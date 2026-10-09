"""
Reporte automático de ventas: lee un Excel y genera un PDF con totales y gráficos.

Uso:
    python generar_reporte.py ventas_ejemplo.xlsx
    python generar_reporte.py ventas.xlsx --negocio "Mi Negocio" --salida reporte.pdf

El Excel necesita estas columnas:
    Fecha, Ticket, Producto, Categoría, Cantidad, Total, Medio de pago
"""

import argparse
import io
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import pandas as pd
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib.utils import ImageReader
from reportlab.pdfgen import canvas

# --- Estilo -----------------------------------------------------------------
AZUL = "#2140B8"
AZUL_CLARO = "#C9D3F5"
TINTA = "#1C2620"
GRIS = "#5B675F"
LINEA = "#D5DBD6"
FONDO = "#F3F5F2"

MESES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio",
         "agosto", "septiembre", "octubre", "noviembre", "diciembre"]
DIAS = ["Lun", "Mar", "Mié", "Jue", "Vie", "Sáb", "Dom"]
COLUMNAS = ["Fecha", "Ticket", "Producto", "Categoría", "Cantidad", "Total", "Medio de pago"]

plt.rcParams.update({
    "font.family": "DejaVu Sans",
    "font.size": 9,
    "axes.edgecolor": LINEA,
    "axes.labelcolor": GRIS,
    "xtick.color": GRIS,
    "ytick.color": GRIS,
    "axes.spines.top": False,
    "axes.spines.right": False,
})


def pesos(valor: float) -> str:
    """Formatea un número como pesos argentinos: $ 1.234.567"""
    return "$ " + f"{valor:,.0f}".replace(",", ".")


# --- Datos ------------------------------------------------------------------
def cargar_ventas(ruta: Path) -> pd.DataFrame:
    df = pd.read_excel(ruta)
    faltan = [c for c in COLUMNAS if c not in df.columns]
    if faltan:
        raise SystemExit(f"Al Excel le faltan estas columnas: {', '.join(faltan)}")
    df["Fecha"] = pd.to_datetime(df["Fecha"])
    return df


def calcular_resumen(df: pd.DataFrame) -> dict:
    total = df["Total"].sum()
    tickets = df["Ticket"].nunique()
    por_dia = df.groupby(df["Fecha"].dt.date)["Total"].sum()
    productos = (df.groupby("Producto")
                   .agg(Unidades=("Cantidad", "sum"), Ventas=("Total", "sum"))
                   .sort_values("Ventas", ascending=False))
    productos["% del total"] = productos["Ventas"] / total * 100
    semana = (df.assign(dia=df["Fecha"].dt.dayofweek)
                .groupby(["dia", df["Fecha"].dt.date])["Total"].sum()
                .groupby("dia").mean()
                .reindex(range(7), fill_value=0))
    return {
        "total": total,
        "tickets": tickets,
        "ticket_promedio": total / tickets,
        "unidades": int(df["Cantidad"].sum()),
        "por_dia": por_dia,
        "mejor_dia": por_dia.idxmax(),
        "por_categoria": df.groupby("Categoría")["Total"].sum().sort_values(),
        "por_medio": df.groupby("Medio de pago")["Total"].sum().sort_values(),
        "productos": productos,
        "semana": semana,
    }


# --- Gráficos (se dibujan con matplotlib y se pegan en el PDF) ---------------
def a_imagen(fig) -> ImageReader:
    buffer = io.BytesIO()
    fig.savefig(buffer, format="png", dpi=200, bbox_inches="tight")
    plt.close(fig)
    buffer.seek(0)
    return ImageReader(buffer)


def grafico_por_dia(por_dia: pd.Series) -> ImageReader:
    fig, ax = plt.subplots(figsize=(7.4, 2.4))
    fechas = pd.to_datetime(por_dia.index)
    color = [AZUL if f.dayofweek >= 5 else AZUL_CLARO for f in fechas]
    ax.bar(fechas.day, por_dia.values / 1000, color=color, width=0.75)
    ax.set_xlabel("Día del mes")
    ax.set_ylabel("Miles de $")
    ax.set_xticks(fechas.day[::2])
    ax.grid(axis="y", color=LINEA, linewidth=0.6)
    ax.set_axisbelow(True)
    return a_imagen(fig)


def grafico_barras_horizontales(serie: pd.Series, total: float) -> ImageReader:
    fig, ax = plt.subplots(figsize=(3.5, 1.9))
    ax.barh(serie.index, serie.values, color=AZUL, height=0.6)
    ax.set_xticks([])
    ax.spines["bottom"].set_visible(False)
    tope = serie.max()
    for i, valor in enumerate(serie.values):
        ax.text(valor + tope * 0.02, i, f"{valor / total * 100:.0f}%",
                va="center", color=TINTA, fontsize=9)
    ax.set_xlim(0, tope * 1.2)
    return a_imagen(fig)


def grafico_semana(semana: pd.Series) -> ImageReader:
    fig, ax = plt.subplots(figsize=(3.5, 2.2))
    color = [AZUL if i >= 5 else AZUL_CLARO for i in range(7)]
    ax.bar(DIAS, semana.values / 1000, color=color, width=0.65)
    ax.set_ylabel("Miles de $ por día")
    ax.grid(axis="y", color=LINEA, linewidth=0.6)
    ax.set_axisbelow(True)
    return a_imagen(fig)


# --- PDF --------------------------------------------------------------------
class Reporte:
    def __init__(self, salida: Path):
        self.c = canvas.Canvas(str(salida), pagesize=A4)
        self.ancho, self.alto = A4
        self.margen = 16 * mm

    def encabezado(self, negocio: str, periodo: str, pagina: int):
        c = self.c
        c.setFillColor(colors.HexColor(AZUL))
        c.rect(0, self.alto - 30 * mm, self.ancho, 30 * mm, stroke=0, fill=1)
        c.setFillColor(colors.white)
        c.setFont("Helvetica", 10)
        c.drawString(self.margen, self.alto - 12 * mm, negocio.upper())
        c.setFont("Helvetica-Bold", 20)
        c.drawString(self.margen, self.alto - 22 * mm, f"Reporte de ventas · {periodo}")
        c.setFont("Helvetica", 9)
        c.drawRightString(self.ancho - self.margen, self.alto - 12 * mm, f"Página {pagina} de 2")

    def titulo(self, texto: str, y: float):
        self.c.setFillColor(colors.HexColor(TINTA))
        self.c.setFont("Helvetica-Bold", 12)
        self.c.drawString(self.margen, y, texto)

    def tarjetas(self, datos: list, y: float):
        c = self.c
        espacio = 4 * mm
        ancho = (self.ancho - 2 * self.margen - 3 * espacio) / 4
        for i, (etiqueta, valor) in enumerate(datos):
            x = self.margen + i * (ancho + espacio)
            c.setFillColor(colors.HexColor(FONDO))
            c.roundRect(x, y, ancho, 22 * mm, 3 * mm, stroke=0, fill=1)
            c.setFillColor(colors.HexColor(GRIS))
            c.setFont("Helvetica", 8.5)
            c.drawString(x + 4 * mm, y + 15 * mm, etiqueta)
            c.setFillColor(colors.HexColor(TINTA))
            c.setFont("Helvetica-Bold", 15)
            c.drawString(x + 4 * mm, y + 6 * mm, valor)

    def imagen(self, img: ImageReader, x: float, y: float, ancho: float):
        w, h = img.getSize()
        alto = ancho * h / w
        self.c.drawImage(img, x, y - alto, ancho, alto)
        return y - alto

    def tabla_productos(self, productos: pd.DataFrame, y: float):
        c = self.c
        columnas = [("Producto", self.margen, "izq"),
                    ("Unidades", self.ancho - self.margen - 80 * mm, "der"),
                    ("Ventas", self.ancho - self.margen - 30 * mm, "der"),
                    ("% del total", self.ancho - self.margen, "der")]

        def fila(valores, fuente, color_texto):
            c.setFont(fuente, 9.5)
            c.setFillColor(colors.HexColor(color_texto))
            for (_, x, alineado), valor in zip(columnas, valores):
                (c.drawString if alineado == "izq" else c.drawRightString)(x, y, valor)

        fila([nombre for nombre, _, _ in columnas], "Helvetica-Bold", GRIS)
        y -= 3 * mm
        for nombre, datos in productos.head(10).iterrows():
            c.setStrokeColor(colors.HexColor(LINEA))
            c.line(self.margen, y, self.ancho - self.margen, y)
            y -= 5.5 * mm
            fila([nombre, f"{int(datos['Unidades'])}", pesos(datos["Ventas"]),
                  f"{datos['% del total']:.1f}%".replace(".", ",")], "Helvetica", TINTA)
            y -= 2.5 * mm
        return y

    def pie(self):
        self.c.setFillColor(colors.HexColor(GRIS))
        self.c.setFont("Helvetica", 8)
        self.c.drawString(self.margen, 10 * mm, "Generado automáticamente con Python")


def generar(ruta_excel: Path, negocio: str, salida: Path):
    df = cargar_ventas(ruta_excel)
    r = calcular_resumen(df)
    inicio = df["Fecha"].min()
    periodo = f"{MESES[inicio.month - 1].capitalize()} {inicio.year}"
    mejor = pd.Timestamp(r["mejor_dia"])

    pdf = Reporte(salida)
    m, ancho_util = pdf.margen, pdf.ancho - 2 * pdf.margen

    # Página 1: resumen
    pdf.encabezado(negocio, periodo, 1)
    y = pdf.alto - 58 * mm
    pdf.tarjetas([("Ventas totales", pesos(r["total"])),
                  ("Tickets", f"{r['tickets']:,}".replace(",", ".")),
                  ("Ticket promedio", pesos(r["ticket_promedio"])),
                  ("Unidades vendidas", f"{r['unidades']:,}".replace(",", "."))], y)
    y -= 12 * mm
    pdf.titulo("Ventas por día", y)
    pdf.c.setFont("Helvetica", 9)
    pdf.c.setFillColor(colors.HexColor(GRIS))
    pdf.c.drawString(m, y - 5 * mm,
                     f"Los fines de semana, en azul fuerte. Mejor día: {DIAS[mejor.dayofweek]} {mejor.day} "
                     f"con {pesos(r['por_dia'].max())}.")
    y = pdf.imagen(grafico_por_dia(r["por_dia"]), m, y - 8 * mm, ancho_util)
    y -= 10 * mm
    mitad = (ancho_util - 8 * mm) / 2
    pdf.titulo("Ventas por categoría", y)
    pdf.c.setFont("Helvetica-Bold", 12)
    pdf.c.drawString(m + mitad + 8 * mm, y, "Medios de pago")
    pdf.imagen(grafico_barras_horizontales(r["por_categoria"], r["total"]), m, y - 4 * mm, mitad)
    pdf.imagen(grafico_barras_horizontales(r["por_medio"], r["total"]), m + mitad + 8 * mm, y - 4 * mm, mitad)
    pdf.pie()
    pdf.c.showPage()

    # Página 2: productos y días de la semana
    pdf.encabezado(negocio, periodo, 2)
    y = pdf.alto - 44 * mm
    pdf.titulo("Los 10 productos que más facturaron", y)
    y = pdf.tabla_productos(r["productos"], y - 9 * mm)
    y -= 8 * mm
    pdf.titulo("Venta promedio según el día de la semana", y)
    pdf.imagen(grafico_semana(r["semana"]), m, y - 4 * mm, ancho_util * 0.62)
    pdf.pie()
    pdf.c.save()
    print(f"Listo: {salida} ({len(df)} ventas de {periodo})")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Genera un reporte PDF a partir de un Excel de ventas.")
    parser.add_argument("excel", type=Path, help="Archivo Excel con las ventas")
    parser.add_argument("--negocio", default="Brote Café de Barrio", help="Nombre del negocio")
    parser.add_argument("--salida", type=Path, default=Path("reporte_ventas.pdf"), help="PDF de salida")
    args = parser.parse_args()
    generar(args.excel, args.negocio, args.salida)
