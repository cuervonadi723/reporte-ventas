"""
Genera un Excel de ventas de ejemplo para probar el reporte.

Los datos son ficticios: un mes de ventas de "Brote Café de Barrio".
Uso:
    python crear_datos_ejemplo.py
"""

import random
from datetime import date, timedelta

import pandas as pd

random.seed(42)

# Producto: (categoría, precio en ARS, popularidad relativa)
PRODUCTOS = {
    "Espresso": ("Café", 2800, 9),
    "Cortado": ("Café", 3200, 12),
    "Flat white": ("Café", 3900, 10),
    "Filtrado V60": ("Café", 4200, 4),
    "Cold brew": ("Café", 4300, 5),
    "Latte de avena": ("Café", 4400, 6),
    "Medialuna de manteca": ("Pastelería", 1300, 14),
    "Budín de limón": ("Pastelería", 2600, 5),
    "Alfajor de maicena": ("Pastelería", 2200, 6),
    "Cookie de chocolate": ("Pastelería", 2400, 5),
    "Tostón de palta": ("Brunch", 7800, 4),
    "Bowl de yogur y granola": ("Brunch", 6200, 3),
    "Revuelto con hongos": ("Brunch", 7400, 2),
    "Combo Brote": ("Brunch", 9900, 4),
}
MEDIOS_DE_PAGO = ["Mercado Pago", "Débito", "Efectivo", "Crédito"]
PESO_MEDIOS = [45, 25, 20, 10]

inicio = date(2026, 9, 1)
filas = []
ticket = 1000

for dia in range(30):
    fecha = inicio + timedelta(days=dia)
    # Más movimiento los fines de semana
    tickets_del_dia = random.randint(45, 70) if fecha.weekday() < 5 else random.randint(70, 95)
    for _ in range(tickets_del_dia):
        ticket += 1
        medio = random.choices(MEDIOS_DE_PAGO, PESO_MEDIOS)[0]
        nombres = random.choices(list(PRODUCTOS), [p[2] for p in PRODUCTOS.values()], k=random.randint(1, 3))
        for nombre in set(nombres):
            categoria, precio, _ = PRODUCTOS[nombre]
            cantidad = random.choices([1, 2, 3], [75, 20, 5])[0]
            filas.append({
                "Fecha": fecha,
                "Ticket": ticket,
                "Producto": nombre,
                "Categoría": categoria,
                "Cantidad": cantidad,
                "Precio unitario": precio,
                "Total": precio * cantidad,
                "Medio de pago": medio,
            })

df = pd.DataFrame(filas)
df.to_excel("ventas_ejemplo.xlsx", index=False, sheet_name="Ventas")
print(f"Listo: ventas_ejemplo.xlsx con {len(df)} filas y {df['Ticket'].nunique()} tickets.")
