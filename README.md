# Reporte automático de ventas (Python)

Script que toma un Excel con las ventas de un negocio y genera en segundos un **reporte en PDF** con totales, gráficos y ranking de productos. Pensado para dueños de comercios que hoy arman estos números a mano.

Proyecto demo de portfolio, con datos ficticios de "Brote Café de Barrio".

## Qué incluye el reporte
- Ventas totales, cantidad de tickets, ticket promedio y unidades vendidas
- Gráfico de ventas por día, con los fines de semana destacados y el mejor día del mes
- Ventas por categoría y por medio de pago
- Los 10 productos que más facturaron
- Venta promedio según el día de la semana

Ejemplo del resultado: [reporte_ventas.pdf](reporte_ventas.pdf)

## Cómo usarlo
```bash
pip install -r requirements.txt
python generar_reporte.py ventas_ejemplo.xlsx
```

Con otro negocio y otro nombre de archivo:
```bash
python generar_reporte.py ventas_octubre.xlsx --negocio "Mi Negocio" --salida reporte_octubre.pdf
```

El Excel necesita las columnas: `Fecha`, `Ticket`, `Producto`, `Categoría`, `Cantidad`, `Total`, `Medio de pago`.

Para generar de nuevo los datos de ejemplo: `python crear_datos_ejemplo.py`

## Tecnologías
Python · pandas · matplotlib · ReportLab

## Autor
Nadir Cunico — Desarrollador web y automatizaciones
