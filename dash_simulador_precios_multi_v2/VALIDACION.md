# Validación contra ERP

La versión anterior aplicaba el descuento sobre `Final actual / bulto`, que incluía una regla adicional del 3% detectada en un archivo de ejemplo. Esa regla no coincide con el cálculo mostrado por el ERP en los dos SKU controlados.

## Lógica ERP usada ahora

Para cada SKU se toman de la lista de precios:

- `Precio Base / bulto` = columna Precio del ERP.
- `P. Final / bulto` = importe final sin bonificación.
- `Unidades / bulto`.

Con un descuento `d`:

```text
Bonif. ERP / bulto = Precio Base / bulto × d
Final ERP / bulto  = P. Final / bulto × (1 - d)
Final / unidad     = Final ERP / bulto ÷ unidades por bulto
Sugerido           = Final / unidad × 1,30
```

Para calcular desde un sugerido objetivo:

```text
Final / unidad objetivo = Sugerido ÷ 1,30
Final ERP / bulto objetivo = Final / unidad objetivo × unidades por bulto
Descuento = 1 - Final ERP / bulto objetivo ÷ P. Final / bulto
Bonif. ERP / bulto = Precio Base / bulto × descuento
```

## Control SKU 2218 — QUILMES 4X6 LAT 473

Datos de lista:

- Precio Base: 35.556,580
- P. Final: 47.339,222
- Unidades/bulto: 24
- Descuento: 5%

Resultado:

- Bonif. ERP: 1.777,829 — coincide con ERP.
- Final calculado: 44.972,261.
- Final mostrado en ERP: 44.972,259.
- Diferencia: 0,002 por bulto.

## Control SKU 30793 — PEPSI BLACK CAN 4X6 354CC 2024

Datos de lista:

- Precio Base: 29.778,450
- P. Final: 38.072,245
- Unidades/bulto: 24
- Descuento: 14%

Resultado:

- Bonif. ERP: 4.168,983 — coincide con ERP.
- Final calculado: 32.742,131.
- Final mostrado en ERP: 32.742,136.
- Diferencia: 0,005 por bulto.

Las diferencias residuales de milésimas indican que el ERP conserva más decimales internos que los expuestos en la lista de precios. Con los valores disponibles, esta es la reproducción más cercana y consistente sin hardcodear SKU individuales.
