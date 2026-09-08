# SIMULADOR MULTI-SKU · ERP

**Esta versión corre en `http://127.0.0.1:8051` para no confundirse con versiones anteriores que usaban el puerto 8050.**

Ejecutar `INICIAR_MULTI_SKU.bat`.

# Simulador de precios por SKU — multi-SKU + Excel

Dashboard en Dash/Plotly para calcular descuento y precio sugerido por SKU con la lógica del ERP, sumar varios SKU a una misma simulación y exportar el resultado a Excel junto con stock y vencimientos de Trelew/Madryn.

## Flujo de uso

1. Ingresar SKU.
2. Elegir cálculo desde `% de descuento` o desde `precio sugerido objetivo`.
3. Definir bultos a bonificar y, opcionalmente, sucursal de frescura.
4. Presionar **＋ Sumar SKU**.
5. Repetir con todos los SKU necesarios.
6. Revisar la alerta de vencimientos del mes corriente.
7. Presionar **Exportar Excel**.

Si se vuelve a sumar el mismo `SKU + sucursal`, el escenario se actualiza en lugar de duplicarse.

## Qué incluye el listado

- SKU y descripción.
- Sucursal de frescura usada.
- Bultos a bonificar.
- Descuento.
- Precio sugerido.
- Final por unidad y por bulto.
- Bonificación ERP por bulto y gasto total.
- Stock total informado.
- Stock asociado a cada fecha de vencimiento.
- Alerta **⚠ VENCE ESTE MES** cuando existe stock con vencimiento dentro del mes corriente.

## Excel exportado

El archivo tiene dos hojas:

### `Simulacion`
Una fila por SKU/sucursal agregado. Incluye fórmulas editables para Bonif. ERP, Final ERP, Final/unidad, Sugerido y Gasto total, además del stock y el próximo vencimiento.

### `Vencimientos`
Una fila por lote, con sucursal, vencimiento, estado, stock, orden de lote, días para bloqueo y marca de mes corriente.

## Fórmula ERP

```text
Bonif. ERP = Precio Base × descuento
Final ERP = P. Final × (1 - descuento)
Final por unidad = Final ERP ÷ unidades por bulto
Sugerido = Final por unidad × 1,30
```

## Ejecutar en Windows

Doble clic en `INICIAR_DASH.bat`.

O manualmente:

```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
python app.py
```

Abrir `http://127.0.0.1:8050`.

## Validación ERP

Ver `VALIDACION.md` para el control de los SKU 2218 y 30793 contra el ERP.

## Fuentes de datos en Google Drive

La versión online intenta leer estas fuentes públicas directamente desde Google Drive:

- Lista de precios: `11bArEuSP1JLoHnQRgJSnt6kRY7xeCuv5`
- Frescura Trelew: `1w5ydKY-p8pkOphqeIRGXiTdrfHnVYFnM`
- Frescura Madryn: `18rSCdJrYcQ_8nyWHULXd12oFO8HqkC7p`

El dashboard muestra el estado de las tres fuentes y tiene el botón **Actualizar datos**. Si Drive no responde, usa los CSV incluidos en `data/` como respaldo para que la app no se caiga.

### Importante para la actualización diaria

Para que Render siga leyendo automáticamente los archivos, mantené los mismos archivos/IDs de Drive y reemplazá o actualizá su contenido. Si se crea un archivo nuevo con otro ID, hay que cambiar el ID en `drive_data.py` o definir en Render estas variables de entorno:

- `DRIVE_PRICE_FILE_ID`
- `DRIVE_TRELEW_FILE_ID`
- `DRIVE_MADRYN_FILE_ID`

Así se puede cambiar una fuente sin modificar código.
