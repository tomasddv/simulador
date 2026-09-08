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
