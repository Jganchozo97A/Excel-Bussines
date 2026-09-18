# Excel-Bussines

Repositorio para el desarrollo, automatización y comercialización de plantillas Excel profesionales para negocios, finanzas y productividad.

## Línea de productos digitales

| # | Producto | Precio orientativo (USD) | Estado |
|---|----------|---------------------------|--------|
| 1 | Control de gastos personales | $6.99 | ✅ Disponible |
| 2 | Presupuesto mensual | $6.99 | ✅ Disponible |
| 3 | Control de deudas | $7.99 | ⏳ Pendiente |
| 4 | Control de inventario | $9.99 | ⏳ Pendiente |
| 5 | Flujo de caja para negocios | $12.99 | ⏳ Pendiente |
| 6 | Contabilidad básica para emprendedores | $14.99 | ⏳ Pendiente |
| 7 | Control de ventas | $9.99 | ⏳ Pendiente |
| 8 | Nómina básica | $12.99 | ⏳ Pendiente |
| 9 | Pack 5 plantillas | $24.99 | ⏳ Pendiente |
| 10 | Pack 10 plantillas | $39.99 | ⏳ Pendiente |

Cada plantilla incluye:

- ✅ Dashboard con indicadores clave y gráficos
- ✅ Fórmulas automáticas
- ✅ Gráficos (pastel, barras, tendencia)
- ✅ Resumen dinámico tipo tabla dinámica cuando corresponde
- ✅ Listas desplegables y validación de datos
- ✅ Hoja de instrucciones
- ✅ Protección de celdas y fórmulas (solo editable donde corresponde)
- ✅ Diseño profesional
- ✅ Ejemplos precargados
- ✅ Selector de moneda integrado (local / USD, con tipo de cambio editable)

## Estructura del repositorio

```
productos/
  01-control-gastos-personales/
    Control-de-Gastos-Personales.xlsx
  02-presupuesto-mensual/
    Presupuesto-Mensual.xlsx
scripts/
  build_control_gastos_personales.py   # genera el .xlsx del producto 1 con openpyxl
  build_presupuesto_mensual.py         # genera el .xlsx del producto 2 con openpyxl
```

Todas las plantillas comparten la misma marca (paleta de colores, tipografía Arial, estructura
de hojas Portada/Instrucciones/Configuración) y el mismo mecanismo de selector de moneda
(Local/USD) y de contraseña de desprotección: `plantilla2026`.

## Producto 1: Control de gastos personales

Hojas: Portada · Instrucciones · Configuración · Categorías · Transacciones · Resumen · Dashboard.

- El selector de moneda vive en la hoja **Configuración** (símbolo local, tipo de cambio y
  "Moneda de visualización": Local/USD). Todas las hojas leen ese factor de conversión, así
  que cambiar la moneda recalcula Transacciones, Resumen y Dashboard sin duplicar hojas.
- El "Resumen" usa fórmulas `SUMIFS` (por categoría y por mes) en lugar de una tabla dinámica
  nativa de Excel, para evitar la fragilidad de generar tablas dinámicas por XML: el resultado
  se comporta igual (se actualiza solo) pero es mucho más robusto.
- Para regenerar el archivo tras editar el script: `python3 scripts/build_control_gastos_personales.py`.

## Producto 2: Presupuesto mensual

Hojas: Portada · Instrucciones · Configuración · Categorías · Presupuesto · Real · Resumen · Dashboard.

- "Presupuesto" y "Real" son grillas categoría × mes (Ene-Dic) donde el usuario escribe cuánto
  planea y cuánto efectivamente ingresó/gastó cada mes.
- "Configuración" añade un selector de **mes** (además del de moneda), y "Resumen" usa
  `INDEX`/`MATCH` para traer, para el mes elegido, el presupuestado y el real de cada categoría.
- El Dashboard separa Ingresos de Gastos (en vez de un total mezclado) para que "% de gastos
  usado" y "Balance" tengan sentido financiero real.
- Para regenerar el archivo: `python3 scripts/build_presupuesto_mensual.py`.
