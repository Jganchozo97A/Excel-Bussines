# Excel-Bussines

Repositorio para el desarrollo, automatización y comercialización de plantillas Excel profesionales para negocios, finanzas y productividad.

## Línea de productos digitales

| # | Producto | Precio orientativo (USD) | Estado |
|---|----------|---------------------------|--------|
| 1 | Control de gastos personales | $6.99 | ✅ Disponible |
| 2 | Presupuesto mensual | $6.99 | ✅ Disponible |
| 3 | Control de deudas | $7.99 | ✅ Disponible |
| 4 | Control de inventario | $9.99 | ✅ Disponible |
| 5 | Flujo de caja para negocios | $12.99 | ✅ Disponible |
| 6 | Contabilidad básica para emprendedores | $14.99 | ✅ Disponible |
| 7 | Control de ventas | $9.99 | ✅ Disponible |
| 8 | Nómina básica | $12.99 | ✅ Disponible |
| 9 | Pack 5 plantillas | $24.99 | ✅ Disponible |
| 10 | Pack completo de plantillas (comercializado como "Pack 10") | $39.99 | ✅ Disponible (8 plantillas hoy) |

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
  03-control-deudas/
    Control-de-Deudas.xlsx
  04-control-inventario/
    Control-de-Inventario.xlsx
  05-flujo-de-caja/
    Flujo-de-Caja-para-Negocios.xlsx
  06-contabilidad-basica/
    Contabilidad-Basica-para-Emprendedores.xlsx
  07-control-ventas/
    Control-de-Ventas.xlsx
  08-nomina-basica/
    Nomina-Basica.xlsx
  09-pack-5-plantillas/
    Pack-5-Plantillas.zip
  10-pack-completo/
    Pack-Completo-Plantillas.zip
scripts/
  build_control_gastos_personales.py   # genera el .xlsx del producto 1 con openpyxl
  build_presupuesto_mensual.py         # genera el .xlsx del producto 2 con openpyxl
  build_control_deudas.py              # genera el .xlsx del producto 3 con openpyxl
  build_control_inventario.py          # genera el .xlsx del producto 4 con openpyxl
  build_flujo_caja.py                  # genera el .xlsx del producto 5 con openpyxl
  build_contabilidad_basica.py         # genera el .xlsx del producto 6 con openpyxl
  build_control_ventas.py              # genera el .xlsx del producto 7 con openpyxl
  build_nomina_basica.py               # genera el .xlsx del producto 8 con openpyxl
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

## Producto 3: Control de deudas

Hojas: Portada · Instrucciones · Configuración · Deudas · Pagos · Resumen · Dashboard.

- "Deudas" es el catálogo de deudas (hasta 10): saldo inicial, tasa de interés anual y pago
  mínimo. "Pagos" es la bitácora de pagos (fecha, deuda, monto, interés); el capital amortizado
  se calcula solo y descuenta el saldo de cada deuda vía `SUMIFS`.
- "Resumen" sugiere un orden de pago con dos estrategias clásicas, calculadas con `RANK`:
  **bola de nieve** (paga primero la deuda con menor saldo) y **avalancha** (paga primero la de
  mayor tasa de interés).
- El Dashboard incluye la evolución mensual de la deuda total (saldo acumulado descontando los
  pagos registrados mes a mes).
- Para regenerar el archivo: `python3 scripts/build_control_deudas.py`.

## Producto 4: Control de inventario

Hojas: Portada · Instrucciones · Configuración · Categorías · Productos · Movimientos · Resumen · Dashboard.

- "Productos" es el catálogo (hasta 12 productos: código, categoría, unidad, stock mínimo, costo,
  precio de venta, stock inicial). "Movimientos" registra entradas/salidas por código de producto.
- "Resumen" calcula el stock actual (`stock inicial + entradas - salidas`), el valor de inventario
  y marca "Bajo stock" cuando el stock actual cae por debajo del mínimo.
- El Dashboard alerta cuántos productos están bajo stock mínimo y grafica el valor de inventario
  por categoría y por producto.
- Nota de diseño: el número de filas de Productos/Deudas se mantiene moderado (10-12) a propósito,
  porque los gráficos de pastel/barras en Excel no ocultan automáticamente las categorías vacías;
  con rangos muy largos la leyenda se llena de entradas en blanco.
- Para regenerar el archivo: `python3 scripts/build_control_inventario.py`.

## Producto 5: Flujo de caja para negocios

Hojas: Portada · Instrucciones · Configuración · Categorías · Flujo de Caja · Resumen · Dashboard.

- "Flujo de Caja" es una grilla categoría × mes (como en el producto 2) pero de movimientos reales
  de efectivo del negocio (sin comparación presupuesto/real).
- "Resumen" arrastra el saldo de caja mes a mes: `Saldo final(mes) = Saldo inicial(mes) + Flujo
  neto(mes)`, y `Saldo inicial(mes N) = Saldo final(mes N-1)`, partiendo del saldo de caja inicial
  configurado.
- El Dashboard muestra la curva de saldo de caja proyectado durante el año, además de ingresos vs.
  egresos por mes y egresos por categoría.
- Para regenerar el archivo: `python3 scripts/build_flujo_caja.py`.

## Producto 6: Contabilidad básica para emprendedores

Hojas: Portada · Instrucciones · Configuración · Plan de Cuentas · Transacciones · Estado de
Resultados · Dashboard.

- "Plan de Cuentas" clasifica cada cuenta contable en uno de 6 tipos (Ingreso, Costo de venta,
  Gasto operativo, Gasto administrativo, Gasto financiero, Impuesto). "Transacciones" registra
  cada movimiento por cuenta; el tipo se resuelve solo con `INDEX`/`MATCH`.
- "Estado de Resultados" arma automáticamente el clásico estado de resultados en cascada (Ventas
  − Costo de venta = Utilidad Bruta; − Gastos operativos/administrativos = Utilidad Operativa;
  − Gastos financieros/Impuestos = Utilidad Neta), mes a mes y en total anual, con márgenes.
- Es contabilidad de ingresos y gastos (base para el Estado de Resultados), no partida doble ni
  balance general; la hoja de instrucciones lo aclara explícitamente.
- Para regenerar el archivo: `python3 scripts/build_contabilidad_basica.py`.

## Producto 7: Control de ventas

Hojas: Portada · Instrucciones · Configuración · Catálogo · Clientes · Ventas · Resumen · Dashboard.

- "Catálogo" (precios) y "Clientes" alimentan las listas desplegables de "Ventas"; el precio
  unitario y el total de cada venta se completan solos vía `INDEX`/`MATCH`.
- "Resumen" agrupa ventas por producto, por cliente y por mes, y calcula el ticket promedio, lo
  pendiente de cobro (`Estado = Pendiente`) y el producto más vendido por monto.
- Para regenerar el archivo: `python3 scripts/build_control_ventas.py`.

## Producto 8: Nómina básica

Hojas: Portada · Instrucciones · Configuración · Empleados · Bonificaciones · Resumen · Dashboard.

- "Empleados" es la ficha de personal (hasta 10 colaboradores, con sueldo base). "Bonificaciones"
  es una grilla colaborador × mes para bonos, comisiones u horas extra.
- Las tasas de descuento (seguro social e impuesto/retención) son porcentajes simples editables en
  Configuración — una simplificación deliberada para que sirva en cualquier país; la hoja de
  instrucciones aclara que no reemplaza el cálculo legal de planillas.
- "Resumen" calcula, para el mes elegido (selector en Configuración, igual que en el producto 2),
  el total de ingresos, los descuentos y el sueldo neto de cada colaborador.
- Para regenerar el archivo: `python3 scripts/build_nomina_basica.py`.

## Producto 9: Pack 5 plantillas

`productos/09-pack-5-plantillas/Pack-5-Plantillas.zip` — bundle de 5 plantillas con descuento
frente a comprarlas por separado: Control de Gastos Personales, Presupuesto Mensual, Control de
Deudas, Flujo de Caja para Negocios y Control de Ventas. Incluye un `LEEME.md` con el índice del
pack.

## Producto 10: Pack completo de plantillas ("Pack 10")

`productos/10-pack-completo/Pack-Completo-Plantillas.zip` — bundle con las 8 plantillas de la
línea. Se comercializa como "Pack 10" (el nombre original de la línea de productos), aunque hoy
incluye 8: el `LEEME.md` del pack lo indica explícitamente y aclara que las próximas plantillas
que se agreguen a la línea se sumarán a este pack sin costo adicional para quienes ya lo compraron.
