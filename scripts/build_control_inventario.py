import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, Protection
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.chart import PieChart, BarChart, LineChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.workbook.defined_name import DefinedName
import datetime
import os

NAVY = "1F3B57"
TEAL = "2CA58D"
GOLD = "C9A24B"
LIGHT_GRAY = "F2F4F6"
MID_GRAY = "D9DEE3"
WHITE = "FFFFFF"
DARK_TEXT = "1A1A1A"
RED_NEG = "B23B3B"
INPUT_FILL = "FFF6D9"
LOCKED_TEXT = "555555"
FONT_NAME = "Arial"

f_h1 = Font(name=FONT_NAME, size=16, bold=True, color=NAVY)
f_h2 = Font(name=FONT_NAME, size=12, bold=True, color=WHITE)
f_h3 = Font(name=FONT_NAME, size=11, bold=True, color=NAVY)
f_body = Font(name=FONT_NAME, size=10, color=DARK_TEXT)
f_body_b = Font(name=FONT_NAME, size=10, bold=True, color=DARK_TEXT)
f_kpi_label = Font(name=FONT_NAME, size=10, bold=True, color=WHITE)
f_kpi_value = Font(name=FONT_NAME, size=16, bold=True, color=WHITE)
f_note = Font(name=FONT_NAME, size=9, italic=True, color=LOCKED_TEXT)
f_input = Font(name=FONT_NAME, size=10, color="0000FF")

fill_header = PatternFill("solid", fgColor=NAVY)
fill_header2 = PatternFill("solid", fgColor=TEAL)
fill_light = PatternFill("solid", fgColor=LIGHT_GRAY)
fill_input = PatternFill("solid", fgColor=INPUT_FILL)
fill_white = PatternFill("solid", fgColor=WHITE)

thin = Side(style="thin", color=MID_GRAY)
border_all = Border(left=thin, right=thin, top=thin, bottom=thin)

align_center = Alignment(horizontal="center", vertical="center")
align_left = Alignment(horizontal="left", vertical="center")
align_wrap = Alignment(horizontal="left", vertical="top", wrap_text=True)

PROTECT_PASSWORD = "plantilla2026"

wb = Workbook()
wb.remove(wb.active)


def style_header_row(ws, row, col_start, col_end, fill=fill_header, font=f_h2, height=22):
    ws.row_dimensions[row].height = height
    for c in range(col_start, col_end + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = fill
        cell.font = font
        cell.alignment = align_center
        cell.border = border_all


def set_col_widths(ws, widths):
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


def lock_all(ws):
    ws.protection.sheet = True
    ws.protection.password = PROTECT_PASSWORD
    ws.protection.formatCells = False
    ws.protection.formatColumns = False
    ws.protection.formatRows = False
    ws.protection.insertRows = False
    ws.protection.selectLockedCells = True
    ws.protection.selectUnlockedCells = True


def unlock_range(ws, cell_range):
    for row in ws[cell_range]:
        for cell in row:
            cell.protection = Protection(locked=False)


def build_portada(title, subtitle, features, price):
    ws = wb.create_sheet("Portada")
    set_col_widths(ws, [4, 22, 22, 22, 22, 22, 4])
    ws.sheet_view.showGridLines = False
    ws.merge_cells("B2:F2")
    ws["B2"] = "LÍNEA DE PRODUCTOS DIGITALES · PLANTILLAS EXCEL"
    ws["B2"].font = Font(name=FONT_NAME, size=10, bold=True, color=TEAL)
    ws.merge_cells("B4:F6")
    ws["B4"] = title
    ws["B4"].font = Font(name=FONT_NAME, size=28, bold=True, color=NAVY)
    ws["B4"].alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    ws.merge_cells("B7:F8")
    ws["B7"] = subtitle
    ws["B7"].font = Font(name=FONT_NAME, size=11, color=DARK_TEXT)
    ws["B7"].alignment = align_wrap
    for r in range(2, 9):
        ws.row_dimensions[r].height = 20
    ws.row_dimensions[4].height = 30
    row0 = 10
    ws.merge_cells(f"B{row0}:F{row0}")
    ws[f"B{row0}"] = "QUÉ INCLUYE ESTA PLANTILLA"
    ws[f"B{row0}"].font = f_h3
    for c in range(2, 7):
        ws.cell(row=row0, column=c).fill = fill_light
    ws.row_dimensions[row0].height = 20
    r = row0 + 1
    for feat in features:
        ws.merge_cells(f"B{r}:F{r}")
        ws[f"B{r}"] = f"✔  {feat}"
        ws[f"B{r}"].font = f_body
        ws.row_dimensions[r].height = 18
        r += 1
    r += 1
    ws.merge_cells(f"B{r}:F{r}")
    ws[f"B{r}"] = f"Producto digital · Precio orientativo: {price}  |  Licencia de uso personal"
    ws[f"B{r}"].font = Font(name=FONT_NAME, size=10, bold=True, color=GOLD)
    r += 2
    ws.merge_cells(f"B{r}:F{r}")
    ws[f"B{r}"] = "© Tu Marca de Plantillas · www.tumarca.com"
    ws[f"B{r}"].font = f_note
    lock_all(ws)


def build_instrucciones(steps):
    ws = wb.create_sheet("Instrucciones")
    set_col_widths(ws, [3, 46, 46, 3])
    ws.sheet_view.showGridLines = False
    ws.merge_cells("B2:C2")
    ws["B2"] = "CÓMO USAR ESTA PLANTILLA"
    ws["B2"].font = f_h1
    ws.row_dimensions[2].height = 26
    r = 4
    for title, body in steps:
        ws.merge_cells(f"B{r}:C{r}")
        ws[f"B{r}"] = title
        ws[f"B{r}"].font = f_h3
        for cc in (2, 3):
            ws.cell(row=r, column=cc).fill = fill_light
        ws.row_dimensions[r].height = 18
        r += 1
        ws.merge_cells(f"B{r}:C{r}")
        ws[f"B{r}"] = body
        ws[f"B{r}"].font = f_body
        ws[f"B{r}"].alignment = align_wrap
        ws.row_dimensions[r].height = 42
        r += 2
    ws.merge_cells(f"B{r}:C{r}")
    ws[f"B{r}"] = "Soporte y otras plantillas de esta línea de productos: www.tumarca.com"
    ws[f"B{r}"].font = f_note
    lock_all(ws)
    return ws


def build_configuracion(extra_rows=None):
    ws = wb.create_sheet("Configuracion")
    set_col_widths(ws, [3, 32, 20, 40, 3])
    ws.sheet_view.showGridLines = False
    ws.merge_cells("B2:D2")
    ws["B2"] = "CONFIGURACIÓN Y MONEDA"
    ws["B2"].font = f_h1
    ws.row_dimensions[2].height = 26
    headers_row = 4
    ws.cell(row=headers_row, column=2, value="Parámetro")
    ws.cell(row=headers_row, column=3, value="Valor")
    ws.cell(row=headers_row, column=4, value="Notas")
    style_header_row(ws, headers_row, 2, 4)

    rows = [
        ("Nombre de moneda local", "Soles (PEN)", "Solo referencia. Escribe el nombre de tu moneda."),
        ("Símbolo de moneda local", "S/", "Símbolo que se usará en el Dashboard cuando la vista esté en LOCAL."),
        ("Tipo de cambio (1 USD =)", 3.75, "Cuántas unidades de tu moneda local equivalen a 1 USD."),
        ("Moneda de visualización", "Local", "Elige LOCAL o USD. Afecta Resumen y Dashboard."),
    ]
    if extra_rows:
        rows.extend(extra_rows)

    row_map = {}
    r = 5
    for label, value, note in rows:
        ws.cell(row=r, column=2, value=label).font = f_body_b
        c = ws.cell(row=r, column=3, value=value)
        c.font = f_input
        c.fill = fill_input
        c.alignment = align_center
        c.border = border_all
        ws.cell(row=r, column=4, value=note).font = f_note
        ws.cell(row=r, column=4).alignment = align_wrap
        ws.row_dimensions[r].height = 18
        ws.cell(row=r, column=2).border = border_all
        row_map[label] = r
        r += 1

    ROW_NOMBRE = row_map["Nombre de moneda local"]
    ROW_SIMBOLO = row_map["Símbolo de moneda local"]
    ROW_TC = row_map["Tipo de cambio (1 USD =)"]
    ROW_VIS = row_map["Moneda de visualización"]

    r_calc_label = r
    ws.cell(row=r_calc_label, column=2, value="Valores calculados (no editar)").font = f_h3
    for cc in (2, 3, 4):
        ws.cell(row=r_calc_label, column=cc).fill = fill_light
    ws.row_dimensions[r_calc_label].height = 18
    r += 1

    ROW_FACTOR = r
    ws.cell(row=ROW_FACTOR, column=2, value="Factor de conversión aplicado").font = f_body_b
    ws.cell(row=ROW_FACTOR, column=3, value=f'=IF(C{ROW_VIS}="USD",1/C{ROW_TC},1)')
    ws.cell(row=ROW_FACTOR, column=3).number_format = "0.0000"
    ws.cell(row=ROW_FACTOR, column=3).border = border_all
    ws.cell(row=ROW_FACTOR, column=4, value="Vale 1 si vista LOCAL; vale 1/Tipo de cambio si vista USD.").font = f_note
    ws.row_dimensions[ROW_FACTOR].height = 18
    r += 1

    ROW_SIMBOLO_ACTUAL = r
    ws.cell(row=ROW_SIMBOLO_ACTUAL, column=2, value="Símbolo de moneda actual").font = f_body_b
    ws.cell(row=ROW_SIMBOLO_ACTUAL, column=3, value=f'=IF(C{ROW_VIS}="USD","US$",C{ROW_SIMBOLO})')
    ws.cell(row=ROW_SIMBOLO_ACTUAL, column=3).border = border_all
    ws.cell(row=ROW_SIMBOLO_ACTUAL, column=3).alignment = align_center
    ws.cell(row=ROW_SIMBOLO_ACTUAL, column=4, value="Símbolo que verás en el Dashboard según la moneda elegida.").font = f_note
    ws.row_dimensions[ROW_SIMBOLO_ACTUAL].height = 18
    r += 1

    note_row = r + 1
    ws.merge_cells(f"B{note_row}:D{note_row}")
    ws[f"B{note_row}"] = "Cambia 'Moneda de visualización' entre Local y USD para recalcular automáticamente toda la plantilla."
    ws[f"B{note_row}"].font = f_note
    ws[f"B{note_row}"].alignment = align_wrap

    lock_all(ws)
    for label in row_map:
        rr = row_map[label]
        unlock_range(ws, f"C{rr}:C{rr}")

    dv_moneda = DataValidation(type="list", formula1='"Local,USD"', allow_blank=False)
    ws.add_data_validation(dv_moneda)
    dv_moneda.add(ws[f"C{ROW_VIS}"])

    return ws, dict(ROW_NOMBRE=ROW_NOMBRE, ROW_SIMBOLO=ROW_SIMBOLO, ROW_TC=ROW_TC, ROW_VIS=ROW_VIS,
                     ROW_FACTOR=ROW_FACTOR, ROW_SIMBOLO_ACTUAL=ROW_SIMBOLO_ACTUAL, row_map=row_map)


# =========================================================
build_portada(
    "CONTROL DE INVENTARIO",
    ("Registra tus productos, controla entradas y salidas de stock, y detecta a tiempo qué "
     "productos necesitas reponer — con panel de control, gráficos automáticos y selector de "
     "moneda (local / USD)."),
    [
        "Dashboard con indicadores clave (KPIs) y gráficos automáticos",
        "Fórmulas automáticas: nada que calcular a mano",
        "Gráficos: pastel (valor por categoría) y barras (stock por producto)",
        "Resumen dinámico tipo tabla dinámica por producto",
        "Alertas automáticas de stock bajo el mínimo",
        "Listas desplegables y validación de datos en cada registro",
        "Selector de moneda: local o USD, con tipo de cambio editable",
        "Hoja de instrucciones paso a paso",
        "Celdas y fórmulas protegidas: solo editas donde debes",
        "Diseño profesional y ejemplos precargados",
    ],
    "$9.99 USD",
)

build_instrucciones([
    ("1. Configuración",
     "Define el símbolo de tu moneda local, el tipo de cambio y si la plantilla se muestra en "
     "moneda LOCAL o USD."),
    ("2. Categorías",
     "En la hoja 'Categorías' revisa o edita las categorías de producto y las unidades de medida "
     "que aparecen como listas desplegables."),
    ("3. Productos",
     "En la hoja 'Productos' registra cada artículo: código, nombre, categoría, unidad, stock "
     "mínimo, costo unitario, precio de venta y stock inicial. Hay espacio para hasta 12 productos."),
    ("4. Movimientos",
     "Cada vez que entre o salga mercadería, regístralo en 'Movimientos': fecha, producto (lista "
     "desplegable), tipo (Entrada/Salida), cantidad y referencia."),
    ("5. Resumen automático",
     "La hoja 'Resumen' calcula el stock actual y el valor de inventario de cada producto, y marca "
     "en rojo los que están por debajo de su stock mínimo."),
    ("6. Dashboard",
     "La hoja 'Dashboard' muestra tu valor total de inventario, unidades totales y productos con "
     "stock bajo, con gráficos que se actualizan solos."),
    ("7. Cambiar de moneda",
     "Cambia 'Moneda de visualización' en 'Configuración' entre Local y USD: todo se recalcula "
     "automáticamente."),
    ("8. Celdas protegidas",
     "Las hojas están protegidas para que no borres fórmulas por accidente. Solo puedes escribir en "
     "las celdas resaltadas en color crema / texto azul. Contraseña de desprotección: plantilla2026."),
])

ws_cfg, cfg = build_configuracion()
ROW_FACTOR = cfg["ROW_FACTOR"]
ROW_SIMBOLO_ACTUAL = cfg["ROW_SIMBOLO_ACTUAL"]
ROW_VIS = cfg["ROW_VIS"]

# =========================================================
# 4. CATEGORIAS
# =========================================================
ws = wb.create_sheet("Categorias")
set_col_widths(ws, [3, 28, 20, 3])
ws.sheet_view.showGridLines = False
ws.merge_cells("B2:C2")
ws["B2"] = "CATEGORÍAS Y UNIDADES"
ws["B2"].font = f_h1
ws.row_dimensions[2].height = 26

ws.cell(row=4, column=2, value="Categoría de producto")
ws.cell(row=4, column=3, value="Unidad de medida")
style_header_row(ws, 4, 2, 3)

categorias_prod = ["Electrónica", "Alimentos y bebidas", "Ropa y calzado", "Hogar y limpieza",
                   "Papelería y oficina", "Otros"]
unidades = ["Unidad", "Caja", "Kg", "Litro", "Paquete", "Docena"]

CAT_FIRST = 5
CAT_LAST = CAT_FIRST + len(categorias_prod) - 1
for i, cat in enumerate(categorias_prod):
    rr = CAT_FIRST + i
    c = ws.cell(row=rr, column=2, value=cat)
    c.font = f_input
    c.fill = fill_input
    c.border = border_all

UNI_FIRST = 5
UNI_LAST = UNI_FIRST + len(unidades) - 1
for i, u in enumerate(unidades):
    rr = UNI_FIRST + i
    c = ws.cell(row=rr, column=3, value=u)
    c.font = f_input
    c.fill = fill_input
    c.border = border_all

note_row = max(CAT_LAST, UNI_LAST) + 2
ws.merge_cells(f"B{note_row}:C{note_row}")
ws[f"B{note_row}"] = "Puedes renombrar o agregar categorías/unidades; las listas de 'Productos' se actualizan solas."
ws[f"B{note_row}"].font = f_note
ws[f"B{note_row}"].alignment = align_wrap

lock_all(ws)
unlock_range(ws, f"B{CAT_FIRST}:B{CAT_LAST+6}")
unlock_range(ws, f"C{UNI_FIRST}:C{UNI_LAST+6}")

CAT_RANGE = f"Categorias!$B${CAT_FIRST}:$B${CAT_LAST}"
UNI_RANGE = f"Categorias!$C${UNI_FIRST}:$C${UNI_LAST}"

# =========================================================
# 5. PRODUCTOS
# =========================================================
ws = wb.create_sheet("Productos")
set_col_widths(ws, [12, 26, 18, 12, 12, 14, 14, 12])
ws.sheet_view.showGridLines = False
ws.freeze_panes = "A5"

ws.merge_cells("A2:H2")
ws["A2"] = "CATÁLOGO DE PRODUCTOS"
ws["A2"].font = f_h1
ws.row_dimensions[2].height = 26
ws.merge_cells("A3:H3")
ws["A3"] = "Completa las columnas A a H para cada producto (hasta 12 productos)."
ws["A3"].font = f_note

P_HEADER = 4
p_headers = ["Código", "Nombre", "Categoría", "Unidad", "Stock mínimo", "Costo unitario",
             "Precio de venta", "Stock inicial"]
for i, h in enumerate(p_headers, start=1):
    ws.cell(row=P_HEADER, column=i, value=h)
style_header_row(ws, P_HEADER, 1, 8, height=28)

P_FIRST = P_HEADER + 1
P_LAST = P_FIRST + 11  # 12 products (keeps chart legends/labels readable)

example_products = [
    ("ELE-001", "Audífonos Bluetooth", "Electrónica", "Unidad", 10, 15, 29.9, 40),
    ("ELE-002", "Cargador USB-C", "Electrónica", "Unidad", 15, 6, 14.9, 60),
    ("ALI-001", "Café molido 500g", "Alimentos y bebidas", "Paquete", 20, 8, 15.9, 80),
    ("ALI-002", "Aceite vegetal 1L", "Alimentos y bebidas", "Litro", 15, 5, 9.9, 12),
    ("ROP-001", "Camiseta básica", "Ropa y calzado", "Unidad", 10, 12, 24.9, 25),
    ("ROP-002", "Zapatillas urbanas", "Ropa y calzado", "Unidad", 5, 35, 69.9, 8),
    ("HOG-001", "Detergente 1L", "Hogar y limpieza", "Litro", 12, 4, 8.9, 30),
    ("PAP-001", "Cuaderno A4", "Papelería y oficina", "Unidad", 20, 2, 4.9, 50),
]

for i in range(P_LAST - P_FIRST + 1):
    rr = P_FIRST + i
    if i < len(example_products):
        vals = example_products[i]
        for col, v in enumerate(vals, start=1):
            ws.cell(row=rr, column=col, value=v)
    for col in range(1, 9):
        cell = ws.cell(row=rr, column=col)
        cell.font = f_input
        cell.fill = fill_input
        cell.border = border_all
        if col in (5, 8):
            cell.number_format = "#,##0"
            cell.alignment = align_center
        elif col in (6, 7):
            cell.number_format = "#,##0.00"
            cell.alignment = align_center
        else:
            cell.alignment = align_left

dv_cat_p = DataValidation(type="list", formula1=f"={CAT_RANGE}", allow_blank=True)
ws.add_data_validation(dv_cat_p)
dv_cat_p.add(f"C{P_FIRST}:C{P_LAST}")

dv_uni_p = DataValidation(type="list", formula1=f"={UNI_RANGE}", allow_blank=True)
ws.add_data_validation(dv_uni_p)
dv_uni_p.add(f"D{P_FIRST}:D{P_LAST}")

dv_num_p = DataValidation(type="decimal", operator="greaterThanOrEqual", formula1="0",
                           allow_blank=True, showErrorMessage=True,
                           errorTitle="Valor inválido", error="Ingresa un número igual o mayor a 0.")
ws.add_data_validation(dv_num_p)
dv_num_p.add(f"E{P_FIRST}:H{P_LAST}")

lock_all(ws)
unlock_range(ws, f"A{P_FIRST}:H{P_LAST}")

PROD_CODE_RANGE = f"Productos!$A${P_FIRST}:$A${P_LAST}"

# =========================================================
# 6. MOVIMIENTOS
# =========================================================
ws = wb.create_sheet("Movimientos")
set_col_widths(ws, [14, 14, 30, 14, 12, 26])
ws.sheet_view.showGridLines = False
ws.freeze_panes = "A5"

ws.merge_cells("A2:F2")
ws["A2"] = "MOVIMIENTOS DE INVENTARIO"
ws["A2"].font = f_h1
ws.row_dimensions[2].height = 26
ws.merge_cells("A3:F3")
ws["A3"] = "Completa las columnas A a E para cada movimiento (entrada o salida de stock)."
ws["A3"].font = f_note

M_HEADER = 4
m_headers = ["Fecha", "Código producto", "Producto", "Tipo", "Cantidad", "Referencia / Motivo"]
for i, h in enumerate(m_headers, start=1):
    ws.cell(row=M_HEADER, column=i, value=h)
style_header_row(ws, M_HEADER, 1, 6, height=22)

M_FIRST = M_HEADER + 1
M_LAST = M_FIRST + 199

example_moves = [
    (datetime.date(2026, 1, 3), "ELE-001", "Entrada", 20, "Compra a proveedor"),
    (datetime.date(2026, 1, 6), "ELE-001", "Salida", 5, "Venta mostrador"),
    (datetime.date(2026, 1, 8), "ALI-001", "Entrada", 40, "Compra a proveedor"),
    (datetime.date(2026, 1, 10), "ALI-001", "Salida", 12, "Venta mostrador"),
    (datetime.date(2026, 1, 12), "ROP-002", "Entrada", 10, "Compra a proveedor"),
    (datetime.date(2026, 1, 15), "ROP-002", "Salida", 6, "Venta online"),
    (datetime.date(2026, 1, 18), "HOG-001", "Entrada", 20, "Compra a proveedor"),
    (datetime.date(2026, 1, 20), "HOG-001", "Salida", 8, "Venta mostrador"),
    (datetime.date(2026, 2, 2), "PAP-001", "Entrada", 30, "Compra a proveedor"),
    (datetime.date(2026, 2, 5), "PAP-001", "Salida", 10, "Venta mostrador"),
    (datetime.date(2026, 2, 8), "ELE-002", "Entrada", 50, "Compra a proveedor"),
    (datetime.date(2026, 2, 10), "ELE-002", "Salida", 15, "Venta online"),
    (datetime.date(2026, 2, 14), "ALI-002", "Entrada", 10, "Compra a proveedor"),
    (datetime.date(2026, 2, 18), "ROP-001", "Salida", 5, "Venta mostrador"),
    (datetime.date(2026, 3, 2), "ELE-001", "Salida", 8, "Venta online"),
    (datetime.date(2026, 3, 5), "ROP-002", "Salida", 10, "Venta de liquidación"),
]

for i in range(M_LAST - M_FIRST + 1):
    rr = M_FIRST + i
    if i < len(example_moves):
        fecha, codigo, tipo, cant, ref = example_moves[i]
        ws.cell(row=rr, column=1, value=fecha).number_format = "DD/MM/YYYY"
        ws.cell(row=rr, column=2, value=codigo)
        ws.cell(row=rr, column=4, value=tipo)
        ws.cell(row=rr, column=5, value=cant)
        ws.cell(row=rr, column=6, value=ref)
    else:
        ws.cell(row=rr, column=1).number_format = "DD/MM/YYYY"

    for col in (1, 2, 4, 5, 6):
        cell = ws.cell(row=rr, column=col)
        cell.font = f_input
        cell.fill = fill_input
        cell.border = border_all
        if col == 5:
            cell.number_format = "#,##0"
            cell.alignment = align_center
        elif col == 1:
            cell.alignment = align_center
        else:
            cell.alignment = align_left

    ncell = ws.cell(row=rr, column=3,
                     value=f'=IF(B{rr}="","",IFERROR(INDEX(Productos!$B${P_FIRST}:$B${P_LAST},MATCH(B{rr},{PROD_CODE_RANGE},0)),"Código no encontrado"))')
    ncell.font = f_body
    ncell.fill = fill_light
    ncell.border = border_all
    ncell.alignment = align_left

dv_codigo = DataValidation(type="list", formula1=f"={PROD_CODE_RANGE}", allow_blank=True)
ws.add_data_validation(dv_codigo)
dv_codigo.add(f"B{M_FIRST}:B{M_LAST}")

dv_tipo_m = DataValidation(type="list", formula1='"Entrada,Salida"', allow_blank=True)
ws.add_data_validation(dv_tipo_m)
dv_tipo_m.add(f"D{M_FIRST}:D{M_LAST}")

dv_cant = DataValidation(type="decimal", operator="greaterThanOrEqual", formula1="0",
                          allow_blank=True, showErrorMessage=True,
                          errorTitle="Cantidad inválida", error="Ingresa un número igual o mayor a 0.")
ws.add_data_validation(dv_cant)
dv_cant.add(f"E{M_FIRST}:E{M_LAST}")

lock_all(ws)
unlock_range(ws, f"A{M_FIRST}:B{M_LAST}")
unlock_range(ws, f"D{M_FIRST}:F{M_LAST}")

MOV_CODIGO_RANGE = f"Movimientos!$B${M_FIRST}:$B${M_LAST}"
MOV_TIPO_RANGE = f"Movimientos!$D${M_FIRST}:$D${M_LAST}"
MOV_CANT_RANGE = f"Movimientos!$E${M_FIRST}:$E${M_LAST}"

# =========================================================
# 7. RESUMEN
# =========================================================
ws = wb.create_sheet("Resumen")
set_col_widths(ws, [3, 26, 18, 12, 12, 12, 12, 16, 12, 3])
ws.sheet_view.showGridLines = False

ws.merge_cells("B2:H2")
ws["B2"] = "RESUMEN DINÁMICO (se actualiza solo)"
ws["B2"].font = f_h1
ws.row_dimensions[2].height = 26
ws.merge_cells("B3:H3")
ws["B3"] = '=" Moneda: " & Configuracion!$C$8'
ws["B3"].font = f_note

hr = 5
res_headers = ["Producto", "Categoría", "Stock mínimo", "Entradas", "Salidas", "Stock actual", "Valor inventario", "Estado"]
for i, h in enumerate(res_headers):
    ws.cell(row=hr, column=2 + i, value=h)
style_header_row(ws, hr, 2, 9, height=28)

RES_FIRST = hr + 1
RES_LAST = RES_FIRST + (P_LAST - P_FIRST)

for i in range(P_LAST - P_FIRST + 1):
    rr = RES_FIRST + i
    p_rr = P_FIRST + i
    ws.cell(row=rr, column=2, value=f'=IF(Productos!$A${p_rr}="","",Productos!$B${p_rr})')
    ws.cell(row=rr, column=3, value=f'=IF(Productos!$A${p_rr}="","",Productos!$C${p_rr})')
    ws.cell(row=rr, column=4, value=f'=IF(Productos!$A${p_rr}="","",Productos!$E${p_rr})')
    ws.cell(row=rr, column=5,
            value=f'=SUMIFS({MOV_CANT_RANGE},{MOV_CODIGO_RANGE},Productos!$A${p_rr},{MOV_TIPO_RANGE},"Entrada")')
    ws.cell(row=rr, column=6,
            value=f'=SUMIFS({MOV_CANT_RANGE},{MOV_CODIGO_RANGE},Productos!$A${p_rr},{MOV_TIPO_RANGE},"Salida")')
    ws.cell(row=rr, column=7,
            value=f'=IF(Productos!$A${p_rr}="","",Productos!$H${p_rr}+E{rr}-F{rr})')
    ws.cell(row=rr, column=8,
            value=f'=IF(Productos!$A${p_rr}="",0,G{rr}*Productos!$F${p_rr}*Configuracion!$C${ROW_FACTOR})')
    ws.cell(row=rr, column=9,
            value=f'=IF(Productos!$A${p_rr}="","",IF(G{rr}<=D{rr},"Bajo stock","Normal"))')
    for c in range(2, 10):
        cell = ws.cell(row=rr, column=c)
        cell.border = border_all
        cell.font = f_body
        if c in (3, 4, 5, 6):
            cell.number_format = "#,##0"
            cell.alignment = align_center
        elif c == 7:
            cell.number_format = "#,##0.00"
            cell.alignment = align_center
        elif c == 9:
            cell.alignment = align_center

tot_row = RES_LAST + 1
ws.cell(row=tot_row, column=2, value="TOTAL").font = f_body_b
for c, col_letter in zip(range(5, 8), ["D", "E", "F"]):
    cell = ws.cell(row=tot_row, column=c,
                    value=f"=SUM({col_letter}{RES_FIRST}:{col_letter}{RES_LAST})")
    cell.number_format = "#,##0"
    cell.font = f_body_b
    cell.alignment = align_center
    cell.fill = fill_light
cell = ws.cell(row=tot_row, column=8, value=f"=SUM(G{RES_FIRST}:G{RES_LAST})")
cell.number_format = "#,##0.00"
cell.font = f_body_b
cell.alignment = align_center
cell.fill = fill_light
n_bajo = ws.cell(row=tot_row, column=9, value=f'=COUNTIF(I{RES_FIRST}:I{RES_LAST},"Bajo stock")')
n_bajo.font = f_body_b
n_bajo.fill = fill_light
n_bajo.alignment = align_center
for c in range(2, 10):
    ws.cell(row=tot_row, column=c).border = border_all
    ws.cell(row=tot_row, column=c).fill = fill_light
RES_TOTAL_ROW = tot_row

# ---- Valor de inventario por categoría (for the pie chart) ----
r1 = tot_row + 3
ws.cell(row=r1, column=2, value="VALOR DE INVENTARIO POR CATEGORÍA")
ws.merge_cells(start_row=r1, start_column=2, end_row=r1, end_column=3)
ws.cell(row=r1, column=2).font = f_h2
for c in range(2, 4):
    ws.cell(row=r1, column=c).fill = fill_header2
ws.row_dimensions[r1].height = 20

hr2 = r1 + 1
ws.cell(row=hr2, column=2, value="Categoría")
ws.cell(row=hr2, column=3, value="Valor inventario")
style_header_row(ws, hr2, 2, 3, height=18)

CATVAL_FIRST = hr2 + 1
for i, cat in enumerate(categorias_prod):
    rr = CATVAL_FIRST + i
    ws.cell(row=rr, column=2, value=f"=Categorias!$B${CAT_FIRST + i}")
    ws.cell(row=rr, column=3,
            value=f'=SUMIFS($H${RES_FIRST}:$H${RES_LAST},$C${RES_FIRST}:$C${RES_LAST},$B{rr})')
    ws.cell(row=rr, column=2).border = border_all
    ws.cell(row=rr, column=3).border = border_all
    ws.cell(row=rr, column=2).font = f_body
    ws.cell(row=rr, column=3).font = f_body
    ws.cell(row=rr, column=3).number_format = "#,##0.00"
    ws.cell(row=rr, column=3).alignment = align_center
CATVAL_LAST = CATVAL_FIRST + len(categorias_prod) - 1

lock_all(ws)

# =========================================================
# 8. DASHBOARD
# =========================================================
ws = wb.create_sheet("Dashboard")
set_col_widths(ws, [3, 16, 16, 16, 16, 16, 16, 16, 3])
ws.sheet_view.showGridLines = False

ws.merge_cells("B2:H2")
ws["B2"] = "DASHBOARD · CONTROL DE INVENTARIO"
ws["B2"].font = f_h1
ws.row_dimensions[2].height = 28
ws.merge_cells("B3:H3")
ws["B3"] = '=" Moneda de visualización: " & Configuracion!$C$8'
ws["B3"].font = f_note

kpi_row = 5
card_defs = [
    ("PRODUCTOS REGISTRADOS", f'=COUNTIF(Productos!$A${P_FIRST}:$A${P_LAST},"<>")', TEAL),
    ("VALOR TOTAL INVENTARIO", f"=Configuracion!$C${ROW_SIMBOLO_ACTUAL}&\" \"&TEXT(Resumen!H{RES_TOTAL_ROW},\"#,##0.00\")", NAVY),
    ("UNIDADES EN STOCK", f"=TEXT(SUM(Resumen!G{RES_FIRST}:G{RES_LAST}),\"#,##0\")", GOLD),
    ("PRODUCTOS BAJO STOCK", f"=Resumen!I{RES_TOTAL_ROW}", RED_NEG),
]
col_starts = [2, 4, 6, 8]
for (label, formula, color), col0 in zip(card_defs, col_starts):
    ws.merge_cells(start_row=kpi_row, start_column=col0, end_row=kpi_row, end_column=col0 + 1)
    lc = ws.cell(row=kpi_row, column=col0, value=label)
    lc.font = f_kpi_label
    lc.alignment = align_center
    for c in range(col0, col0 + 2):
        ws.cell(row=kpi_row, column=c).fill = PatternFill("solid", fgColor=color)
    ws.merge_cells(start_row=kpi_row + 1, start_column=col0, end_row=kpi_row + 2, end_column=col0 + 1)
    vc = ws.cell(row=kpi_row + 1, column=col0, value=formula)
    vc.font = f_kpi_value
    vc.alignment = align_center
    for rr in (kpi_row + 1, kpi_row + 2):
        for c in range(col0, col0 + 2):
            ws.cell(row=rr, column=c).fill = PatternFill("solid", fgColor=color)

ws.row_dimensions[kpi_row].height = 18
ws.row_dimensions[kpi_row + 1].height = 22
ws.row_dimensions[kpi_row + 2].height = 22

chart_top = kpi_row + 4

pie = PieChart()
pie.title = "Valor de inventario por categoría"
data_pie = Reference(wb["Resumen"], min_col=3, min_row=hr2, max_row=CATVAL_LAST)
cats_pie = Reference(wb["Resumen"], min_col=2, min_row=CATVAL_FIRST, max_row=CATVAL_LAST)
pie.add_data(data_pie, titles_from_data=True)
pie.set_categories(cats_pie)
pie.height = 9
pie.width = 15
pie.dataLabels = DataLabelList()
pie.dataLabels.showPercent = True
pie.dataLabels.showVal = False
pie.dataLabels.showCatName = False
pie.dataLabels.showSerName = False
pie.dataLabels.showLegendKey = False
pie.dataLabels.showBubbleSize = False
ws.add_chart(pie, f"B{chart_top}")

bar = BarChart()
bar.type = "col"
bar.title = "Stock actual por producto"
bar.style = 10
bdata = Reference(wb["Resumen"], min_col=7, min_row=hr, max_row=RES_LAST)
bcats = Reference(wb["Resumen"], min_col=2, min_row=RES_FIRST, max_row=RES_LAST)
bar.add_data(bdata, titles_from_data=True)
bar.set_categories(bcats)
bar.height = 9
bar.width = 15
ws.add_chart(bar, f"E{chart_top}")

bar2 = BarChart()
bar2.type = "bar"
bar2.title = "Valor de inventario por producto"
bar2.style = 12
bdata2 = Reference(wb["Resumen"], min_col=8, min_row=hr, max_row=RES_LAST)
bcats2 = Reference(wb["Resumen"], min_col=2, min_row=RES_FIRST, max_row=RES_LAST)
bar2.add_data(bdata2, titles_from_data=True)
bar2.set_categories(bcats2)
bar2.height = 9
bar2.width = 30
ws.add_chart(bar2, f"B{chart_top + 19}")

lock_all(ws)

# =========================================================
wb.defined_names["TipoCambio"] = DefinedName("TipoCambio", attr_text=f"Configuracion!$C${cfg['ROW_TC']}")
wb.defined_names["MonedaVisualizacion"] = DefinedName("MonedaVisualizacion", attr_text=f"Configuracion!$C${ROW_VIS}")
wb.defined_names["FactorConversion"] = DefinedName("FactorConversion", attr_text=f"Configuracion!$C${ROW_FACTOR}")

order = ["Portada", "Instrucciones", "Configuracion", "Categorias", "Productos", "Movimientos", "Resumen", "Dashboard"]
wb._sheets = [wb[name] for name in order]
tab_colors = {
    "Portada": NAVY, "Instrucciones": TEAL, "Configuracion": GOLD, "Categorias": GOLD,
    "Productos": TEAL, "Movimientos": TEAL, "Resumen": NAVY, "Dashboard": NAVY,
}
for name, color in tab_colors.items():
    wb[name].sheet_properties.tabColor = color

wb.active = wb.index(wb["Dashboard"])
wb.calculation.fullCalcOnLoad = True

for name in order:
    psheet = wb[name]
    psheet.page_setup.orientation = "landscape"
    psheet.page_setup.fitToWidth = 1
    psheet.page_setup.fitToHeight = 0
    psheet.sheet_properties.pageSetUpPr.fitToPage = True
    psheet.page_margins.left = 0.4
    psheet.page_margins.right = 0.4
    psheet.page_margins.top = 0.5
    psheet.page_margins.bottom = 0.5

out_path = os.path.join(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
    "productos", "04-control-inventario", "Control-de-Inventario.xlsx",
)
os.makedirs(os.path.dirname(out_path), exist_ok=True)
wb.save(out_path)
print("Saved:", out_path)
print("Key rows:", dict(RES_TOTAL_ROW=RES_TOTAL_ROW, P_LAST=P_LAST))
