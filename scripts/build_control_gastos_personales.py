import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, NamedStyle, Protection
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.chart import PieChart, BarChart, LineChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.formatting.rule import CellIsRule
import datetime
import os

# ---------- Palette / Fonts ----------
NAVY = "1F3B57"
NAVY_DARK = "16283C"
TEAL = "2CA58D"
GOLD = "C9A24B"
LIGHT_GRAY = "F2F4F6"
MID_GRAY = "D9DEE3"
WHITE = "FFFFFF"
DARK_TEXT = "1A1A1A"
RED_NEG = "B23B3B"
GREEN_POS = "2E7D46"
INPUT_FILL = "FFF6D9"   # editable cells
LOCKED_TEXT = "555555"

FONT_NAME = "Arial"

f_title = Font(name=FONT_NAME, size=22, bold=True, color=WHITE)
f_subtitle = Font(name=FONT_NAME, size=12, color=WHITE)
f_h1 = Font(name=FONT_NAME, size=16, bold=True, color=NAVY)
f_h2 = Font(name=FONT_NAME, size=12, bold=True, color=WHITE)
f_h3 = Font(name=FONT_NAME, size=11, bold=True, color=NAVY)
f_body = Font(name=FONT_NAME, size=10, color=DARK_TEXT)
f_body_b = Font(name=FONT_NAME, size=10, bold=True, color=DARK_TEXT)
f_kpi_label = Font(name=FONT_NAME, size=10, bold=True, color=WHITE)
f_kpi_value = Font(name=FONT_NAME, size=18, bold=True, color=WHITE)
f_note = Font(name=FONT_NAME, size=9, italic=True, color=LOCKED_TEXT)
f_input = Font(name=FONT_NAME, size=10, color="0000FF")

fill_header = PatternFill("solid", fgColor=NAVY)
fill_header2 = PatternFill("solid", fgColor=TEAL)
fill_light = PatternFill("solid", fgColor=LIGHT_GRAY)
fill_input = PatternFill("solid", fgColor=INPUT_FILL)
fill_gold = PatternFill("solid", fgColor=GOLD)
fill_white = PatternFill("solid", fgColor=WHITE)

thin = Side(style="thin", color=MID_GRAY)
border_all = Border(left=thin, right=thin, top=thin, bottom=thin)

align_center = Alignment(horizontal="center", vertical="center")
align_left = Alignment(horizontal="left", vertical="center")
align_wrap = Alignment(horizontal="left", vertical="top", wrap_text=True)


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

# =========================================================
# 1. PORTADA
# =========================================================
ws = wb.create_sheet("Portada")
set_col_widths(ws, [4, 22, 22, 22, 22, 22, 4])
ws.sheet_view.showGridLines = False

ws.merge_cells("B2:F2")
ws["B2"] = "LÍNEA DE PRODUCTOS DIGITALES · PLANTILLAS EXCEL"
ws["B2"].font = Font(name=FONT_NAME, size=10, bold=True, color=TEAL)
ws["B2"].alignment = align_left

ws.merge_cells("B4:F6")
ws["B4"] = "CONTROL DE GASTOS PERSONALES"
ws["B4"].font = Font(name=FONT_NAME, size=28, bold=True, color=NAVY)
ws["B4"].alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

ws.merge_cells("B7:F8")
ws["B7"] = "Plantilla profesional para registrar tus ingresos y gastos, visualizar tu balance mensual y controlar tus finanzas personales — con panel de control, gráficos automáticos y selector de moneda (local / USD)."
ws["B7"].font = Font(name=FONT_NAME, size=11, color=DARK_TEXT)
ws["B7"].alignment = align_wrap

for r in range(2, 9):
    ws.row_dimensions[r].height = 20
ws.row_dimensions[4].height = 30

features = [
    "Dashboard con indicadores clave (KPIs) y gráficos automáticos",
    "Fórmulas automáticas: nada que calcular a mano",
    "Gráficos: pastel, barras y tendencia mensual",
    "Resumen dinámico tipo tabla dinámica por categoría y por mes",
    "Listas desplegables y validación de datos en cada registro",
    "Selector de moneda: local o USD, con tipo de cambio editable",
    "Hoja de instrucciones paso a paso",
    "Celdas y fórmulas protegidas: solo editas donde debes",
    "Diseño profesional, listo para usar",
    "Ejemplos precargados para que veas cómo funciona",
]

row0 = 10
ws.merge_cells(f"B{row0}:F{row0}")
ws[f"B{row0}"] = "QUÉ INCLUYE ESTA PLANTILLA"
ws[f"B{row0}"].font = f_h3
ws[f"B{row0}"].fill = fill_light
ws.row_dimensions[row0].height = 20
for c in range(2, 7):
    ws.cell(row=row0, column=c).fill = fill_light

r = row0 + 1
for feat in features:
    ws.merge_cells(f"B{r}:F{r}")
    ws[f"B{r}"] = f"✔  {feat}"
    ws[f"B{r}"].font = f_body
    ws.row_dimensions[r].height = 18
    r += 1

r += 1
ws.merge_cells(f"B{r}:F{r}")
ws[f"B{r}"] = "Producto digital · Precio orientativo: $6.99 USD  |  Licencia de uso personal"
ws[f"B{r}"].font = Font(name=FONT_NAME, size=10, bold=True, color=GOLD)
r += 2
ws.merge_cells(f"B{r}:F{r}")
ws[f"B{r}"] = "© Tu Marca de Plantillas · www.tumarca.com"
ws[f"B{r}"].font = f_note

lock_all(ws)

# =========================================================
# 2. INSTRUCCIONES
# =========================================================
ws = wb.create_sheet("Instrucciones")
set_col_widths(ws, [3, 46, 46, 3])
ws.sheet_view.showGridLines = False

ws.merge_cells("B2:C2")
ws["B2"] = "CÓMO USAR ESTA PLANTILLA"
ws["B2"].font = f_h1
ws.row_dimensions[2].height = 26

steps = [
    ("1. Configuración",
     "Ve a la hoja 'Configuración' y define el símbolo de tu moneda local, su nombre y el tipo de cambio frente al USD. "
     "Ahí mismo eliges si el resto de la plantilla se muestra en moneda LOCAL o en USD."),
    ("2. Categorías",
     "En la hoja 'Categorías' puedes revisar o editar la lista de categorías de ingreso/gasto y los métodos de pago "
     "que aparecen como listas desplegables en 'Transacciones'."),
    ("3. Registrar movimientos",
     "En la hoja 'Transacciones' registra cada ingreso o gasto: fecha, categoría (lista desplegable), tipo, descripción, "
     "método de pago (lista desplegable) y monto en tu moneda local. La columna 'Monto en moneda seleccionada' se calcula sola."),
    ("4. Resumen automático",
     "La hoja 'Resumen' agrupa automáticamente tus movimientos por categoría y por mes, como una tabla dinámica, "
     "sin que tengas que hacer nada manualmente."),
    ("5. Dashboard",
     "La hoja 'Dashboard' muestra tus indicadores clave (ingresos, gastos, balance, tasa de ahorro) y tres gráficos "
     "que se actualizan solos cada vez que agregas movimientos."),
    ("6. Cambiar de moneda",
     "En cualquier momento cambia la celda 'Moneda de visualización' en 'Configuración' entre LOCAL y USD: "
     "Transacciones, Resumen y Dashboard se recalculan automáticamente al nuevo tipo de cambio."),
    ("7. Celdas protegidas",
     "Las hojas están protegidas para que no borres fórmulas por accidente. Solo puedes escribir en las celdas "
     "resaltadas en color crema / texto azul. Si necesitas editar otra celda, ve a Revisar → Desproteger hoja (no pide contraseña)."),
    ("8. Agregar más filas",
     "Si necesitas más de 200 movimientos, selecciona una fila dentro de la tabla de 'Transacciones', copia el formato "
     "hacia abajo y continúa registrando: las fórmulas de conversión ya están preparadas en ese rango."),
]

r = 4
for title, body in steps:
    ws.merge_cells(f"B{r}:C{r}")
    ws[f"B{r}"] = title
    ws[f"B{r}"].font = f_h3
    ws[f"B{r}"].fill = fill_light
    ws.cell(row=r, column=2).fill = fill_light
    ws.cell(row=r, column=3).fill = fill_light
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

# =========================================================
# 3. CONFIGURACIÓN
# =========================================================
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
    ("Tipo de cambio (1 USD =)", 3.75, "Cuántas unidades de tu moneda local equivalen a 1 USD. Edítalo cuando cambie."),
    ("Moneda de visualización", "Local", "Elige LOCAL o USD en la lista desplegable. Afecta Transacciones, Resumen y Dashboard."),
    ("Año de análisis", 2026, "Año que se usa para agrupar el Resumen Mensual."),
]

r = 5
cell_refs = {}
labels_key = ["nombre_moneda", "simbolo_local", "tipo_cambio", "moneda_vis", "anio"]
for key, (label, value, note) in zip(labels_key, rows):
    ws.cell(row=r, column=2, value=label).font = f_body_b
    c = ws.cell(row=r, column=3, value=value)
    c.font = f_input
    c.fill = fill_input
    c.alignment = align_center
    c.border = border_all
    ws.cell(row=r, column=4, value=note).font = f_note
    ws.cell(row=r, column=4).alignment = align_wrap
    ws.row_dimensions[r].height = 18
    cell_refs[key] = f"C{r}"
    r += 1

ROW_NOMBRE, ROW_SIMBOLO, ROW_TC, ROW_VIS, ROW_ANIO = 5, 6, 7, 8, 9

# Computed helper cells
r_calc_label = r + 1
ws.cell(row=r_calc_label, column=2, value="Valores calculados (no editar)").font = f_h3
ws.cell(row=r_calc_label, column=2).fill = fill_light
ws.cell(row=r_calc_label, column=3).fill = fill_light
ws.cell(row=r_calc_label, column=4).fill = fill_light
ws.row_dimensions[r_calc_label].height = 18

ROW_FACTOR = r_calc_label + 1
ws.cell(row=ROW_FACTOR, column=2, value="Factor de conversión aplicado").font = f_body_b
ws.cell(row=ROW_FACTOR, column=3, value=f'=IF(C{ROW_VIS}="USD",1/C{ROW_TC},1)')
ws.cell(row=ROW_FACTOR, column=3).number_format = "0.0000"
ws.cell(row=ROW_FACTOR, column=3).border = border_all
ws.cell(row=ROW_FACTOR, column=4, value="Vale 1 si vista LOCAL; vale 1/Tipo de cambio si vista USD.").font = f_note
ws.row_dimensions[ROW_FACTOR].height = 18

ROW_SIMBOLO_ACTUAL = ROW_FACTOR + 1
ws.cell(row=ROW_SIMBOLO_ACTUAL, column=2, value="Símbolo de moneda actual").font = f_body_b
ws.cell(row=ROW_SIMBOLO_ACTUAL, column=3, value=f'=IF(C{ROW_VIS}="USD","US$",C{ROW_SIMBOLO})')
ws.cell(row=ROW_SIMBOLO_ACTUAL, column=3).border = border_all
ws.cell(row=ROW_SIMBOLO_ACTUAL, column=3).alignment = align_center
ws.cell(row=ROW_SIMBOLO_ACTUAL, column=4, value="Símbolo que verás en el Dashboard según la moneda elegida.").font = f_note
ws.row_dimensions[ROW_SIMBOLO_ACTUAL].height = 18

for rr in (ROW_NOMBRE, ROW_SIMBOLO, ROW_TC, ROW_VIS, ROW_ANIO, ROW_FACTOR, ROW_SIMBOLO_ACTUAL):
    ws.cell(row=rr, column=2).border = border_all

note_row = ROW_SIMBOLO_ACTUAL + 2
ws.merge_cells(f"B{note_row}:D{note_row}")
ws[f"B{note_row}"] = "Cambia 'Moneda de visualización' entre Local y USD para recalcular automáticamente toda la plantilla."
ws[f"B{note_row}"].font = f_note
ws[f"B{note_row}"].alignment = align_wrap

lock_all(ws)
unlock_range(ws, f"C{ROW_NOMBRE}:C{ROW_NOMBRE}")
unlock_range(ws, f"C{ROW_SIMBOLO}:C{ROW_SIMBOLO}")
unlock_range(ws, f"C{ROW_TC}:C{ROW_TC}")
unlock_range(ws, f"C{ROW_VIS}:C{ROW_VIS}")
unlock_range(ws, f"C{ROW_ANIO}:C{ROW_ANIO}")

dv_moneda = DataValidation(type="list", formula1='"Local,USD"', allow_blank=False, showDropDown=False)
ws.add_data_validation(dv_moneda)
dv_moneda.add(ws[f"C{ROW_VIS}"])

# =========================================================
# 4. CATEGORIAS
# =========================================================
ws = wb.create_sheet("Categorias")
set_col_widths(ws, [3, 30, 26, 3])
ws.sheet_view.showGridLines = False

ws.merge_cells("B2:C2")
ws["B2"] = "CATEGORÍAS Y MÉTODOS DE PAGO"
ws["B2"].font = f_h1
ws.row_dimensions[2].height = 26

ws.cell(row=4, column=2, value="Categoría (ingreso/gasto)")
ws.cell(row=4, column=3, value="Método de pago")
style_header_row(ws, 4, 2, 3)

categorias = [
    "Salario", "Freelance / Ingresos extra", "Otros ingresos",
    "Vivienda (alquiler/hipoteca)", "Alimentación", "Transporte",
    "Servicios (luz, agua, internet)", "Salud", "Educación",
    "Entretenimiento", "Ropa y cuidado personal", "Ahorro e inversión",
    "Deudas y préstamos", "Otros gastos",
]
metodos = ["Efectivo", "Tarjeta de débito", "Tarjeta de crédito", "Transferencia bancaria"]

CAT_FIRST_ROW = 5
CAT_LAST_ROW = CAT_FIRST_ROW + len(categorias) - 1
for i, cat in enumerate(categorias):
    rr = CAT_FIRST_ROW + i
    c = ws.cell(row=rr, column=2, value=cat)
    c.font = f_input
    c.fill = fill_input
    c.border = border_all
    ws.row_dimensions[rr].height = 16

MET_FIRST_ROW = 5
MET_LAST_ROW = MET_FIRST_ROW + len(metodos) - 1
for i, met in enumerate(metodos):
    rr = MET_FIRST_ROW + i
    c = ws.cell(row=rr, column=3, value=met)
    c.font = f_input
    c.fill = fill_input
    c.border = border_all

note_row = max(CAT_LAST_ROW, MET_LAST_ROW) + 2
ws.merge_cells(f"B{note_row}:C{note_row}")
ws[f"B{note_row}"] = "Puedes renombrar o agregar categorías/métodos aquí; las listas desplegables de 'Transacciones' se actualizan solas."
ws[f"B{note_row}"].font = f_note
ws[f"B{note_row}"].alignment = align_wrap

lock_all(ws)
unlock_range(ws, f"B{CAT_FIRST_ROW}:B{CAT_LAST_ROW+10}")
unlock_range(ws, f"C{MET_FIRST_ROW}:C{MET_LAST_ROW+10}")

# =========================================================
# 5. TRANSACCIONES
# =========================================================
ws = wb.create_sheet("Transacciones")
set_col_widths(ws, [14, 26, 12, 32, 22, 19, 23])
ws.sheet_view.showGridLines = False
ws.freeze_panes = "A5"

ws.merge_cells("A2:G2")
ws["A2"] = "REGISTRO DE INGRESOS Y GASTOS"
ws["A2"].font = f_h1
ws.row_dimensions[2].height = 26

ws.merge_cells("A3:G3")
ws["A3"] = "Completa las columnas A a F para cada movimiento. La columna G se calcula automáticamente."
ws["A3"].font = f_note

HEADER_ROW = 4
headers = ["Fecha", "Categoría", "Tipo", "Descripción", "Método de pago",
           "Monto (moneda local)", "Monto (moneda seleccionada)"]
for i, h in enumerate(headers, start=1):
    ws.cell(row=HEADER_ROW, column=i, value=h)
style_header_row(ws, HEADER_ROW, 1, 7)

FIRST_DATA_ROW = HEADER_ROW + 1
LAST_DATA_ROW = FIRST_DATA_ROW + 195  # up to 200 rows total

example_data = [
    (datetime.date(2026, 1, 3), "Salario", "Ingreso", "Sueldo mensual", "Transferencia bancaria", 4500),
    (datetime.date(2026, 1, 5), "Vivienda (alquiler/hipoteca)", "Gasto", "Alquiler de enero", "Transferencia bancaria", 1200),
    (datetime.date(2026, 1, 6), "Alimentación", "Gasto", "Supermercado semanal", "Tarjeta de débito", 180),
    (datetime.date(2026, 1, 8), "Transporte", "Gasto", "Gasolina", "Tarjeta de crédito", 60),
    (datetime.date(2026, 1, 10), "Servicios (luz, agua, internet)", "Gasto", "Recibo de electricidad", "Transferencia bancaria", 75),
    (datetime.date(2026, 1, 12), "Freelance / Ingresos extra", "Ingreso", "Proyecto de diseño", "Transferencia bancaria", 350),
    (datetime.date(2026, 1, 15), "Entretenimiento", "Gasto", "Cine y cena", "Efectivo", 45),
    (datetime.date(2026, 1, 20), "Salud", "Gasto", "Consulta médica", "Efectivo", 90),
    (datetime.date(2026, 1, 25), "Ahorro e inversión", "Gasto", "Aporte a fondo de ahorro", "Transferencia bancaria", 300),
    (datetime.date(2026, 2, 3), "Salario", "Ingreso", "Sueldo mensual", "Transferencia bancaria", 4500),
    (datetime.date(2026, 2, 5), "Vivienda (alquiler/hipoteca)", "Gasto", "Alquiler de febrero", "Transferencia bancaria", 1200),
    (datetime.date(2026, 2, 7), "Alimentación", "Gasto", "Supermercado semanal", "Tarjeta de débito", 210),
    (datetime.date(2026, 2, 9), "Transporte", "Gasto", "Mantenimiento de auto", "Tarjeta de crédito", 120),
    (datetime.date(2026, 2, 14), "Entretenimiento", "Gasto", "Suscripciones streaming", "Tarjeta de crédito", 25),
    (datetime.date(2026, 2, 18), "Educación", "Gasto", "Curso en línea", "Tarjeta de crédito", 80),
    (datetime.date(2026, 2, 22), "Otros ingresos", "Ingreso", "Venta de artículo usado", "Efectivo", 150),
    (datetime.date(2026, 3, 3), "Salario", "Ingreso", "Sueldo mensual", "Transferencia bancaria", 4600),
    (datetime.date(2026, 3, 5), "Vivienda (alquiler/hipoteca)", "Gasto", "Alquiler de marzo", "Transferencia bancaria", 1200),
    (datetime.date(2026, 3, 8), "Alimentación", "Gasto", "Supermercado semanal", "Tarjeta de débito", 195),
    (datetime.date(2026, 3, 15), "Deudas y préstamos", "Gasto", "Cuota de préstamo", "Transferencia bancaria", 250),
]

CAT_RANGE = f"Categorias!$B${CAT_FIRST_ROW}:$B${CAT_LAST_ROW}"
MET_RANGE = f"Categorias!$C${MET_FIRST_ROW}:$C${MET_LAST_ROW}"

for i in range(0, LAST_DATA_ROW - FIRST_DATA_ROW + 1):
    rr = FIRST_DATA_ROW + i
    if i < len(example_data):
        fecha, cat, tipo, desc, met, monto = example_data[i]
        ws.cell(row=rr, column=1, value=fecha).number_format = "DD/MM/YYYY"
        ws.cell(row=rr, column=2, value=cat)
        ws.cell(row=rr, column=3, value=tipo)
        ws.cell(row=rr, column=4, value=desc)
        ws.cell(row=rr, column=5, value=met)
        ws.cell(row=rr, column=6, value=monto)
    else:
        ws.cell(row=rr, column=1).number_format = "DD/MM/YYYY"

    for col in range(1, 7):
        cell = ws.cell(row=rr, column=col)
        cell.font = f_input
        cell.fill = fill_input if i < 200 else fill_white
        cell.border = border_all
        if col == 6:
            cell.number_format = "#,##0.00"
            cell.alignment = align_center
        elif col == 1:
            cell.alignment = align_center
        else:
            cell.alignment = align_left

    gcell = ws.cell(row=rr, column=7,
                     value=f'=IF(F{rr}="","",F{rr}*Configuracion!$C${ROW_FACTOR})')
    gcell.number_format = "#,##0.00"
    gcell.font = f_body
    gcell.fill = fill_light
    gcell.border = border_all
    gcell.alignment = align_center

# Data validations for the whole editable range
dv_cat = DataValidation(type="list", formula1=f"={CAT_RANGE}", allow_blank=True)
ws.add_data_validation(dv_cat)
dv_cat.add(f"B{FIRST_DATA_ROW}:B{LAST_DATA_ROW}")

dv_tipo = DataValidation(type="list", formula1='"Ingreso,Gasto"', allow_blank=True)
ws.add_data_validation(dv_tipo)
dv_tipo.add(f"C{FIRST_DATA_ROW}:C{LAST_DATA_ROW}")

dv_met = DataValidation(type="list", formula1=f"={MET_RANGE}", allow_blank=True)
ws.add_data_validation(dv_met)
dv_met.add(f"E{FIRST_DATA_ROW}:E{LAST_DATA_ROW}")

dv_monto = DataValidation(type="decimal", operator="greaterThanOrEqual", formula1="0",
                           allow_blank=True, showErrorMessage=True,
                           errorTitle="Monto inválido", error="Ingresa un número igual o mayor a 0.")
ws.add_data_validation(dv_monto)
dv_monto.add(f"F{FIRST_DATA_ROW}:F{LAST_DATA_ROW}")

lock_all(ws)
unlock_range(ws, f"A{FIRST_DATA_ROW}:F{LAST_DATA_ROW}")

TRANS_G_RANGE = f"Transacciones!$G${FIRST_DATA_ROW}:$G${LAST_DATA_ROW}"
TRANS_B_RANGE = f"Transacciones!$B${FIRST_DATA_ROW}:$B${LAST_DATA_ROW}"
TRANS_C_RANGE = f"Transacciones!$C${FIRST_DATA_ROW}:$C${LAST_DATA_ROW}"
TRANS_A_RANGE = f"Transacciones!$A${FIRST_DATA_ROW}:$A${LAST_DATA_ROW}"

# =========================================================
# 6. RESUMEN  (SUMIFS-driven dynamic summary tables)
# =========================================================
ws = wb.create_sheet("Resumen")
set_col_widths(ws, [3, 30, 16, 16, 16, 16, 3])
ws.sheet_view.showGridLines = False

ws.merge_cells("B2:F2")
ws["B2"] = "RESUMEN DINÁMICO (se actualiza solo)"
ws["B2"].font = f_h1
ws.row_dimensions[2].height = 26
ws.merge_cells("B3:F3")
ws["B3"] = "Tablas de resumen automáticas por categoría y por mes, equivalentes a una tabla dinámica, calculadas con fórmulas."
ws["B3"].font = f_note

# --- Tabla 1: Resumen por categoría ---
r0 = 5
ws.cell(row=r0, column=2, value="RESUMEN POR CATEGORÍA")
ws.merge_cells(start_row=r0, start_column=2, end_row=r0, end_column=5)
ws.cell(row=r0, column=2).font = f_h2
ws.cell(row=r0, column=2).fill = fill_header2
ws.cell(row=r0, column=2).alignment = align_center
for c in range(2, 6):
    ws.cell(row=r0, column=c).fill = fill_header2
ws.row_dimensions[r0].height = 20

hr = r0 + 1
cat_headers = ["Categoría", "Ingresos", "Gastos", "Neto"]
for i, h in enumerate(cat_headers):
    ws.cell(row=hr, column=2 + i, value=h)
style_header_row(ws, hr, 2, 5, fill=fill_header, height=18)

CAT_SUMMARY_FIRST = hr + 1
for i, cat in enumerate(categorias):
    rr = CAT_SUMMARY_FIRST + i
    ws.cell(row=rr, column=2, value=f"=Categorias!$B${CAT_FIRST_ROW + i}")
    ws.cell(row=rr, column=3,
            value=f'=SUMIFS({TRANS_G_RANGE},{TRANS_B_RANGE},$B{rr},{TRANS_C_RANGE},"Ingreso")')
    ws.cell(row=rr, column=4,
            value=f'=SUMIFS({TRANS_G_RANGE},{TRANS_B_RANGE},$B{rr},{TRANS_C_RANGE},"Gasto")')
    ws.cell(row=rr, column=5, value=f"=C{rr}-D{rr}")
    for c in range(2, 6):
        cell = ws.cell(row=rr, column=c)
        cell.border = border_all
        cell.font = f_body
        if c >= 3:
            cell.number_format = "#,##0.00"
            cell.alignment = align_center
CAT_SUMMARY_LAST = CAT_SUMMARY_FIRST + len(categorias) - 1

tot_row = CAT_SUMMARY_LAST + 1
ws.cell(row=tot_row, column=2, value="TOTAL").font = f_body_b
for c, col_letter in zip(range(3, 6), ["C", "D", "E"]):
    cell = ws.cell(row=tot_row, column=c,
                    value=f"=SUM({col_letter}{CAT_SUMMARY_FIRST}:{col_letter}{CAT_SUMMARY_LAST})")
    cell.number_format = "#,##0.00"
    cell.font = f_body_b
    cell.alignment = align_center
    cell.fill = fill_light
for c in range(2, 6):
    ws.cell(row=tot_row, column=c).border = border_all
    ws.cell(row=tot_row, column=c).fill = fill_light

CAT_TOTAL_ROW = tot_row

# --- Tabla 2: Resumen mensual ---
r1 = tot_row + 3
ws.cell(row=r1, column=2, value="RESUMEN MENSUAL")
ws.merge_cells(start_row=r1, start_column=2, end_row=r1, end_column=6)
ws.cell(row=r1, column=2).font = f_h2
for c in range(2, 7):
    ws.cell(row=r1, column=c).fill = fill_header2
ws.row_dimensions[r1].height = 20

hr2 = r1 + 1
mon_headers = ["Mes", "Ingresos", "Gastos", "Ahorro neto", "% Ahorro"]
for i, h in enumerate(mon_headers):
    ws.cell(row=hr2, column=2 + i, value=h)
style_header_row(ws, hr2, 2, 6, fill=fill_header, height=18)

MONTHS_ES = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
             "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]

MONTH_SUMMARY_FIRST = hr2 + 1
for i, mname in enumerate(MONTHS_ES):
    rr = MONTH_SUMMARY_FIRST + i
    mnum = i + 1
    ws.cell(row=rr, column=2, value=mname)
    start_expr = f'DATE(Configuracion!$C${ROW_ANIO},{mnum},1)'
    end_expr = f'EOMONTH(DATE(Configuracion!$C${ROW_ANIO},{mnum},1),0)'
    ws.cell(row=rr, column=3,
            value=(f'=SUMIFS({TRANS_G_RANGE},{TRANS_C_RANGE},"Ingreso",'
                   f'{TRANS_A_RANGE},">="&{start_expr},{TRANS_A_RANGE},"<="&{end_expr})'))
    ws.cell(row=rr, column=4,
            value=(f'=SUMIFS({TRANS_G_RANGE},{TRANS_C_RANGE},"Gasto",'
                   f'{TRANS_A_RANGE},">="&{start_expr},{TRANS_A_RANGE},"<="&{end_expr})'))
    ws.cell(row=rr, column=5, value=f"=C{rr}-D{rr}")
    ws.cell(row=rr, column=6, value=f'=IF(C{rr}=0,0,E{rr}/C{rr})')
    for c in range(2, 7):
        cell = ws.cell(row=rr, column=c)
        cell.border = border_all
        cell.font = f_body
        if c in (3, 4, 5):
            cell.number_format = "#,##0.00"
            cell.alignment = align_center
        elif c == 6:
            cell.number_format = "0.0%"
            cell.alignment = align_center
MONTH_SUMMARY_LAST = MONTH_SUMMARY_FIRST + len(MONTHS_ES) - 1

lock_all(ws)

# =========================================================
# 7. DASHBOARD
# =========================================================
ws = wb.create_sheet("Dashboard")
set_col_widths(ws, [3, 16, 16, 16, 16, 16, 16, 16, 3])
ws.sheet_view.showGridLines = False

ws.merge_cells("B2:H2")
ws["B2"] = "DASHBOARD · CONTROL DE GASTOS PERSONALES"
ws["B2"].font = f_h1
ws.row_dimensions[2].height = 28

ws.merge_cells("B3:H3")
ws["B3"] = f'="Moneda de visualización: " & Configuracion!$C${ROW_VIS} & "   |   Año: " & Configuracion!$C${ROW_ANIO}'
ws["B3"].font = f_note

# KPI cards row
kpi_row = 5
kpi_height = 4
card_defs = [
    ("TOTAL INGRESOS", f"=Configuracion!$C${ROW_SIMBOLO_ACTUAL}&\" \"&TEXT(Resumen!C{CAT_TOTAL_ROW},\"#,##0.00\")", TEAL),
    ("TOTAL GASTOS", f"=Configuracion!$C${ROW_SIMBOLO_ACTUAL}&\" \"&TEXT(Resumen!D{CAT_TOTAL_ROW},\"#,##0.00\")", RED_NEG),
    ("BALANCE NETO", f"=Configuracion!$C${ROW_SIMBOLO_ACTUAL}&\" \"&TEXT(Resumen!E{CAT_TOTAL_ROW},\"#,##0.00\")", NAVY),
    ("TASA DE AHORRO", f'=IF(Resumen!C{CAT_TOTAL_ROW}=0,"0.0%",TEXT(Resumen!E{CAT_TOTAL_ROW}/Resumen!C{CAT_TOTAL_ROW},"0.0%"))', GOLD),
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

# ---- Charts ----
chart_top = kpi_row + 4

EXPENSE_CAT_FIRST = CAT_SUMMARY_FIRST + 3  # skip the 3 income categories (Salario, Freelance, Otros ingresos)

pie = PieChart()
pie.title = "Gastos por categoría"
data_expense = Reference(wb["Resumen"], min_col=4, min_row=EXPENSE_CAT_FIRST, max_row=CAT_SUMMARY_LAST)
cats = Reference(wb["Resumen"], min_col=2, min_row=EXPENSE_CAT_FIRST, max_row=CAT_SUMMARY_LAST)
pie.add_data(data_expense, titles_from_data=False)
pie.set_categories(cats)
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
bar.title = "Ingresos vs. Gastos por mes"
bar.style = 10
bdata = Reference(wb["Resumen"], min_col=3, max_col=4, min_row=hr2, max_row=MONTH_SUMMARY_LAST)
bcats = Reference(wb["Resumen"], min_col=2, min_row=MONTH_SUMMARY_FIRST, max_row=MONTH_SUMMARY_LAST)
bar.add_data(bdata, titles_from_data=True)
bar.set_categories(bcats)
bar.height = 9
bar.width = 15
bar.y_axis.title = None
ws.add_chart(bar, f"E{chart_top}")

line = LineChart()
line.title = "Tendencia de ahorro neto mensual"
ldata = Reference(wb["Resumen"], min_col=5, min_row=hr2, max_row=MONTH_SUMMARY_LAST)
lcats = Reference(wb["Resumen"], min_col=2, min_row=MONTH_SUMMARY_FIRST, max_row=MONTH_SUMMARY_LAST)
line.add_data(ldata, titles_from_data=True)
line.set_categories(lcats)
line.height = 9
line.width = 30
ws.add_chart(line, f"B{chart_top + 19}")

lock_all(ws)

# =========================================================
# Defined names (global, handy for power users)
# =========================================================
wb.defined_names["TipoCambio"] = DefinedName("TipoCambio", attr_text=f"Configuracion!$C${ROW_TC}")
wb.defined_names["MonedaVisualizacion"] = DefinedName("MonedaVisualizacion", attr_text=f"Configuracion!$C${ROW_VIS}")
wb.defined_names["FactorConversion"] = DefinedName("FactorConversion", attr_text=f"Configuracion!$C${ROW_FACTOR}")

# Tab order + colors
order = ["Portada", "Instrucciones", "Configuracion", "Categorias", "Transacciones", "Resumen", "Dashboard"]
wb._sheets = [wb[name] for name in order]
tab_colors = {
    "Portada": NAVY, "Instrucciones": TEAL, "Configuracion": GOLD,
    "Categorias": GOLD, "Transacciones": TEAL, "Resumen": NAVY, "Dashboard": NAVY,
}
for name, color in tab_colors.items():
    wb[name].sheet_properties.tabColor = color

wb["Dashboard"].sheet_view.showGridLines = False
wb.active = wb.index(wb["Dashboard"])
wb.calculation.fullCalcOnLoad = True

# Print / page setup so nothing gets clipped when printed or exported to PDF
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
    "productos", "01-control-gastos-personales", "Control-de-Gastos-Personales.xlsx",
)
wb.save(out_path)
print("Saved:", out_path)
print("Key rows:", dict(ROW_NOMBRE=ROW_NOMBRE, ROW_SIMBOLO=ROW_SIMBOLO, ROW_TC=ROW_TC, ROW_VIS=ROW_VIS,
                         ROW_ANIO=ROW_ANIO, ROW_FACTOR=ROW_FACTOR, ROW_SIMBOLO_ACTUAL=ROW_SIMBOLO_ACTUAL,
                         CAT_TOTAL_ROW=CAT_TOTAL_ROW, hr=hr, hr2=hr2))
