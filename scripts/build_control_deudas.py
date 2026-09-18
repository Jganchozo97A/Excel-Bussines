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

f_title = Font(name=FONT_NAME, size=22, bold=True, color=WHITE)
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


# =========================================================
# 1. PORTADA
# =========================================================
ws = wb.create_sheet("Portada")
set_col_widths(ws, [4, 22, 22, 22, 22, 22, 4])
ws.sheet_view.showGridLines = False

ws.merge_cells("B2:F2")
ws["B2"] = "LÍNEA DE PRODUCTOS DIGITALES · PLANTILLAS EXCEL"
ws["B2"].font = Font(name=FONT_NAME, size=10, bold=True, color=TEAL)

ws.merge_cells("B4:F6")
ws["B4"] = "CONTROL DE DEUDAS"
ws["B4"].font = Font(name=FONT_NAME, size=28, bold=True, color=NAVY)
ws["B4"].alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

ws.merge_cells("B7:F8")
ws["B7"] = ("Registra todas tus deudas, controla tus pagos (capital e interés) y descubre qué "
            "deuda conviene pagar primero — con panel de control, gráficos automáticos y selector "
            "de moneda (local / USD).")
ws["B7"].font = Font(name=FONT_NAME, size=11, color=DARK_TEXT)
ws["B7"].alignment = align_wrap

for r in range(2, 9):
    ws.row_dimensions[r].height = 20
ws.row_dimensions[4].height = 30

features = [
    "Dashboard con indicadores clave (KPIs) y gráficos automáticos",
    "Fórmulas automáticas: nada que calcular a mano",
    "Gráficos: pastel, barras y evolución de la deuda total",
    "Resumen dinámico tipo tabla dinámica por deuda",
    "Estrategias sugeridas de pago: bola de nieve y avalancha",
    "Listas desplegables y validación de datos en cada registro",
    "Selector de moneda: local o USD, con tipo de cambio editable",
    "Hoja de instrucciones paso a paso",
    "Celdas y fórmulas protegidas: solo editas donde debes",
    "Diseño profesional y ejemplos precargados",
]
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
ws[f"B{r}"] = "Producto digital · Precio orientativo: $7.99 USD  |  Licencia de uso personal"
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
     "Define el símbolo de tu moneda local, el tipo de cambio y si la plantilla se muestra en "
     "moneda LOCAL o USD."),
    ("2. Deudas",
     "En la hoja 'Deudas' registra cada deuda: nombre, tipo, saldo inicial, tasa de interés anual "
     "y pago mínimo mensual. El saldo actual y el % pagado se calculan solos. Hay espacio para "
     "hasta 10 deudas; si necesitas más, copia el formato de una fila hacia abajo en Deudas y Resumen."),
    ("3. Pagos",
     "Cada vez que pagues una deuda, agrégalo en la hoja 'Pagos': fecha, deuda (lista desplegable), "
     "monto total pagado e interés de ese pago. El capital amortizado se calcula solo."),
    ("4. Resumen automático",
     "La hoja 'Resumen' agrupa automáticamente tus pagos por deuda y te sugiere un orden de pago "
     "con dos estrategias: 'Bola de nieve' (pagar primero la deuda más pequeña) y 'Avalancha' "
     "(pagar primero la de mayor interés)."),
    ("5. Dashboard",
     "La hoja 'Dashboard' muestra tu deuda total, lo pagado, el interés acumulado y la evolución "
     "de tu deuda total mes a mes, con gráficos que se actualizan solos."),
    ("6. Cambiar de moneda",
     "Cambia 'Moneda de visualización' en 'Configuración' entre Local y USD: todo se recalcula "
     "automáticamente."),
    ("7. Celdas protegidas",
     "Las hojas están protegidas para que no borres fórmulas por accidente. Solo puedes escribir en "
     "las celdas resaltadas en color crema / texto azul. Contraseña de desprotección: plantilla2026."),
]

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

ROW_NOMBRE, ROW_SIMBOLO, ROW_TC, ROW_VIS, ROW_ANIO = 5, 6, 7, 8, 9
rows = [
    (ROW_NOMBRE, "Nombre de moneda local", "Soles (PEN)", "Solo referencia. Escribe el nombre de tu moneda."),
    (ROW_SIMBOLO, "Símbolo de moneda local", "S/", "Símbolo que se usará en el Dashboard cuando la vista esté en LOCAL."),
    (ROW_TC, "Tipo de cambio (1 USD =)", 3.75, "Cuántas unidades de tu moneda local equivalen a 1 USD."),
    (ROW_VIS, "Moneda de visualización", "Local", "Elige LOCAL o USD. Afecta Resumen y Dashboard."),
    (ROW_ANIO, "Año de análisis", 2026, "Año que se usa para agrupar la evolución mensual."),
]
for rr, label, value, note in rows:
    ws.cell(row=rr, column=2, value=label).font = f_body_b
    c = ws.cell(row=rr, column=3, value=value)
    c.font = f_input
    c.fill = fill_input
    c.alignment = align_center
    c.border = border_all
    ws.cell(row=rr, column=4, value=note).font = f_note
    ws.cell(row=rr, column=4).alignment = align_wrap
    ws.row_dimensions[rr].height = 18
    ws.cell(row=rr, column=2).border = border_all

r_calc_label = ROW_ANIO + 1
ws.cell(row=r_calc_label, column=2, value="Valores calculados (no editar)").font = f_h3
for cc in (2, 3, 4):
    ws.cell(row=r_calc_label, column=cc).fill = fill_light
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

note_row = ROW_SIMBOLO_ACTUAL + 2
ws.merge_cells(f"B{note_row}:D{note_row}")
ws[f"B{note_row}"] = "Cambia 'Moneda de visualización' entre Local y USD para recalcular automáticamente toda la plantilla."
ws[f"B{note_row}"].font = f_note
ws[f"B{note_row}"].alignment = align_wrap

lock_all(ws)
for rr in (ROW_NOMBRE, ROW_SIMBOLO, ROW_TC, ROW_VIS, ROW_ANIO):
    unlock_range(ws, f"C{rr}:C{rr}")

dv_moneda = DataValidation(type="list", formula1='"Local,USD"', allow_blank=False)
ws.add_data_validation(dv_moneda)
dv_moneda.add(ws[f"C{ROW_VIS}"])

# =========================================================
# 4. DEUDAS
# =========================================================
ws = wb.create_sheet("Deudas")
set_col_widths(ws, [22, 20, 15, 13, 16, 15, 15, 15, 14, 11])
ws.sheet_view.showGridLines = False
ws.freeze_panes = "A5"

ws.merge_cells("A2:J2")
ws["A2"] = "REGISTRO DE DEUDAS"
ws["A2"].font = f_h1
ws.row_dimensions[2].height = 26
ws.merge_cells("A3:J3")
ws["A3"] = "Completa las columnas A a E para cada deuda. Las columnas F a J se calculan automáticamente."
ws["A3"].font = f_note

HEADER_ROW = 4
headers = ["Deuda", "Tipo", "Saldo inicial", "Tasa interés anual", "Pago mínimo mensual",
           "Total pagado", "Interés pagado", "Capital pagado", "Saldo actual", "% pagado"]
for i, h in enumerate(headers, start=1):
    ws.cell(row=HEADER_ROW, column=i, value=h)
style_header_row(ws, HEADER_ROW, 1, 10, height=30)

FIRST_ROW = HEADER_ROW + 1
tipos = ["Tarjeta de crédito", "Préstamo personal", "Préstamo vehicular", "Hipoteca",
         "Línea de crédito", "Otro"]

example_debts = [
    ("Tarjeta Visa", "Tarjeta de crédito", 3000, 0.28, 150),
    ("Tarjeta Mastercard", "Tarjeta de crédito", 1500, 0.32, 80),
    ("Préstamo auto", "Préstamo vehicular", 8000, 0.12, 320),
    ("Préstamo personal", "Préstamo personal", 2000, 0.18, 150),
    ("Hipoteca", "Hipoteca", 45000, 0.08, 450),
]
LAST_ROW = FIRST_ROW + 9  # allow up to 10 debts (keeps pie/bar chart legends readable)

for i in range(LAST_ROW - FIRST_ROW + 1):
    rr = FIRST_ROW + i
    if i < len(example_debts):
        name, tipo, saldo, tasa, pago = example_debts[i]
        ws.cell(row=rr, column=1, value=name)
        ws.cell(row=rr, column=2, value=tipo)
        ws.cell(row=rr, column=3, value=saldo)
        ws.cell(row=rr, column=4, value=tasa)
        ws.cell(row=rr, column=5, value=pago)
    for col in range(1, 6):
        cell = ws.cell(row=rr, column=col)
        cell.font = f_input
        cell.fill = fill_input
        cell.border = border_all
        if col == 3:
            cell.number_format = "#,##0.00"
            cell.alignment = align_center
        elif col == 4:
            cell.number_format = "0.0%"
            cell.alignment = align_center
        elif col == 5:
            cell.number_format = "#,##0.00"
            cell.alignment = align_center
        else:
            cell.alignment = align_left

    # placeholders for the formula columns, filled after PAGOS sheet exists
    for col in range(6, 11):
        cell = ws.cell(row=rr, column=col)
        cell.font = f_body
        cell.fill = fill_light
        cell.border = border_all
        cell.alignment = align_center
        if col in (6, 7, 8, 9):
            cell.number_format = "#,##0.00"
        elif col == 10:
            cell.number_format = "0.0%"

dv_tipo = DataValidation(type="list", formula1='"' + ",".join(tipos) + '"', allow_blank=True)
ws.add_data_validation(dv_tipo)
dv_tipo.add(f"B{FIRST_ROW}:B{LAST_ROW}")

dv_num = DataValidation(type="decimal", operator="greaterThanOrEqual", formula1="0",
                         allow_blank=True, showErrorMessage=True,
                         errorTitle="Valor inválido", error="Ingresa un número igual o mayor a 0.")
ws.add_data_validation(dv_num)
dv_num.add(f"C{FIRST_ROW}:E{LAST_ROW}")

lock_all(ws)
unlock_range(ws, f"A{FIRST_ROW}:E{LAST_ROW}")

DEBT_NAME_RANGE = f"Deudas!$A${FIRST_ROW}:$A${LAST_ROW}"

# =========================================================
# 5. PAGOS
# =========================================================
ws = wb.create_sheet("Pagos")
set_col_widths(ws, [14, 22, 16, 16, 16, 3])
ws.sheet_view.showGridLines = False
ws.freeze_panes = "A5"

ws.merge_cells("A2:E2")
ws["A2"] = "REGISTRO DE PAGOS"
ws["A2"].font = f_h1
ws.row_dimensions[2].height = 26
ws.merge_cells("A3:E3")
ws["A3"] = "Completa las columnas A a D para cada pago. La columna E (capital) se calcula automáticamente."
ws["A3"].font = f_note

PAY_HEADER_ROW = 4
pay_headers = ["Fecha", "Deuda", "Monto pagado", "Interés del pago", "Capital amortizado"]
for i, h in enumerate(pay_headers, start=1):
    ws.cell(row=PAY_HEADER_ROW, column=i, value=h)
style_header_row(ws, PAY_HEADER_ROW, 1, 5, height=22)

PAY_FIRST_ROW = PAY_HEADER_ROW + 1
PAY_LAST_ROW = PAY_FIRST_ROW + 199

example_payments = [
    (datetime.date(2026, 1, 5), "Tarjeta Visa", 250, 65),
    (datetime.date(2026, 1, 8), "Tarjeta Mastercard", 120, 38),
    (datetime.date(2026, 1, 10), "Préstamo auto", 320, 78),
    (datetime.date(2026, 1, 12), "Préstamo personal", 150, 28),
    (datetime.date(2026, 1, 15), "Hipoteca", 450, 295),
    (datetime.date(2026, 2, 5), "Tarjeta Visa", 250, 60),
    (datetime.date(2026, 2, 8), "Tarjeta Mastercard", 120, 35),
    (datetime.date(2026, 2, 10), "Préstamo auto", 320, 74),
    (datetime.date(2026, 2, 12), "Préstamo personal", 150, 26),
    (datetime.date(2026, 2, 15), "Hipoteca", 450, 292),
    (datetime.date(2026, 3, 5), "Tarjeta Visa", 300, 58),
    (datetime.date(2026, 3, 8), "Tarjeta Mastercard", 120, 33),
    (datetime.date(2026, 3, 10), "Préstamo auto", 320, 71),
    (datetime.date(2026, 3, 12), "Préstamo personal", 150, 24),
    (datetime.date(2026, 3, 15), "Hipoteca", 450, 289),
]

for i in range(PAY_LAST_ROW - PAY_FIRST_ROW + 1):
    rr = PAY_FIRST_ROW + i
    if i < len(example_payments):
        fecha, deuda, monto, interes = example_payments[i]
        ws.cell(row=rr, column=1, value=fecha).number_format = "DD/MM/YYYY"
        ws.cell(row=rr, column=2, value=deuda)
        ws.cell(row=rr, column=3, value=monto)
        ws.cell(row=rr, column=4, value=interes)
    else:
        ws.cell(row=rr, column=1).number_format = "DD/MM/YYYY"

    for col in range(1, 5):
        cell = ws.cell(row=rr, column=col)
        cell.font = f_input
        cell.fill = fill_input
        cell.border = border_all
        if col in (3, 4):
            cell.number_format = "#,##0.00"
            cell.alignment = align_center
        elif col == 1:
            cell.alignment = align_center
        else:
            cell.alignment = align_left

    ecell = ws.cell(row=rr, column=5,
                     value=f'=IF(C{rr}="","",C{rr}-D{rr})')
    ecell.number_format = "#,##0.00"
    ecell.font = f_body
    ecell.fill = fill_light
    ecell.border = border_all
    ecell.alignment = align_center

dv_deuda = DataValidation(type="list", formula1=f"={DEBT_NAME_RANGE}", allow_blank=True)
ws.add_data_validation(dv_deuda)
dv_deuda.add(f"B{PAY_FIRST_ROW}:B{PAY_LAST_ROW}")

dv_num2 = DataValidation(type="decimal", operator="greaterThanOrEqual", formula1="0",
                          allow_blank=True, showErrorMessage=True,
                          errorTitle="Valor inválido", error="Ingresa un número igual o mayor a 0.")
ws.add_data_validation(dv_num2)
dv_num2.add(f"C{PAY_FIRST_ROW}:D{PAY_LAST_ROW}")

lock_all(ws)
unlock_range(ws, f"A{PAY_FIRST_ROW}:D{PAY_LAST_ROW}")

PAY_FECHA_RANGE = f"Pagos!$A${PAY_FIRST_ROW}:$A${PAY_LAST_ROW}"
PAY_DEUDA_RANGE = f"Pagos!$B${PAY_FIRST_ROW}:$B${PAY_LAST_ROW}"
PAY_MONTO_RANGE = f"Pagos!$C${PAY_FIRST_ROW}:$C${PAY_LAST_ROW}"
PAY_INTERES_RANGE = f"Pagos!$D${PAY_FIRST_ROW}:$D${PAY_LAST_ROW}"
PAY_CAPITAL_RANGE = f"Pagos!$E${PAY_FIRST_ROW}:$E${PAY_LAST_ROW}"

# ---- Now fill the formula columns on Deudas (needs Pagos ranges) ----
ws_d = wb["Deudas"]
for i in range(LAST_ROW - FIRST_ROW + 1):
    rr = FIRST_ROW + i
    ws_d.cell(row=rr, column=6,
              value=f'=SUMIFS({PAY_MONTO_RANGE},{PAY_DEUDA_RANGE},$A{rr})')
    ws_d.cell(row=rr, column=7,
              value=f'=SUMIFS({PAY_INTERES_RANGE},{PAY_DEUDA_RANGE},$A{rr})')
    ws_d.cell(row=rr, column=8,
              value=f'=SUMIFS({PAY_CAPITAL_RANGE},{PAY_DEUDA_RANGE},$A{rr})')
    ws_d.cell(row=rr, column=9,
              value=f'=IF($A{rr}="","",$C{rr}-$H{rr})')
    ws_d.cell(row=rr, column=10,
              value=f'=IF(OR($A{rr}="",$C{rr}=0),0,$H{rr}/$C{rr})')

DEBT_TOTAL_ROW = LAST_ROW + 1
ws_d.cell(row=DEBT_TOTAL_ROW, column=1, value="TOTAL").font = f_body_b
for col, letter in zip(range(3, 10), ["C", "D", "E", "F", "G", "H", "I"]):
    if col == 4:
        continue  # tasa de interés: no se totaliza
    cell = ws_d.cell(row=DEBT_TOTAL_ROW, column=col,
                      value=f"=SUM({letter}{FIRST_ROW}:{letter}{LAST_ROW})")
    cell.number_format = "#,##0.00"
    cell.font = f_body_b
    cell.fill = fill_light
    cell.alignment = align_center
    cell.border = border_all
ws_d.cell(row=DEBT_TOTAL_ROW, column=10,
          value=f'=IF(C{DEBT_TOTAL_ROW}=0,0,H{DEBT_TOTAL_ROW}/C{DEBT_TOTAL_ROW})')
ws_d.cell(row=DEBT_TOTAL_ROW, column=10).number_format = "0.0%"
ws_d.cell(row=DEBT_TOTAL_ROW, column=10).font = f_body_b
ws_d.cell(row=DEBT_TOTAL_ROW, column=10).fill = fill_light
ws_d.cell(row=DEBT_TOTAL_ROW, column=10).border = border_all
ws_d.cell(row=DEBT_TOTAL_ROW, column=10).alignment = align_center
for col in (1, 2, 4):
    ws_d.cell(row=DEBT_TOTAL_ROW, column=col).fill = fill_light
    ws_d.cell(row=DEBT_TOTAL_ROW, column=col).border = border_all

# =========================================================
# 6. RESUMEN
# =========================================================
ws = wb.create_sheet("Resumen")
set_col_widths(ws, [3, 22, 15, 15, 13, 16, 16, 3])
ws.sheet_view.showGridLines = False

ws.merge_cells("B2:G2")
ws["B2"] = "RESUMEN DINÁMICO (se actualiza solo)"
ws["B2"].font = f_h1
ws.row_dimensions[2].height = 26
ws.merge_cells("B3:G3")
ws["B3"] = '=" Moneda: " & Configuracion!$C$8'
ws["B3"].font = f_note

r0 = 5
ws.cell(row=r0, column=2, value="DEUDAS Y ESTRATEGIA DE PAGO SUGERIDA")
ws.merge_cells(start_row=r0, start_column=2, end_row=r0, end_column=7)
ws.cell(row=r0, column=2).font = f_h2
for c in range(2, 8):
    ws.cell(row=r0, column=c).fill = fill_header2
ws.row_dimensions[r0].height = 20

hr = r0 + 1
res_headers = ["Deuda", "Saldo inicial", "Saldo actual", "% pagado", "Tasa interés", "Bola de nieve", "Avalancha"]
for i, h in enumerate(res_headers):
    ws.cell(row=hr, column=2 + i, value=h)
style_header_row(ws, hr, 2, 8, height=28)

N_DEBTS = len(example_debts)  # visible example rows used for charts; formulas below cover full range
RES_FIRST = hr + 1
RES_LAST = RES_FIRST + (LAST_ROW - FIRST_ROW)  # same count as Deudas rows

for i in range(LAST_ROW - FIRST_ROW + 1):
    rr = RES_FIRST + i
    d_rr = FIRST_ROW + i
    ws.cell(row=rr, column=2, value=f"=Deudas!$A${d_rr}")
    ws.cell(row=rr, column=3, value=f"=Deudas!$C${d_rr}*Configuracion!$C${ROW_FACTOR}")
    ws.cell(row=rr, column=4,
            value=f'=IF(Deudas!$A${d_rr}="",0,Deudas!$I${d_rr}*Configuracion!$C${ROW_FACTOR})')
    ws.cell(row=rr, column=5, value=f"=Deudas!$J${d_rr}")
    ws.cell(row=rr, column=6, value=f"=Deudas!$D${d_rr}")
    ws.cell(row=rr, column=7,
            value=f'=IF(Deudas!$A${d_rr}="","",IF(Deudas!$I${d_rr}<=0,"Pagada",RANK(Deudas!$I${d_rr},Deudas!$I${FIRST_ROW}:$I${LAST_ROW},1)))')
    ws.cell(row=rr, column=8,
            value=f'=IF(Deudas!$A${d_rr}="","",IF(Deudas!$I${d_rr}<=0,"Pagada",RANK(Deudas!$D${d_rr},Deudas!$D${FIRST_ROW}:$D${LAST_ROW},0)))')
    for c in range(2, 9):
        cell = ws.cell(row=rr, column=c)
        cell.border = border_all
        cell.font = f_body
        if c in (3, 4):
            cell.number_format = "#,##0.00"
            cell.alignment = align_center
        elif c == 5:
            cell.number_format = "0.0%"
            cell.alignment = align_center
        elif c == 6:
            cell.number_format = "0.0%"
            cell.alignment = align_center
        elif c in (7, 8):
            cell.alignment = align_center

tot_row = RES_LAST + 1
ws.cell(row=tot_row, column=2, value="TOTAL").font = f_body_b
for c, col_letter in zip(range(3, 5), ["C", "D"]):
    cell = ws.cell(row=tot_row, column=c,
                    value=f"=SUM({col_letter}{RES_FIRST}:{col_letter}{RES_LAST})")
    cell.number_format = "#,##0.00"
    cell.font = f_body_b
    cell.alignment = align_center
    cell.fill = fill_light
ws.cell(row=tot_row, column=5, value=f'=IF(C{tot_row}=0,0,1-D{tot_row}/C{tot_row})')
ws.cell(row=tot_row, column=5).number_format = "0.0%"
ws.cell(row=tot_row, column=5).font = f_body_b
ws.cell(row=tot_row, column=5).fill = fill_light
ws.cell(row=tot_row, column=5).alignment = align_center
for c in range(2, 9):
    ws.cell(row=tot_row, column=c).border = border_all
    ws.cell(row=tot_row, column=c).fill = fill_light
RES_TOTAL_ROW = tot_row

# ---- Evolución mensual de la deuda total ----
r1 = tot_row + 3
ws.cell(row=r1, column=2, value="EVOLUCIÓN DE LA DEUDA TOTAL (saldo acumulado al cierre de cada mes)")
ws.merge_cells(start_row=r1, start_column=2, end_row=r1, end_column=4)
ws.cell(row=r1, column=2).font = f_h2
for c in range(2, 5):
    ws.cell(row=r1, column=c).fill = fill_header2
ws.row_dimensions[r1].height = 20

hr2 = r1 + 1
ws.cell(row=hr2, column=2, value="Mes")
ws.cell(row=hr2, column=3, value="Saldo total")
style_header_row(ws, hr2, 2, 3, height=18)

MONTHS_ES = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
             "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]

MONTH_FIRST = hr2 + 1
for i, mname in enumerate(MONTHS_ES):
    rr = MONTH_FIRST + i
    mnum = i + 1
    end_expr = f'EOMONTH(DATE(Configuracion!$C${ROW_ANIO},{mnum},1),0)'
    ws.cell(row=rr, column=2, value=mname)
    ws.cell(row=rr, column=3,
            value=(f'=(SUM(Deudas!$C${FIRST_ROW}:$C${LAST_ROW})'
                    f'-SUMIFS({PAY_CAPITAL_RANGE},{PAY_FECHA_RANGE},"<="&{end_expr}))'
                    f'*Configuracion!$C${ROW_FACTOR}'))
    for c in (2, 3):
        cell = ws.cell(row=rr, column=c)
        cell.border = border_all
        cell.font = f_body
        if c == 3:
            cell.number_format = "#,##0.00"
            cell.alignment = align_center
MONTH_LAST = MONTH_FIRST + len(MONTHS_ES) - 1

lock_all(ws)

# =========================================================
# 7. DASHBOARD
# =========================================================
ws = wb.create_sheet("Dashboard")
set_col_widths(ws, [3, 16, 16, 16, 16, 16, 16, 16, 3])
ws.sheet_view.showGridLines = False

ws.merge_cells("B2:H2")
ws["B2"] = "DASHBOARD · CONTROL DE DEUDAS"
ws["B2"].font = f_h1
ws.row_dimensions[2].height = 28

ws.merge_cells("B3:H3")
ws["B3"] = '=" Moneda de visualización: " & Configuracion!$C$8'
ws["B3"].font = f_note

kpi_row = 5
card_defs = [
    ("DEUDA INICIAL TOTAL", f"=Configuracion!$C${ROW_SIMBOLO_ACTUAL}&\" \"&TEXT(Resumen!C{RES_TOTAL_ROW},\"#,##0.00\")", NAVY),
    ("SALDO ACTUAL TOTAL", f"=Configuracion!$C${ROW_SIMBOLO_ACTUAL}&\" \"&TEXT(Resumen!D{RES_TOTAL_ROW},\"#,##0.00\")", RED_NEG),
    ("TOTAL PAGADO", f"=Configuracion!$C${ROW_SIMBOLO_ACTUAL}&\" \"&TEXT(Resumen!C{RES_TOTAL_ROW}-Resumen!D{RES_TOTAL_ROW},\"#,##0.00\")", TEAL),
    ("% DEUDA PAGADA", f'=TEXT(Resumen!E{RES_TOTAL_ROW},"0.0%")', GOLD),
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
pie.title = "Saldo actual por deuda"
data_pie = Reference(wb["Resumen"], min_col=4, min_row=hr, max_row=RES_LAST)
cats_pie = Reference(wb["Resumen"], min_col=2, min_row=RES_FIRST, max_row=RES_LAST)
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
bar.title = "Saldo inicial vs. saldo actual por deuda"
bar.style = 10
bdata = Reference(wb["Resumen"], min_col=3, max_col=4, min_row=hr, max_row=RES_LAST)
bcats = Reference(wb["Resumen"], min_col=2, min_row=RES_FIRST, max_row=RES_LAST)
bar.add_data(bdata, titles_from_data=True)
bar.set_categories(bcats)
bar.height = 9
bar.width = 15
ws.add_chart(bar, f"E{chart_top}")

line = LineChart()
line.title = "Evolución de la deuda total"
ldata = Reference(wb["Resumen"], min_col=3, min_row=hr2, max_row=MONTH_LAST)
lcats = Reference(wb["Resumen"], min_col=2, min_row=MONTH_FIRST, max_row=MONTH_LAST)
line.add_data(ldata, titles_from_data=True)
line.set_categories(lcats)
line.height = 9
line.width = 30
ws.add_chart(line, f"B{chart_top + 19}")

lock_all(ws)

# =========================================================
wb.defined_names["TipoCambio"] = DefinedName("TipoCambio", attr_text=f"Configuracion!$C${ROW_TC}")
wb.defined_names["MonedaVisualizacion"] = DefinedName("MonedaVisualizacion", attr_text=f"Configuracion!$C${ROW_VIS}")
wb.defined_names["FactorConversion"] = DefinedName("FactorConversion", attr_text=f"Configuracion!$C${ROW_FACTOR}")

order = ["Portada", "Instrucciones", "Configuracion", "Deudas", "Pagos", "Resumen", "Dashboard"]
wb._sheets = [wb[name] for name in order]
tab_colors = {
    "Portada": NAVY, "Instrucciones": TEAL, "Configuracion": GOLD,
    "Deudas": TEAL, "Pagos": TEAL, "Resumen": NAVY, "Dashboard": NAVY,
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
    "productos", "03-control-deudas", "Control-de-Deudas.xlsx",
)
os.makedirs(os.path.dirname(out_path), exist_ok=True)
wb.save(out_path)
print("Saved:", out_path)
print("Key rows:", dict(ROW_FACTOR=ROW_FACTOR, DEBT_TOTAL_ROW=DEBT_TOTAL_ROW, RES_TOTAL_ROW=RES_TOTAL_ROW))
