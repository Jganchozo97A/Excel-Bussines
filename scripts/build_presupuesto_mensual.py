import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, Protection
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.chart import PieChart, BarChart, LineChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.workbook.defined_name import DefinedName
import os

# ---------- Palette / Fonts (same brand as product 1) ----------
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


MONTHS_ES = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
             "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
MONTHS_SHORT = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"]

categorias = [
    "Salario", "Freelance / Ingresos extra", "Otros ingresos",
    "Vivienda (alquiler/hipoteca)", "Alimentación", "Transporte",
    "Servicios (luz, agua, internet)", "Salud", "Educación",
    "Entretenimiento", "Ropa y cuidado personal", "Ahorro e inversión",
    "Deudas y préstamos", "Otros gastos",
]
N_INCOME = 3
N_CATS = len(categorias)

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
ws["B4"] = "PRESUPUESTO MENSUAL"
ws["B4"].font = Font(name=FONT_NAME, size=28, bold=True, color=NAVY)
ws["B4"].alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

ws.merge_cells("B7:F8")
ws["B7"] = ("Planifica tus ingresos y gastos mes a mes, compara lo presupuestado contra lo real "
            "y visualiza tus desviaciones al instante — con panel de control, gráficos automáticos "
            "y selector de moneda (local / USD).")
ws["B7"].font = Font(name=FONT_NAME, size=11, color=DARK_TEXT)
ws["B7"].alignment = align_wrap

for r in range(2, 9):
    ws.row_dimensions[r].height = 20
ws.row_dimensions[4].height = 30

features = [
    "Dashboard con indicadores clave (KPIs) y gráficos automáticos",
    "Fórmulas automáticas: nada que calcular a mano",
    "Gráficos: barras (presupuestado vs. real), tendencia anual y pastel",
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
     "Ve a la hoja 'Configuración' y define el símbolo de tu moneda local, el tipo de cambio, "
     "el mes que quieres analizar y si el resto de la plantilla se muestra en moneda LOCAL o USD."),
    ("2. Categorías",
     "En la hoja 'Categorías' puedes revisar o editar las categorías de ingreso/gasto que se usan "
     "en 'Presupuesto' y 'Real'."),
    ("3. Presupuesto",
     "En la hoja 'Presupuesto' escribe cuánto planeas ingresar/gastar en cada categoría, mes a mes. "
     "La columna 'Total anual' se calcula sola."),
    ("4. Real",
     "En la hoja 'Real' registra lo que efectivamente ingresaste/gastaste cada mes, en las mismas "
     "categorías y meses que en 'Presupuesto'."),
    ("5. Resumen automático",
     "La hoja 'Resumen' compara Presupuestado vs. Real para el mes elegido en 'Configuración', "
     "categoría por categoría, y además totaliza los 12 meses del año, como una tabla dinámica."),
    ("6. Dashboard",
     "La hoja 'Dashboard' muestra tus indicadores clave del mes elegido (presupuestado, real, "
     "disponible, % usado) y gráficos que se actualizan solos."),
    ("7. Cambiar de mes o de moneda",
     "En 'Configuración' cambia 'Mes seleccionado' para analizar otro mes, o 'Moneda de visualización' "
     "entre Local y USD: todo se recalcula automáticamente."),
    ("8. Celdas protegidas",
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

ROW_NOMBRE, ROW_SIMBOLO, ROW_TC, ROW_VIS, ROW_ANIO, ROW_MES = 5, 6, 7, 8, 9, 10
rows = [
    (ROW_NOMBRE, "Nombre de moneda local", "Soles (PEN)", "Solo referencia. Escribe el nombre de tu moneda."),
    (ROW_SIMBOLO, "Símbolo de moneda local", "S/", "Símbolo que se usará en el Dashboard cuando la vista esté en LOCAL."),
    (ROW_TC, "Tipo de cambio (1 USD =)", 3.75, "Cuántas unidades de tu moneda local equivalen a 1 USD."),
    (ROW_VIS, "Moneda de visualización", "Local", "Elige LOCAL o USD. Afecta Resumen y Dashboard."),
    (ROW_ANIO, "Año de análisis", 2026, "Año que se usa como referencia en Presupuesto y Real."),
    (ROW_MES, "Mes seleccionado", "Enero", "Mes que se muestra en Resumen y Dashboard."),
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

r_calc_label = ROW_MES + 1
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

ROW_MES_NUM = ROW_SIMBOLO_ACTUAL + 1
ws.cell(row=ROW_MES_NUM, column=2, value="Número de mes seleccionado").font = f_body_b
ws.cell(row=ROW_MES_NUM, column=3,
        value=f'=MATCH(C{ROW_MES},{{"Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"}},0)')
ws.cell(row=ROW_MES_NUM, column=3).border = border_all
ws.cell(row=ROW_MES_NUM, column=3).alignment = align_center
ws.cell(row=ROW_MES_NUM, column=4, value="1=Enero ... 12=Diciembre. Se usa para ubicar la columna del mes en Presupuesto y Real.").font = f_note
ws.row_dimensions[ROW_MES_NUM].height = 18

note_row = ROW_MES_NUM + 2
ws.merge_cells(f"B{note_row}:D{note_row}")
ws[f"B{note_row}"] = "Cambia 'Mes seleccionado' o 'Moneda de visualización' para recalcular automáticamente Resumen y Dashboard."
ws[f"B{note_row}"].font = f_note
ws[f"B{note_row}"].alignment = align_wrap

lock_all(ws)
unlock_range(ws, f"C{ROW_NOMBRE}:C{ROW_NOMBRE}")
unlock_range(ws, f"C{ROW_SIMBOLO}:C{ROW_SIMBOLO}")
unlock_range(ws, f"C{ROW_TC}:C{ROW_TC}")
unlock_range(ws, f"C{ROW_VIS}:C{ROW_VIS}")
unlock_range(ws, f"C{ROW_ANIO}:C{ROW_ANIO}")
unlock_range(ws, f"C{ROW_MES}:C{ROW_MES}")

dv_moneda = DataValidation(type="list", formula1='"Local,USD"', allow_blank=False)
ws.add_data_validation(dv_moneda)
dv_moneda.add(ws[f"C{ROW_VIS}"])

dv_mes = DataValidation(type="list", formula1='"' + ",".join(MONTHS_ES) + '"', allow_blank=False)
ws.add_data_validation(dv_mes)
dv_mes.add(ws[f"C{ROW_MES}"])

# =========================================================
# 4. CATEGORIAS
# =========================================================
ws = wb.create_sheet("Categorias")
set_col_widths(ws, [3, 32, 3])
ws.sheet_view.showGridLines = False

ws.merge_cells("B2:B2")
ws["B2"] = "CATEGORÍAS DE PRESUPUESTO"
ws["B2"].font = f_h1
ws.row_dimensions[2].height = 26

ws.cell(row=4, column=2, value="Categoría (ingreso/gasto)")
style_header_row(ws, 4, 2, 2)

CAT_FIRST_ROW = 5
CAT_LAST_ROW = CAT_FIRST_ROW + len(categorias) - 1
for i, cat in enumerate(categorias):
    rr = CAT_FIRST_ROW + i
    c = ws.cell(row=rr, column=2, value=cat)
    c.font = f_input
    c.fill = fill_input
    c.border = border_all
    ws.row_dimensions[rr].height = 16

note_row = CAT_LAST_ROW + 2
ws.merge_cells(f"B{note_row}:B{note_row}")
ws[f"B{note_row}"] = "Estas categorías deben coincidir, en el mismo orden, con las filas de 'Presupuesto' y 'Real'."
ws[f"B{note_row}"].font = f_note
ws[f"B{note_row}"].alignment = align_wrap

lock_all(ws)
unlock_range(ws, f"B{CAT_FIRST_ROW}:B{CAT_LAST_ROW}")

# =========================================================
# Helper to build the Presupuesto / Real grids
# =========================================================
def build_grid_sheet(sheet_name, title, example_income, example_expense):
    ws = wb.create_sheet(sheet_name)
    set_col_widths(ws, [30] + [10] * 12 + [13])
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "B5"

    ws.merge_cells("A2:N2")
    ws["A2"] = title
    ws["A2"].font = f_h1
    ws.row_dimensions[2].height = 26

    header_row = 4
    ws.cell(row=header_row, column=1, value="Categoría")
    for i, m in enumerate(MONTHS_SHORT):
        ws.cell(row=header_row, column=2 + i, value=m)
    ws.cell(row=header_row, column=14, value="Total anual")
    style_header_row(ws, header_row, 1, 14, height=20)

    first_row = header_row + 1
    last_row = first_row + N_CATS - 1

    for i, cat in enumerate(categorias):
        rr = first_row + i
        is_income = i < N_INCOME
        ws.cell(row=rr, column=1, value=f"=Categorias!$B${CAT_FIRST_ROW + i}")
        ws.cell(row=rr, column=1).font = f_body_b if is_income else f_body
        ws.cell(row=rr, column=1).border = border_all
        ws.cell(row=rr, column=1).fill = fill_light if is_income else fill_white

        example_row = example_income if is_income else example_expense
        example_vals = example_row.get(cat, [None] * 12) if example_row else [None] * 12

        for m in range(12):
            cell = ws.cell(row=rr, column=2 + m)
            val = example_vals[m] if m < len(example_vals) else None
            if val is not None:
                cell.value = val
            cell.font = f_input
            cell.fill = fill_input
            cell.border = border_all
            cell.number_format = "#,##0"
            cell.alignment = align_center

        tot = ws.cell(row=rr, column=14,
                       value=f"=SUM(B{rr}:M{rr})")
        tot.number_format = "#,##0"
        tot.font = f_body_b
        tot.fill = fill_light
        tot.border = border_all
        tot.alignment = align_center

    total_row = last_row + 1
    ws.cell(row=total_row, column=1, value="TOTAL").font = f_body_b
    ws.cell(row=total_row, column=1).border = border_all
    ws.cell(row=total_row, column=1).fill = fill_light
    for m in range(12):
        col = 2 + m
        letter = get_column_letter(col)
        c = ws.cell(row=total_row, column=col,
                     value=f"=SUM({letter}{first_row}:{letter}{last_row})")
        c.number_format = "#,##0"
        c.font = f_body_b
        c.fill = fill_light
        c.border = border_all
        c.alignment = align_center
    c = ws.cell(row=total_row, column=14, value=f"=SUM(N{first_row}:N{last_row})")
    c.number_format = "#,##0"
    c.font = f_body_b
    c.fill = fill_light
    c.border = border_all
    c.alignment = align_center

    dv_num = DataValidation(type="decimal", operator="greaterThanOrEqual", formula1="0",
                             allow_blank=True, showErrorMessage=True,
                             errorTitle="Monto inválido", error="Ingresa un número igual o mayor a 0.")
    ws.add_data_validation(dv_num)
    dv_num.add(f"B{first_row}:M{last_row}")

    lock_all(ws)
    unlock_range(ws, f"B{first_row}:M{last_row}")

    return ws, first_row, last_row, total_row


# Example data: (values roughly matching product 1's monthly totals for Jan-Mar, 0 afterwards)
presupuesto_income = {
    "Salario": [4500] * 12,
    "Freelance / Ingresos extra": [300] * 12,
    "Otros ingresos": [0] * 12,
}
presupuesto_expense = {
    "Vivienda (alquiler/hipoteca)": [1200] * 12,
    "Alimentación": [200] * 12,
    "Transporte": [100] * 12,
    "Servicios (luz, agua, internet)": [80] * 12,
    "Salud": [60] * 12,
    "Educación": [80] * 12,
    "Entretenimiento": [60] * 12,
    "Ropa y cuidado personal": [50] * 12,
    "Ahorro e inversión": [300] * 12,
    "Deudas y préstamos": [250] * 12,
    "Otros gastos": [50] * 12,
}

real_income = {
    "Salario": [4500, 4500, 4600] + [None] * 9,
    "Freelance / Ingresos extra": [350, 0, 0] + [None] * 9,
    "Otros ingresos": [0, 150, 0] + [None] * 9,
}
real_expense = {
    "Vivienda (alquiler/hipoteca)": [1200, 1200, 1200] + [None] * 9,
    "Alimentación": [180, 210, 195] + [None] * 9,
    "Transporte": [60, 120, 0] + [None] * 9,
    "Servicios (luz, agua, internet)": [75, 0, 0] + [None] * 9,
    "Salud": [90, 0, 0] + [None] * 9,
    "Educación": [0, 80, 0] + [None] * 9,
    "Entretenimiento": [45, 25, 0] + [None] * 9,
    "Ropa y cuidado personal": [0, 0, 0] + [None] * 9,
    "Ahorro e inversión": [300, 0, 0] + [None] * 9,
    "Deudas y préstamos": [0, 0, 250] + [None] * 9,
    "Otros gastos": [0, 0, 0] + [None] * 9,
}

ws_pres, PRES_FIRST, PRES_LAST, PRES_TOTAL_ROW = build_grid_sheet(
    "Presupuesto", "PRESUPUESTO ANUAL (por categoría y mes)", presupuesto_income, presupuesto_expense
)
ws_real, REAL_FIRST, REAL_LAST, REAL_TOTAL_ROW = build_grid_sheet(
    "Real", "GASTOS E INGRESOS REALES (por categoría y mes)", real_income, real_expense
)

# =========================================================
# 6. RESUMEN
# =========================================================
ws = wb.create_sheet("Resumen")
set_col_widths(ws, [3, 30, 16, 16, 16, 16, 3])
ws.sheet_view.showGridLines = False

ws.merge_cells("B2:F2")
ws["B2"] = "RESUMEN DINÁMICO (se actualiza solo)"
ws["B2"].font = f_h1
ws.row_dimensions[2].height = 26
ws.merge_cells("B3:F3")
ws["B3"] = f'=" Mes analizado: " & Configuracion!$C${ROW_MES} & "   |   Moneda: " & Configuracion!$C${ROW_VIS}'
ws["B3"].font = f_note

r0 = 5
ws.cell(row=r0, column=2, value="PRESUPUESTADO VS. REAL POR CATEGORÍA (mes seleccionado)")
ws.merge_cells(start_row=r0, start_column=2, end_row=r0, end_column=6)
ws.cell(row=r0, column=2).font = f_h2
for c in range(2, 7):
    ws.cell(row=r0, column=c).fill = fill_header2
ws.row_dimensions[r0].height = 20

hr = r0 + 1
cat_headers = ["Categoría", "Presupuestado", "Real", "Diferencia", "% usado"]
for i, h in enumerate(cat_headers):
    ws.cell(row=hr, column=2 + i, value=h)
style_header_row(ws, hr, 2, 6, height=18)

CAT_SUMMARY_FIRST = hr + 1
for i, cat in enumerate(categorias):
    rr = CAT_SUMMARY_FIRST + i
    pres_row = PRES_FIRST + i
    real_row = REAL_FIRST + i
    ws.cell(row=rr, column=2, value=f"=Categorias!$B${CAT_FIRST_ROW + i}")
    ws.cell(row=rr, column=3,
            value=(f'=INDEX(Presupuesto!$B${pres_row}:$M${pres_row},Configuracion!$C${ROW_MES_NUM})'
                   f'*Configuracion!$C${ROW_FACTOR}'))
    ws.cell(row=rr, column=4,
            value=(f'=INDEX(Real!$B${real_row}:$M${real_row},Configuracion!$C${ROW_MES_NUM})'
                   f'*Configuracion!$C${ROW_FACTOR}'))
    ws.cell(row=rr, column=5, value=f"=C{rr}-D{rr}")
    ws.cell(row=rr, column=6, value=f'=IF(C{rr}=0,0,D{rr}/C{rr})')
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
CAT_SUMMARY_LAST = CAT_SUMMARY_FIRST + N_CATS - 1
INCOME_ROW_FIRST = CAT_SUMMARY_FIRST
INCOME_ROW_LAST = CAT_SUMMARY_FIRST + N_INCOME - 1
EXPENSE_ROW_FIRST = INCOME_ROW_LAST + 1

totals_spec = [
    ("TOTAL INGRESOS", INCOME_ROW_FIRST, INCOME_ROW_LAST),
    ("TOTAL GASTOS", EXPENSE_ROW_FIRST, CAT_SUMMARY_LAST),
]
total_rows = {}
tot_row = CAT_SUMMARY_LAST
for label, r_first, r_last in totals_spec:
    tot_row += 1
    ws.cell(row=tot_row, column=2, value=label).font = f_body_b
    for c, col_letter in zip(range(3, 6), ["C", "D", "E"]):
        cell = ws.cell(row=tot_row, column=c,
                        value=f"=SUM({col_letter}{r_first}:{col_letter}{r_last})")
        cell.number_format = "#,##0.00"
        cell.font = f_body_b
        cell.alignment = align_center
        cell.fill = fill_light
    ws.cell(row=tot_row, column=6, value=f'=IF(C{tot_row}=0,0,D{tot_row}/C{tot_row})')
    ws.cell(row=tot_row, column=6).number_format = "0.0%"
    ws.cell(row=tot_row, column=6).font = f_body_b
    ws.cell(row=tot_row, column=6).fill = fill_light
    ws.cell(row=tot_row, column=6).alignment = align_center
    for c in range(2, 7):
        ws.cell(row=tot_row, column=c).border = border_all
        ws.cell(row=tot_row, column=c).fill = fill_light
    total_rows[label] = tot_row

TOTAL_INGRESOS_ROW = total_rows["TOTAL INGRESOS"]
TOTAL_GASTOS_ROW = total_rows["TOTAL GASTOS"]

tot_row += 1
ws.cell(row=tot_row, column=2, value="BALANCE (Ingresos - Gastos)").font = f_body_b
for c, col_letter in zip(range(3, 6), ["C", "D", "E"]):
    cell = ws.cell(row=tot_row, column=c,
                    value=f"={col_letter}{TOTAL_INGRESOS_ROW}-{col_letter}{TOTAL_GASTOS_ROW}")
    cell.number_format = "#,##0.00"
    cell.font = f_body_b
    cell.alignment = align_center
    cell.fill = fill_light
for c in range(2, 7):
    ws.cell(row=tot_row, column=c).border = border_all
    ws.cell(row=tot_row, column=c).fill = fill_light
BALANCE_ROW = tot_row
CAT_TOTAL_ROW = TOTAL_GASTOS_ROW  # kept for backward-compat naming; expense total drives % usado

r1 = tot_row + 3
ws.cell(row=r1, column=2, value="EVOLUCIÓN MENSUAL (todo el año)")
ws.merge_cells(start_row=r1, start_column=2, end_row=r1, end_column=5)
ws.cell(row=r1, column=2).font = f_h2
for c in range(2, 6):
    ws.cell(row=r1, column=c).fill = fill_header2
ws.row_dimensions[r1].height = 20

hr2 = r1 + 1
mon_headers = ["Mes", "Presupuestado", "Real", "Diferencia"]
for i, h in enumerate(mon_headers):
    ws.cell(row=hr2, column=2 + i, value=h)
style_header_row(ws, hr2, 2, 5, height=18)

MONTH_SUMMARY_FIRST = hr2 + 1
for i, mname in enumerate(MONTHS_ES):
    rr = MONTH_SUMMARY_FIRST + i
    col_letter = get_column_letter(2 + i)  # matches Presupuesto/Real month columns (B..M)
    ws.cell(row=rr, column=2, value=mname)
    ws.cell(row=rr, column=3,
            value=f'=Presupuesto!{col_letter}${PRES_TOTAL_ROW}*Configuracion!$C${ROW_FACTOR}')
    ws.cell(row=rr, column=4,
            value=f'=Real!{col_letter}${REAL_TOTAL_ROW}*Configuracion!$C${ROW_FACTOR}')
    ws.cell(row=rr, column=5, value=f"=C{rr}-D{rr}")
    for c in range(2, 6):
        cell = ws.cell(row=rr, column=c)
        cell.border = border_all
        cell.font = f_body
        if c in (3, 4, 5):
            cell.number_format = "#,##0.00"
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
ws["B2"] = "DASHBOARD · PRESUPUESTO MENSUAL"
ws["B2"].font = f_h1
ws.row_dimensions[2].height = 28

ws.merge_cells("B3:H3")
ws["B3"] = (f'=" Mes: " & Configuracion!$C${ROW_MES} & "   |   Moneda: " & Configuracion!$C${ROW_VIS} '
            f'& "   |   Año: " & Configuracion!$C${ROW_ANIO}')
ws["B3"].font = f_note

kpi_row = 5
card_defs = [
    ("INGRESO REAL (MES)", f"=Configuracion!$C${ROW_SIMBOLO_ACTUAL}&\" \"&TEXT(Resumen!D{TOTAL_INGRESOS_ROW},\"#,##0.00\")", TEAL),
    ("GASTO PRESUPUESTADO (MES)", f"=Configuracion!$C${ROW_SIMBOLO_ACTUAL}&\" \"&TEXT(Resumen!C{TOTAL_GASTOS_ROW},\"#,##0.00\")", GOLD),
    ("GASTO REAL (MES)", f"=Configuracion!$C${ROW_SIMBOLO_ACTUAL}&\" \"&TEXT(Resumen!D{TOTAL_GASTOS_ROW},\"#,##0.00\")", RED_NEG),
    ("% DE GASTOS USADO", f'=IF(Resumen!C{TOTAL_GASTOS_ROW}=0,"0.0%",TEXT(Resumen!D{TOTAL_GASTOS_ROW}/Resumen!C{TOTAL_GASTOS_ROW},"0.0%"))', NAVY),
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

EXPENSE_CAT_FIRST = CAT_SUMMARY_FIRST + N_INCOME

pie = PieChart()
pie.title = "Gastos reales por categoría (mes)"
data_expense = Reference(wb["Resumen"], min_col=4, min_row=EXPENSE_CAT_FIRST, max_row=CAT_SUMMARY_LAST)
cats_expense = Reference(wb["Resumen"], min_col=2, min_row=EXPENSE_CAT_FIRST, max_row=CAT_SUMMARY_LAST)
pie.add_data(data_expense, titles_from_data=False)
pie.set_categories(cats_expense)
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
bar.title = "Presupuestado vs. Real por categoría (mes)"
bar.style = 10
bdata = Reference(wb["Resumen"], min_col=3, max_col=4, min_row=hr, max_row=CAT_SUMMARY_LAST)
bcats = Reference(wb["Resumen"], min_col=2, min_row=CAT_SUMMARY_FIRST, max_row=CAT_SUMMARY_LAST)
bar.add_data(bdata, titles_from_data=True)
bar.set_categories(bcats)
bar.height = 9
bar.width = 15
ws.add_chart(bar, f"E{chart_top}")

line = LineChart()
line.title = "Evolución anual: presupuestado vs. real"
ldata = Reference(wb["Resumen"], min_col=3, max_col=4, min_row=hr2, max_row=MONTH_SUMMARY_LAST)
lcats = Reference(wb["Resumen"], min_col=2, min_row=MONTH_SUMMARY_FIRST, max_row=MONTH_SUMMARY_LAST)
line.add_data(ldata, titles_from_data=True)
line.set_categories(lcats)
line.height = 9
line.width = 30
ws.add_chart(line, f"B{chart_top + 19}")

lock_all(ws)

# =========================================================
# Defined names
# =========================================================
wb.defined_names["TipoCambio"] = DefinedName("TipoCambio", attr_text=f"Configuracion!$C${ROW_TC}")
wb.defined_names["MonedaVisualizacion"] = DefinedName("MonedaVisualizacion", attr_text=f"Configuracion!$C${ROW_VIS}")
wb.defined_names["FactorConversion"] = DefinedName("FactorConversion", attr_text=f"Configuracion!$C${ROW_FACTOR}")
wb.defined_names["MesSeleccionado"] = DefinedName("MesSeleccionado", attr_text=f"Configuracion!$C${ROW_MES}")

order = ["Portada", "Instrucciones", "Configuracion", "Categorias", "Presupuesto", "Real", "Resumen", "Dashboard"]
wb._sheets = [wb[name] for name in order]
tab_colors = {
    "Portada": NAVY, "Instrucciones": TEAL, "Configuracion": GOLD, "Categorias": GOLD,
    "Presupuesto": TEAL, "Real": TEAL, "Resumen": NAVY, "Dashboard": NAVY,
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
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))) if "__file__" in dir() else ".",
    "productos", "02-presupuesto-mensual", "Presupuesto-Mensual.xlsx",
)
os.makedirs(os.path.dirname(out_path), exist_ok=True)
wb.save(out_path)
print("Saved:", out_path)
print("Key rows:", dict(ROW_FACTOR=ROW_FACTOR, ROW_MES_NUM=ROW_MES_NUM, CAT_TOTAL_ROW=CAT_TOTAL_ROW))
