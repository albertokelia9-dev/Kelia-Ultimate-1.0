# ===========================
# ALBERTO KELIA — Hevy-Style Hypertrophy Tracker (Excel)
# Google Colab — Copy/Paste & Run
# ===========================

# --- (1) Install deps (Colab usually has them, but we make it bulletproof)
!pip -q install xlsxwriter pillow

from google.colab import files
import os
from PIL import Image
import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell, xl_range

# --- (2) Upload logo (PNG recommended)
uploaded = files.upload()
if not uploaded:
    raise RuntimeError("No subiste el logo. Sube un PNG y vuelve a ejecutar.")
LOGO_PATH = list(uploaded.keys())[0]

# --- (3) Output paths
OUT_PATH = "ALBERTO_KELIA_Hypertrophy_Tracker_HevyStyle.xlsx"
WATERMARK_PATH = "AK_logo_watermark_10pct.png"

# ===========================
# THEME (Ultimate Dark Mode)
# ===========================
BG_BLACK   = "#000000"
BG_DARK    = "#121212"
BG_INPUT   = "#1E1E1E"
TXT_WHITE  = "#FFFFFF"
GOLD       = "#C5A059"
GREY_BORDER= "#2F2F2F"
RED_DARK   = "#3A0000"

# ===========================
# Build watermark @ ~10% opacity
# ===========================
img = Image.open(LOGO_PATH).convert("RGBA")
r, g, b, a = img.split()
a = a.point(lambda p: int(p * 0.10))
Image.merge("RGBA", (r, g, b, a)).save(WATERMARK_PATH)

# ===========================
# Workbook
# ===========================
workbook = xlsxwriter.Workbook(OUT_PATH)

def with_gold_separator(d: dict):
    # Thick gold bottom border to separate exercises
    return {**d, "bottom": 2, "bottom_color": GOLD}

# ---------------------------
# FORMAT DICTS
# ---------------------------
D_BG = {"bg_color": BG_BLACK, "font_name": "Segoe UI", "font_color": TXT_WHITE}

D_TITLE = {
    "bg_color": BG_BLACK, "font_name": "Segoe UI", "font_color": GOLD,
    "bold": True, "font_size": 22, "align": "left", "valign": "vcenter"
}

D_HEADER = {
    "bg_color": BG_DARK, "font_name": "Segoe UI", "font_color": TXT_WHITE,
    "bold": True, "font_size": 11, "align": "center", "valign": "vcenter",
    "border": 2, "border_color": GOLD
}

D_GROUP = {
    "bg_color": BG_BLACK, "font_name": "Segoe UI", "font_color": GOLD,
    "bold": True, "font_size": 11, "align": "center", "valign": "vcenter",
    "border": 2, "border_color": GOLD
}

D_SUB = {
    "bg_color": BG_DARK, "font_name": "Segoe UI", "font_color": TXT_WHITE,
    "bold": True, "font_size": 10, "align": "center", "valign": "vcenter",
    "border": 1, "border_color": GOLD
}

D_INPUT = {
    "bg_color": BG_INPUT, "font_name": "Segoe UI", "font_color": TXT_WHITE,
    "border": 1, "border_color": GREY_BORDER, "valign": "vcenter"
}
D_INPUT_CENTER = {**D_INPUT, "align": "center"}

D_WEIGHT = {**D_INPUT_CENTER, "num_format": "0.0"}
D_REPS   = {**D_INPUT_CENTER, "num_format": "0"}
D_RIR    = {**D_INPUT_CENTER}

D_FORMULA = {
    "bg_color": BG_DARK, "font_name": "Segoe UI", "font_color": TXT_WHITE,
    "border": 1, "border_color": GREY_BORDER, "align": "center", "valign": "vcenter",
    "num_format": "0"
}

D_TOTAL_LABEL = {
    "bg_color": BG_BLACK, "font_name": "Segoe UI", "font_color": GOLD,
    "bold": True, "align": "right", "valign": "vcenter",
    "border": 2, "border_color": GOLD
}

D_TOTAL_VALUE = {
    "bg_color": BG_BLACK, "font_name": "Segoe UI", "font_color": TXT_WHITE,
    "bold": True, "align": "center", "valign": "vcenter",
    "num_format": "0",
    "border": 2, "border_color": GOLD
}

D_RIR_ALERT = {
    "bg_color": RED_DARK, "font_name": "Segoe UI", "font_color": GOLD,
    "bold": True, "align": "center", "valign": "vcenter",
    "border": 1, "border_color": GREY_BORDER
}

D_BTN = {
    "bg_color": GOLD, "font_name": "Segoe UI", "font_color": BG_BLACK,
    "bold": True, "align": "center", "valign": "vcenter",
    "border": 2, "border_color": GOLD
}

D_HINT = {
    "bg_color": BG_BLACK, "font_name": "Segoe UI", "font_color": TXT_WHITE,
    "align": "left", "valign": "top"
}

# ---------------------------
# CREATE FORMATS
# ---------------------------
fmt_bg        = workbook.add_format(D_BG)
fmt_title     = workbook.add_format(D_TITLE)
fmt_header    = workbook.add_format(D_HEADER)
fmt_group     = workbook.add_format(D_GROUP)
fmt_subheader = workbook.add_format(D_SUB)

fmt_input     = workbook.add_format(D_INPUT)
fmt_input_c   = workbook.add_format(D_INPUT_CENTER)
fmt_weight    = workbook.add_format(D_WEIGHT)
fmt_reps      = workbook.add_format(D_REPS)
fmt_rir       = workbook.add_format(D_RIR)
fmt_formula   = workbook.add_format(D_FORMULA)

fmt_input_sep   = workbook.add_format(with_gold_separator(D_INPUT))
fmt_weight_sep  = workbook.add_format(with_gold_separator(D_WEIGHT))
fmt_reps_sep    = workbook.add_format(with_gold_separator(D_REPS))
fmt_rir_sep     = workbook.add_format(with_gold_separator(D_RIR))
fmt_formula_sep = workbook.add_format(with_gold_separator(D_FORMULA))

fmt_total_label = workbook.add_format(D_TOTAL_LABEL)
fmt_total_value = workbook.add_format(D_TOTAL_VALUE)

fmt_rir_alert = workbook.add_format(D_RIR_ALERT)

fmt_btn  = workbook.add_format(D_BTN)
fmt_hint = workbook.add_format(D_HINT)

# ===========================
# Column layout (Día sheets)
# ===========================
COLS = [
    ("Ejercicio", 40),
    ("Notas Técnicas / Setup", 34),

    ("S1 Peso", 8), ("S1 Reps", 6), ("S1 RIR", 6), ("S1 Vol", 10),
    ("S2 Peso", 8), ("S2 Reps", 6), ("S2 RIR", 6), ("S2 Vol", 10),
    ("S3 Peso", 8), ("S3 Reps", 6), ("S3 RIR", 6), ("S3 Vol", 10),
    ("S4 Peso", 8), ("S4 Reps", 6), ("S4 RIR", 6), ("S4 Vol", 10),

    ("TOTAL VOLUMEN", 14),
    ("Feedback Post-Ejercicio", 34),
]

# ===========================
# TRAINING SHEET BUILDER
# - Adds PR Tracker area
# - Adds day total cell for Dashboard
# ===========================
def build_training_sheet(name: str, num_exercises: int = 28):
    ws = workbook.add_worksheet(name)
    ws.hide_gridlines(2)
    ws.set_zoom(110)
    ws.set_default_row(22)
    ws.set_tab_color(GOLD)

    # Set widths
    for c, (_, w) in enumerate(COLS):
        ws.set_column(c, c, w)

    # Rows 1-5 empty (black) for manual logo insertion
    for r in range(0, 5):
        ws.set_row(r, 24, fmt_bg)
        ws.merge_range(r, 0, r, len(COLS) - 1, "", fmt_bg)

    group_row, sub_row, data_start = 5, 6, 7
    ws.set_row(group_row, 26)
    ws.set_row(sub_row, 22)

    # Header merges
    ws.merge_range(group_row, 0, sub_row, 0, "EJERCICIO", fmt_header)
    ws.merge_range(group_row, 1, sub_row, 1, "NOTAS TÉCNICAS / SETUP", fmt_header)
    ws.merge_range(group_row, 18, sub_row, 18, "TOTAL VOLUMEN", fmt_header)
    ws.merge_range(group_row, 19, sub_row, 19, "FEEDBACK POST-EJERCICIO", fmt_header)

    # Set groups
    ws.merge_range(group_row, 2, group_row, 5, "SET 1", fmt_group)
    ws.merge_range(group_row, 6, group_row, 9, "SET 2", fmt_group)
    ws.merge_range(group_row, 10, group_row, 13, "SET 3", fmt_group)
    ws.merge_range(group_row, 14, group_row, 17, "SET 4", fmt_group)

    # Subheaders
    for base_col in [2, 6, 10, 14]:
        ws.write(sub_row, base_col + 0, "Peso", fmt_subheader)
        ws.write(sub_row, base_col + 1, "Reps", fmt_subheader)
        ws.write(sub_row, base_col + 2, "RIR", fmt_subheader)
        ws.write(sub_row, base_col + 3, "Vol", fmt_subheader)

    # UX: freeze & filter
    ws.freeze_panes(data_start, 2)
    ws.autofilter(sub_row, 0, sub_row, len(COLS) - 1)

    # Watermark
    ws.insert_image(10, 6, WATERMARK_PATH, {
        "x_scale": 1.6, "y_scale": 1.6,
        "x_offset": 10, "y_offset": 0,
        "object_position": 2
    })

    # Column indices
    weight_cols = [2, 6, 10, 14]
    reps_cols   = [3, 7, 11, 15]
    rir_cols    = [4, 8, 12, 16]
    vol_cols    = [5, 9, 13, 17]

    # Main table rows
    for r in range(data_start, data_start + num_exercises):
        ws.write_blank(r, 0, None, fmt_input_sep)   # Exercise
        ws.write_blank(r, 1, None, fmt_input_sep)   # Notes

        for s in range(4):
            wcol, repcol, rircol, vcol = weight_cols[s], reps_cols[s], rir_cols[s], vol_cols[s]
            ws.write_blank(r, wcol,   None, fmt_weight_sep)
            ws.write_blank(r, repcol, None, fmt_reps_sep)
            ws.write_blank(r, rircol, None, fmt_rir_sep)

            wcell = xl_rowcol_to_cell(r, wcol)
            rcell = xl_rowcol_to_cell(r, repcol)
            ws.write_formula(
                r, vcol,
                f'=IF(OR({wcell}="",{rcell}=""),"",{wcell}*{rcell})',
                fmt_formula_sep
            )

        v1 = xl_rowcol_to_cell(r, vol_cols[0])
        v2 = xl_rowcol_to_cell(r, vol_cols[1])
        v3 = xl_rowcol_to_cell(r, vol_cols[2])
        v4 = xl_rowcol_to_cell(r, vol_cols[3])
        ws.write_formula(r, 18, f"=SUM({v1},{v2},{v3},{v4})", fmt_formula_sep)  # Total volume per exercise
        ws.write_blank(r, 19, None, fmt_input_sep)  # Feedback

    # RIR dropdown + conditional format (RIR = 0 or Fallo)
    rir_list = ["0", "1", "2", "3", "4", "Fallo"]
    for c in rir_cols:
        ws.data_validation(data_start, c, data_start + num_exercises - 1, c, {
            "validate": "list",
            "source": rir_list,
            "input_title": "RIR",
            "input_message": "Selecciona 0–4 o 'Fallo'",
            "error_title": "Valor inválido",
            "error_message": "Usa 0, 1, 2, 3, 4 o Fallo."
        })
        ws.conditional_format(data_start, c, data_start + num_exercises - 1, c, {
            "type": "cell", "criteria": "==", "value": 0, "format": fmt_rir_alert
        })
        ws.conditional_format(data_start, c, data_start + num_exercises - 1, c, {
            "type": "cell", "criteria": "==", "value": '"Fallo"', "format": fmt_rir_alert
        })

    # Total day row
    total_row = data_start + num_exercises + 1
    ws.set_row(total_row, 26)
    ws.write(total_row, 17, "TOTAL DÍA", fmt_total_label)
    total_range = xl_range(data_start, 18, data_start + num_exercises - 1, 18)
    ws.write_formula(total_row, 18, f"=SUM({total_range})", fmt_total_value)

    # ===========================
    # PR TRACKER (Hevy-like)
    # ===========================
    # Placement: starts a few rows below total
    pr_start = total_row + 3
    ws.set_row(pr_start, 24)
    ws.merge_range(pr_start, 0, pr_start, 8, "PR TRACKER (Top por Ejercicio)", fmt_header)

    # PR headers
    pr_headers = ["Ejercicio", "Top Weight", "Top Reps", "Top Set Vol", "PR Vol Anterior", "¿Nuevo PR?", "Fecha"]
    pr_cols = [0, 1, 2, 3, 4, 5, 6]
    for i, h in enumerate(pr_headers):
        ws.write(pr_start + 1, pr_cols[i], h, fmt_subheader)

    # Set widths for PR area
    ws.set_column(0, 0, 40)  # already
    ws.set_column(1, 6, 14)

    # Fill PR rows matching each exercise row (same count)
    # Simple + robust formulas:
    # Top Weight = MAX of set weights in row
    # Top Reps   = MAX of reps in row (simple)
    # Top Set Vol= MAX of set vols in row
    # "¿Nuevo PR?" compares Top Set Vol to input PR Vol Anterior
    # Conditional formatting marks new PR in gold background? (we do gold border + bold)
    for i in range(num_exercises):
        src_r = data_start + i
        dst_r = pr_start + 2 + i

        # Link exercise name to main table
        ex_cell = xl_rowcol_to_cell(src_r, 0)
        ws.write_formula(dst_r, 0, f"={ex_cell}", fmt_input_sep)

        # Cells of sets in that row
        w_cells = [xl_rowcol_to_cell(src_r, c) for c in weight_cols]
        r_cells = [xl_rowcol_to_cell(src_r, c) for c in reps_cols]
        v_cells = [xl_rowcol_to_cell(src_r, c) for c in vol_cols]

        ws.write_formula(dst_r, 1, f"=MAX({','.join(w_cells)})", fmt_formula_sep)
        ws.write_formula(dst_r, 2, f"=MAX({','.join(r_cells)})", fmt_formula_sep)
        ws.write_formula(dst_r, 3, f"=MAX({','.join(v_cells)})", fmt_formula_sep)

        # Previous PR Vol (input)
        ws.write_blank(dst_r, 4, None, fmt_input_sep)

        # New PR?
        topvol = xl_rowcol_to_cell(dst_r, 3)
        prevpr = xl_rowcol_to_cell(dst_r, 4)
        ws.write_formula(dst_r, 5, f'=IF(OR({topvol}="",{prevpr}=""),"",IF({topvol}>{prevpr},"PR",""))', fmt_formula_sep)

        # Date (manual input)
        ws.write_blank(dst_r, 6, None, fmt_input_sep)

    # Conditional format: if "¿Nuevo PR?" == "PR" -> highlight
    ws.conditional_format(pr_start + 2, 5, pr_start + 2 + num_exercises - 1, 5, {
        "type": "cell",
        "criteria": "==",
        "value": '"PR"',
        "format": workbook.add_format({
            "bg_color": BG_DARK,
            "font_name": "Segoe UI",
            "font_color": GOLD,
            "bold": True,
            "align": "center",
            "valign": "vcenter",
            "border": 2,
            "border_color": GOLD
        })
    })

    return ws, (total_row, 18)

# ===========================
# Build Day sheets
# ===========================
day_total_cells = {}
for i in range(1, 6):
    sheet_name = f"Día {i}"
    _, (tr, tc) = build_training_sheet(sheet_name, num_exercises=28)
    day_total_cells[sheet_name] = (tr, tc)

# ===========================
# DASHBOARD
# - Tonelage per day
# - Units dropdown (kg/lb)
# - 12-week history + line chart
# - "Button" cell
# ===========================
dash = workbook.add_worksheet("DASHBOARD")
dash.hide_gridlines(2)
dash.set_zoom(120)
dash.set_default_row(22)
dash.set_tab_color(GOLD)

dash.set_column(0, 0, 18)
dash.set_column(1, 1, 22)
dash.set_column(2, 2, 18)
dash.set_column(3, 3, 18)
dash.set_column(4, 8, 14)

dash.merge_range(0, 0, 1, 8, "ALBERTO KELIA — DASHBOARD", fmt_title)

# Units selector
dash.write(3, 0, "UNIDADES", fmt_header)
dash.write(3, 1, "kg / lb", fmt_header)
dash.write(4, 0, "Selecciona:", fmt_input)
dash.write(4, 1, "kg", fmt_input_c)

dash.data_validation(4, 1, 4, 1, {
    "validate": "list",
    "source": ["kg", "lb"],
    "input_title": "Unidades",
    "input_message": "Selecciona kg o lb"
})

# Optional: display conversion factor (only for display — no auto conversion of inputs)
dash.write(4, 2, "Factor (kg→lb)", fmt_input)
dash.write_formula(4, 3, '=IF($B$5="lb",2.20462,1)', fmt_formula)  # Excel row 5

# Day summary table
dash.write(6, 0, "DÍA", fmt_header)
dash.write(6, 1, "SESIÓN", fmt_header)
dash.write(6, 2, "TONELAJE", fmt_header)
dash.write(6, 3, "TONELAJE (Display)", fmt_header)

for idx in range(5):
    day = f"Día {idx+1}"
    r = 7 + idx
    dash.write(r, 0, idx + 1, fmt_input_c)
    dash.write(r, 1, day, fmt_input)
    tr, tc = day_total_cells[day]
    abs_cell = xl_rowcol_to_cell(tr, tc, row_abs=True, col_abs=True)
    dash.write_formula(r, 2, f"='{day}'!{abs_cell}", fmt_formula)
    # Display converted (if user wants to "see" lb)
    dash.write_formula(r, 3, f"=C{r+1}*$D$5", fmt_formula)

# Weekly total
dash.write(13, 1, "TOTAL SEMANAL", fmt_total_label)
dash.write_formula(13, 2, "=SUM(C8:C12)", fmt_total_value)
dash.write_formula(13, 3, "=SUM(D8:D12)", fmt_total_value)

# Watermark on dashboard
dash.insert_image(7, 6, WATERMARK_PATH, {"x_scale": 0.9, "y_scale": 0.9, "object_position": 2})

# "Button" (cell-style) — New Mesocycle
dash.merge_range(15, 0, 16, 3, "NUEVO MESOCICLO (Duplicar plantilla)", fmt_btn)
dash.merge_range(
    17, 0, 20, 8,
    "INSTRUCCIONES:\n"
    "1) Para iniciar un nuevo mesociclo, vuelve a ejecutar este notebook.\n"
    "2) Renombra el archivo exportado con la fecha (ej: Mesociclo_YYYY-MM-DD.xlsx).\n"
    "3) (Opcional) Puedes guardar en Drive y llevar histórico por semanas en la tabla inferior.",
    fmt_hint
)

# ===========================
# 12-WEEK HISTORY + CHART
# ===========================
hist_title_row = 22
dash.merge_range(hist_title_row, 0, hist_title_row, 3, "HISTÓRICO — 12 SEMANAS (Tonelaje Semanal)", fmt_header)

dash.write(hist_title_row + 1, 0, "Semana", fmt_subheader)
dash.write(hist_title_row + 1, 1, "Tonelaje", fmt_subheader)
dash.write(hist_title_row + 1, 2, "Tonelaje (Display)", fmt_subheader)

# Week 1 auto = current weekly total (C14)
dash.write(hist_title_row + 2, 0, "Semana 1 (Actual)", fmt_input)
dash.write_formula(hist_title_row + 2, 1, "=C14", fmt_formula)
dash.write_formula(hist_title_row + 2, 2, "=D14", fmt_formula)

# Weeks 2-12 manual entry (but auto display conversion)
for i in range(2, 13):
    rr = hist_title_row + 1 + i
    dash.write(rr, 0, f"Semana {i}", fmt_input)
    dash.write_blank(rr, 1, None, fmt_input_c)        # user input weekly tonnage
    dash.write_formula(rr, 2, f"=B{rr+1}*$D$5", fmt_formula)

# Chart (line) for 12 weeks
chart = workbook.add_chart({"type": "line"})
chart.set_style(10)

# categories: Week labels
cat_range = f"=DASHBOARD!$A${hist_title_row+3}:$A${hist_title_row+14}"
val_range = f"=DASHBOARD!$C${hist_title_row+3}:$C${hist_title_row+14}"  # display series
chart.add_series({
    "name": "Tonelaje (Display)",
    "categories": cat_range,
    "values": val_range,
    "line": {"color": GOLD},
})
chart.set_title({"name": "Tendencia 12 semanas"})
chart.set_x_axis({"name": "Semanas"})
chart.set_y_axis({"name": "Tonelaje"})
chart.set_legend({"position": "bottom"})

# Insert chart
dash.insert_chart(hist_title_row + 2, 4, chart, {"x_scale": 1.25, "y_scale": 1.25})

# ===========================
# Finalize
# ===========================
workbook.close()

print("✅ Exportado:", OUT_PATH)

# --- (4) Download file
files.download(OUT_PATH)
