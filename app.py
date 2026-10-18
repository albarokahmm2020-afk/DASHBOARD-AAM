from pathlib import Path
import io
import base64
import json
import hashlib
import inspect
import html
import math
import os
import pandas as pd
import requests
import streamlit as st
import altair as alt

st.session_state["_drive_copy_button_counter"] = 0

_streamlit_dataframe = st.dataframe
_streamlit_data_editor = st.data_editor

def _all_rows_table_kwargs(data, kwargs, row_height):
    source = data.data if not isinstance(data, pd.DataFrame) and isinstance(getattr(data, "data", None), pd.DataFrame) else data
    try:
        row_count = len(source)
    except (TypeError, ValueError):
        return kwargs
    # Expand the component to all rows; the page itself remains the scroll surface.
    kwargs["height"] = max(60, 30 + row_count * (row_height + 2))
    kwargs["row_height"] = row_height
    return kwargs

def _performance_value(value):
    """Convert displayed/numeric A/T and deviation values into comparable floats."""
    if pd.isna(value):
        return None
    if isinstance(value, str):
        cleaned = value.strip().replace("%", "").replace(",", "").replace(" ", "")
        if cleaned in {"", "—", "-", "None", "nan"}:
            return None
        try:
            return float(cleaned)
        except ValueError:
            return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None

def _color_performance_columns(table):
    """Highlight achieved A/T in green and negative deviations in red."""
    is_dataframe = isinstance(table, pd.DataFrame)
    is_styler = not is_dataframe and isinstance(getattr(table, "data", None), pd.DataFrame) and hasattr(table, "map")
    if not is_dataframe and not is_styler:
        return table
    source = table.data if is_styler else table
    at_columns = [
        column for column in source.columns
        if str(column).strip().casefold().startswith("a/t")
        or str(column).strip().casefold().startswith("achievement")
        or str(column).strip().casefold().startswith("pencapaian")
    ]
    deviation_columns = [
        column for column in source.columns
        if any(token in str(column).strip().casefold() for token in ("dev", "deviation", "gap"))
    ]
    styled = table if is_styler else table.style
    row_backgrounds = ["#F4F7FB" if index % 2 == 0 else "#E8EEF6" for index in range(len(source))]
    styled = styled.set_table_styles([
        {"selector": "table", "props": [("border-collapse", "collapse"), ("width", "100%"), ("border", "1px solid #C7D4E5")]},
        {"selector": "th", "props": [("background", "linear-gradient(180deg,#335A9F 0%,#24477F 100%)"), ("color", "#FFFFFF"), ("font-weight", "800"), ("font-size", ".82rem"), ("line-height", "1.1"), ("border", "1px solid #1D3B70"), ("padding", "3px 3px")]},
        {"selector": "td", "props": [("font-size", ".82rem"), ("line-height", "1.1"), ("border", "1px solid #D4DEEA"), ("color", "#1B2E46"), ("padding", "2px 3px")]},
    ], overwrite=False)
    styled = styled.set_properties(**{"background-color": "#F4F7FB", "color": "#1B2E46"})
    styled = styled.apply(lambda column: [f"background-color:{row_backgrounds[index]}" for index in range(len(column))], axis=0)
    def total_row_style(row):
        is_total = any(
            isinstance(value, str) and value.strip().casefold().startswith(("total", "grand total", "subtotal"))
            for value in row.tolist()
        )
        if is_total:
            return ["font-weight:900;border-top:2px solid #335A9F;background-color:#DCE6F3;color:#1B2E46"] * len(row)
        return [""] * len(row)
    styled = styled.apply(total_row_style, axis=1)
    for column in at_columns:
        def at_style(value):
            number = _performance_value(value)
            if number is not None and number >= 100:
                return "background-color:#D7F2E2;color:#146C3E;font-weight:800"
            return ""
        styled = styled.map(at_style, subset=[column])
    for column in deviation_columns:
        def deviation_style(value):
            number = _performance_value(value)
            if number is None:
                return ""
            if number < 0:
                return "background-color:#FCE1E4;color:#9E1B32;font-weight:800"
            if number > 0:
                return "background-color:#D7F2E2;color:#146C3E;font-weight:700"
            return ""
        styled = styled.map(deviation_style, subset=[column])
    return styled

def _dashboard_dataframe(data, *args, **kwargs):
    kwargs = _all_rows_table_kwargs(data, kwargs, row_height=26)
    return _streamlit_dataframe(_color_performance_columns(data), *args, **kwargs)

st.dataframe = _dashboard_dataframe

def _dashboard_data_editor(data, *args, **kwargs):
    kwargs = _all_rows_table_kwargs(data, kwargs, row_height=28)
    return _streamlit_data_editor(data, *args, **kwargs)

st.data_editor = _dashboard_data_editor

_original_download_button=st.download_button
# ── PPT export ────────────────────────────────────────────────────────────────
try:
    from pptx import Presentation
    from pptx.util import Inches, Pt, Emu
    from pptx.dml.color import RGBColor
    from pptx.enum.text import PP_ALIGN
    _PPTX_OK = True
except ImportError:
    _PPTX_OK = False

try:
    from meeting_kasie import generate_meeting_kasie_pptx
except ImportError:
    generate_meeting_kasie_pptx = None

def _rgb(hex_str):
    h = hex_str.lstrip("#")
    return RGBColor(int(h[0:2],16), int(h[2:4],16), int(h[4:6],16))

# ── Colours ───────────────────────────────────────────────────────────────────
_C = {
    "bg":      "#000737",  # slide background
    "navy":    "#0C1D42",  # header bar
    "blue":    "#22396F",  # accent
    "cream":   "#FCF1CF",  # primary text
    "silver":  "#B9CADB",  # subtitle
    "green":   "#39C77B",  # good
    "yellow":  "#F5C542",  # medium
    "red":     "#C8102E",  # bad
    "row_a":   "#F4F6F9",  # table row even
    "row_b":   "#E9EEF5",  # table row odd
    "header":  "#22396F",  # table header
    "dark_txt":"#26364A",  # table cell text
    "insight_bg":"#0C1D42",
}

def _clean(text):
    """Strip markdown bold markers and emoji for PPT text."""
    import re
    t = str(text)
    t = re.sub(r"\*\*(.+?)\*\*", r"\1", t)   # **bold** → plain
    t = re.sub(r"\*(.+?)\*",   r"\1", t)   # *italic* → plain
    # keep emoji as-is; PowerPoint 2016+ renders them fine
    return t.strip()

def _slide_bg(slide):
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = _rgb(_C["bg"])

def _header_bar(slide, title, subtitle=""):
    """Navy gradient header bar at top of slide."""
    bar = slide.shapes.add_shape(1, 0, 0, Inches(13.33), Inches(1.1))
    bar.fill.solid(); bar.fill.fore_color.rgb = _rgb(_C["navy"])
    bar.line.fill.background()
    tf = bar.text_frame; tf.word_wrap = True; tf.margin_left = Inches(0.3)
    p = tf.paragraphs[0]; p.text = _clean(title)
    p.alignment = PP_ALIGN.LEFT
    r = p.runs[0]; r.font.size = Pt(28); r.font.bold = True
    r.font.color.rgb = _rgb(_C["cream"])
    if subtitle:
        p2 = tf.add_paragraph(); p2.text = _clean(subtitle)
        p2.alignment = PP_ALIGN.LEFT
        r2 = p2.runs[0]; r2.font.size = Pt(12)
        r2.font.color.rgb = _rgb(_C["silver"])

def _metric_row(slide, metrics, top, card_h=None):
    """metrics = list of (label, value, color_hex). Fills full width evenly."""
    if card_h is None:
        card_h = Inches(1.35)
    n = max(1, len(metrics))
    card_w = Inches(13.33) / n
    pad = Inches(0.1)
    for i, (label, value, color) in enumerate(metrics):
        x = card_w * i + pad
        bg = slide.shapes.add_shape(1, x, top, card_w - pad*2, card_h)
        bg.fill.solid(); bg.fill.fore_color.rgb = _rgb(color)
        bg.line.color.rgb = _rgb("#344a78")
        tf = bg.text_frame; tf.word_wrap = True
        tf.margin_top = Inches(0.08); tf.margin_left = Inches(0.1)
        p1 = tf.paragraphs[0]; p1.text = _clean(label)
        p1.alignment = PP_ALIGN.CENTER
        r1 = p1.runs[0]; r1.font.size = Pt(9); r1.font.color.rgb = _rgb(_C["silver"])
        p2 = tf.add_paragraph(); p2.text = _clean(str(value))
        p2.alignment = PP_ALIGN.CENTER
        r2 = p2.runs[0]; r2.font.size = Pt(18); r2.font.bold = True
        r2.font.color.rgb = _rgb(_C["cream"])

def _table(slide, df, top, left=None, width=None, max_rows=25,
           first_col_pct=0.30):
    """Render DataFrame as PPT table. Returns actual bottom y position."""
    if left is None:
        left = Inches(0.3)
    if width is None:
        width = Inches(12.73)
    if df is None or df.empty: return top
    data = df.head(max_rows)
    nrows = len(data) + 1; ncols = len(df.columns)
    if ncols == 0: return top
    available_h = Inches(7.5) - top - Inches(0.15)
    row_h = min(Inches(0.30), available_h / nrows)
    actual_h = row_h * nrows
    tbl_shape = slide.shapes.add_table(nrows, ncols, left, top, width, actual_h)
    tbl = tbl_shape.table
    # column widths
    first_w = int(width * first_col_pct)
    rest_w  = int((width - first_w) / max(ncols-1, 1))
    tbl.columns[0].width = first_w
    for ci in range(1, ncols): tbl.columns[ci].width = rest_w
    # header row
    for ci, cname in enumerate(df.columns):
        cell = tbl.cell(0, ci)
        cell.text = _clean(str(cname))
        cell.fill.solid(); cell.fill.fore_color.rgb = _rgb(_C["header"])
        pp = cell.text_frame.paragraphs[0]
        pp.alignment = PP_ALIGN.CENTER
        if pp.runs:
            pp.runs[0].font.size = Pt(9); pp.runs[0].font.bold = True
            pp.runs[0].font.color.rgb = _rgb(_C["cream"])
    # data rows
    for ri, (_, row) in enumerate(data.iterrows()):
        rc = _C["row_a"] if ri%2==0 else _C["row_b"]
        for ci, cname in enumerate(df.columns):
            cell = tbl.cell(ri+1, ci)
            val = row[cname]
            cell.text = "—" if (val is None or (isinstance(val,float) and pd.isna(val))) else _clean(str(val))
            cell.fill.solid(); cell.fill.fore_color.rgb = _rgb(rc)
            pp = cell.text_frame.paragraphs[0]
            pp.alignment = PP_ALIGN.LEFT if ci==0 else PP_ALIGN.CENTER
            if pp.runs:
                pp.runs[0].font.size = Pt(8)
                pp.runs[0].font.color.rgb = _rgb(_C["dark_txt"])
    return top + actual_h

def _text_block(slide, lines, left, top, width, height, font_size=11,
                bg=None, title=None, title_color=None):
    """Multi-line text block with optional background and bold title line."""
    box = slide.shapes.add_textbox(left, top, width, height)
    if bg:
        box.fill.solid(); box.fill.fore_color.rgb = _rgb(bg)
    box.line.fill.background()
    tf = box.text_frame; tf.word_wrap = True
    tf.margin_left = Inches(0.12); tf.margin_top = Inches(0.08)
    first = True
    if title:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        p.text = _clean(title); first = False
        p.alignment = PP_ALIGN.LEFT
        if p.runs:
            p.runs[0].font.size = Pt(font_size+1); p.runs[0].font.bold = True
            p.runs[0].font.color.rgb = _rgb(title_color or _C["yellow"])
    for line in lines:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        p.text = _clean(line); first = False
        p.alignment = PP_ALIGN.LEFT
        if p.runs:
            p.runs[0].font.size = Pt(font_size)
            p.runs[0].font.color.rgb = _rgb(_C["cream"])

def _divider(slide, y):
    """Thin accent line."""
    line = slide.shapes.add_shape(1, Inches(0.3), y, Inches(12.73), Inches(0.03))
    line.fill.solid(); line.fill.fore_color.rgb = _rgb(_C["blue"])
    line.line.fill.background()

# ── Cover slide ───────────────────────────────────────────────────────────────
def _cover_slide(prs, title, subtitle, generated_on=""):
    blank = prs.slide_layouts[6]
    sl = prs.slides.add_slide(blank)
    _slide_bg(sl)
    # big navy block
    bg = sl.shapes.add_shape(1, 0, Inches(2.5), Inches(13.33), Inches(2.8))
    bg.fill.solid(); bg.fill.fore_color.rgb = _rgb(_C["navy"])
    bg.line.fill.background()
    # red accent strip
    strip = sl.shapes.add_shape(1, 0, Inches(2.5), Inches(0.18), Inches(2.8))
    strip.fill.solid(); strip.fill.fore_color.rgb = _rgb(_C["red"])
    strip.line.fill.background()
    _text_block(sl, [], Inches(0.5), Inches(2.65), Inches(12.3), Inches(1.0),
                font_size=32, title=title, title_color=_C["cream"])
    _text_block(sl, [subtitle], Inches(0.5), Inches(3.65), Inches(12.3), Inches(0.7),
                font_size=16, bg=None)
    if generated_on:
        _text_block(sl, [f"Generated: {generated_on}"], Inches(0.5), Inches(6.8),
                    Inches(6), Inches(0.4), font_size=10)
    # CHP AAM label
    _text_block(sl, [f"CHP AAM · CABANG {ACTIVE_BRANCH_NAME.upper()}"], Inches(9.0), Inches(6.8),
                Inches(4.0), Inches(0.4), font_size=10)
    return sl

# ── Insight slide ─────────────────────────────────────────────────────────────
def _insight_slide(prs, title, subtitle, status_label, status_color,
                   headline_lines, rekomendasi_lines, extra_sections=None):
    """
    One slide dedicated to AI Insight + Rekomendasi.
    extra_sections = list of (section_title, [lines])
    """
    blank = prs.slide_layouts[6]
    sl = prs.slides.add_slide(blank)
    _slide_bg(sl)
    _header_bar(sl, title, subtitle)

    y = Inches(1.18)
    # Status badge
    badge_w = Inches(3.5)
    badge = sl.shapes.add_shape(1, Inches(0.3), y, badge_w, Inches(0.45))
    badge.fill.solid(); badge.fill.fore_color.rgb = _rgb(status_color)
    badge.line.fill.background()
    p = badge.text_frame.paragraphs[0]; p.text = _clean(status_label)
    p.alignment = PP_ALIGN.CENTER
    if p.runs:
        p.runs[0].font.size = Pt(13); p.runs[0].font.bold = True
        p.runs[0].font.color.rgb = _rgb(_C["cream"])

    y += Inches(0.52)
    # Headline block
    if headline_lines:
        hl_h = Inches(0.38) * len(headline_lines) + Inches(0.25)
        _text_block(sl, headline_lines, Inches(0.3), y,
                    Inches(12.73), min(hl_h, Inches(1.4)),
                    font_size=12, bg=_C["insight_bg"])
        y += min(hl_h, Inches(1.4)) + Inches(0.1)

    _divider(sl, y); y += Inches(0.1)

    # Rekomendasi
    if rekomendasi_lines:
        rek_h = min(Inches(0.32)*len(rekomendasi_lines)+Inches(0.3), Inches(7.5)-y-Inches(0.2))
        _text_block(sl, rekomendasi_lines, Inches(0.3), y,
                    Inches(12.73), rek_h,
                    font_size=10, bg="#0A1530",
                    title="⚡ Rekomendasi Taktis", title_color=_C["yellow"])
        y += rek_h + Inches(0.1)

    # Extra sections (e.g. per-SP insight)
    if extra_sections:
        for sec_title, sec_lines in extra_sections:
            if y >= Inches(7.2): break
            sec_h = min(Inches(0.28)*len(sec_lines)+Inches(0.3), Inches(7.5)-y-Inches(0.1))
            _text_block(sl, sec_lines, Inches(0.3), y, Inches(12.73), sec_h,
                        font_size=10, bg="#111E3F",
                        title=sec_title, title_color=_C["silver"])
            y += sec_h + Inches(0.08)
    return sl

# ═══════════════════════════════════════════════════════════════════════════════
# BUILD FUNCTIONS
# ═══════════════════════════════════════════════════════════════════════════════

def build_ppt_performance(perf_data):
    """
    PPT Performance Sales.
    perf_data keys: subtitle, month_name, year, m_actual/m_target/m_ach/m_dev,
                    y_actual/y_target/y_ach/y_dev, fot, headline, status_label,
                    status_color, rekomendasi(list), sp_insight_lines(list),
                    trend_dir, best_month_name, monthly_over, daily_run_rate,
                    daily_needed, days_left, gap, principal_df, salesperson_df.
    """
    if not _PPTX_OK: return None
    import datetime
    prs = Presentation(); prs.slide_width = Inches(13.33); prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]
    now = datetime.datetime.now().strftime("%d %b %Y %H:%M")

    # ── SLIDE 1: Cover ────────────────────────────────────────────────────────
    _cover_slide(prs, "PERFORMANCE SALES · GROSS",
                 perf_data.get("subtitle",f"CHP AAM · Cabang {ACTIVE_BRANCH_NAME}"), now)

    # ── SLIDE 2: KPI Overview ─────────────────────────────────────────────────
    sl2 = prs.slides.add_slide(blank); _slide_bg(sl2)
    _header_bar(sl2, "KPI Overview", perf_data.get("subtitle",""))
    mn = perf_data.get("month_name","")
    yr = perf_data.get("year","")
    m_ach = perf_data.get("m_ach") or 0
    ach_color = _C["red"] if m_ach<60 else (_C["yellow"] if m_ach<90 else _C["green"])
    metrics_m = [
        (f"Actual · {mn}", f"Rp {perf_data.get('m_actual',0):,.0f}", _C["blue"]),
        (f"Target · {mn}", f"Rp {perf_data.get('m_target',0):,.0f}", _C["navy"]),
        ("A/T Bulanan",    f"{m_ach:.1f}%",                           ach_color),
        ("Deviasi Bulanan",f"Rp {perf_data.get('m_dev',0):,.0f}",     _C["navy"]),
    ]
    _metric_row(sl2, metrics_m, top=Inches(1.2), card_h=Inches(1.45))
    y_ach = perf_data.get("y_ach") or 0
    metrics_y = [
        ("Actual YTD",  f"Rp {perf_data.get('y_actual',0):,.0f}", _C["navy"]),
        ("Target YTD",  f"Rp {perf_data.get('y_target',0):,.0f}", _C["navy"]),
        ("A/T YTD",    f"{y_ach:.1f}%", _C["navy"]),
        ("Deviasi YTD",f"Rp {perf_data.get('y_dev',0):,.0f}",     _C["navy"]),
        (f"FOT {yr}",  f"{perf_data.get('fot',0)}×",               _C["navy"]),
    ]
    _metric_row(sl2, metrics_y, top=Inches(2.78), card_h=Inches(1.15))
    # Key numbers strip
    dl = perf_data.get("days_left",0); gap = perf_data.get("gap",0)
    drr = perf_data.get("daily_run_rate",0); dnd = perf_data.get("daily_needed",0)
    kn_lines = [
        f"Run rate harian saat ini: {drr:,.0f}/hari  |  Run rate dibutuhkan: {dnd:,.0f}/hari",
        f"Gap tersisa: Rp {gap:,.0f}  |  Hari tersisa: {dl} hari",
        f"Trend A/T: {perf_data.get('trend_dir','—')}  |  "
        f"Bulan terbaik: {perf_data.get('best_month_name','—')}  |  "
        f"Principal ≥100% bulanan: {perf_data.get('monthly_over',0)}",
    ]
    _text_block(sl2, kn_lines, Inches(0.3), Inches(4.05), Inches(12.73), Inches(0.9),
                font_size=10, bg=_C["navy"])

    # ── SLIDE 3: AI Insight & Rekomendasi ─────────────────────────────────────
    _insight_slide(
        prs,
        title="AI Insight · Strategi Capai Target Bulanan",
        subtitle=perf_data.get("subtitle",""),
        status_label=perf_data.get("status_label","—"),
        status_color=_C["red"] if "KRITIS" in perf_data.get("status_label","") else
                     (_C["yellow"] if "HAMPIR" in perf_data.get("status_label","") or
                                      "AKSELERASI" in perf_data.get("status_label","") else _C["green"]),
        headline_lines=[perf_data.get("headline","")],
        rekomendasi_lines=perf_data.get("rekomendasi",[]),
        extra_sections=[("👤 Fokus Sales Person Title",
                          perf_data.get("sp_insight_lines",[]))]
                         if perf_data.get("sp_insight_lines") else None,
    )

    # ── SLIDE 4: Principal ID Detail ──────────────────────────────────────────
    pid_df = perf_data.get("principal_df")
    if pid_df is not None and not pid_df.empty:
        sl4 = prs.slides.add_slide(blank); _slide_bg(sl4)
        _header_bar(sl4, f"Detail by Principal ID · {mn} {yr}", perf_data.get("subtitle",""))
        _table(sl4, pid_df, top=Inches(1.2), left=Inches(0.3), width=Inches(12.73))

    # ── SLIDE 5: Sales Person Title Detail ────────────────────────────────────
    sp_df = perf_data.get("salesperson_df")
    if sp_df is not None and not sp_df.empty:
        sl5 = prs.slides.add_slide(blank); _slide_bg(sl5)
        _header_bar(sl5, f"Detail by Sales Person Title · {mn} {yr}", perf_data.get("subtitle",""))
        _table(sl5, sp_df, top=Inches(1.2), left=Inches(0.3), width=Inches(12.73))

    buf = io.BytesIO(); prs.save(buf); buf.seek(0)
    return buf.getvalue()


def build_ppt_summary_program(prog_data_list, month_label_str=""):
    """
    PPT Summary 5 Program Push.
    prog_data_list item keys: label, actual_text, target_text, achievement,
                              deviation_text, status, title_df,
                              insight_parts (list of (text,is_neg)),
                              insight_status, insight_level.
    """
    if not _PPTX_OK: return None
    import datetime
    prs = Presentation(); prs.slide_width = Inches(13.33); prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]
    now = datetime.datetime.now().strftime("%d %b %Y %H:%M")

    # ── Cover ─────────────────────────────────────────────────────────────────
    _cover_slide(prs, f"SUMMARY {len(prog_data_list)} PROGRAM PUSH", month_label_str, now)

    # ── Slide 2: Dashboard ringkasan ──────────────────────────────────────────
    sl2 = prs.slides.add_slide(blank); _slide_bg(sl2)
    _header_bar(sl2, f"Ringkasan Pencapaian {len(prog_data_list)} Program", month_label_str)
    metrics = []
    for pd_item in prog_data_list:
        ach = pd_item.get("achievement")
        color = _C["red"] if ach is None or ach<60 else (_C["yellow"] if ach<90 else _C["green"])
        metrics.append((pd_item.get("label",""), f"A/T {ach:.1f}%" if ach else "—", color))
    if metrics:
        _metric_row(sl2, metrics, top=Inches(1.2), card_h=Inches(1.55))
    rows = []
    for pd_item in prog_data_list:
        ach = pd_item.get("achievement")
        rows.append({
            "Program": pd_item.get("label",""),
            "Actual":  pd_item.get("actual_text","—"),
            "Target":  pd_item.get("target_text","—"),
            "A/T":     f"{ach:.1f}%" if ach else "—",
            "Deviasi": pd_item.get("deviation_text","—"),
            "Status":  pd_item.get("status","—"),
            "AI Insight": pd_item.get("insight_status","—"),
        })
    _table(sl2, pd.DataFrame(rows), top=Inches(2.9), left=Inches(0.3), width=Inches(12.73))

    # ── Per-program slides ────────────────────────────────────────────────────
    for pd_item in prog_data_list:
        prog_label = pd_item.get("label","Program")
        ach = pd_item.get("achievement")
        color = _C["red"] if ach is None or ach<60 else (_C["yellow"] if ach<90 else _C["green"])

        # Slide A: KPI + tabel title
        slA = prs.slides.add_slide(blank); _slide_bg(slA)
        _header_bar(slA, prog_label, month_label_str)
        kpi_metrics = [
            ("Actual",  pd_item.get("actual_text","—"),              _C["blue"]),
            ("Target",  pd_item.get("target_text","—"),              _C["navy"]),
            ("A/T",     f"{ach:.1f}%" if ach else "—",              color),
            ("Deviasi", pd_item.get("deviation_text","—"),           _C["navy"]),
            ("Status",  pd_item.get("status","—"),                   color),
        ]
        _metric_row(slA, kpi_metrics, top=Inches(1.2), card_h=Inches(1.2))
        title_df = pd_item.get("title_df")
        if title_df is not None and not title_df.empty:
            disp = title_df[["Sales Person Title","Actual","Target","A/T (%)","Deviation"]].copy()
            disp["A/T (%)"]   = disp["A/T (%)"].apply(lambda v: f"{v:.1f}%" if pd.notna(v) else "—")
            disp["Actual"]    = disp["Actual"].apply(lambda v: f"{v:,.0f}" if pd.notna(v) else "—")
            disp["Target"]    = disp["Target"].apply(lambda v: f"{v:,.0f}" if pd.notna(v) else "—")
            disp["Deviation"] = disp["Deviation"].apply(lambda v: f"{v:,.0f}" if pd.notna(v) else "—")
            _table(slA, disp, top=Inches(2.55), left=Inches(0.3), width=Inches(12.73))

        # Slide B: AI Insight per program
        insight_parts = pd_item.get("insight_parts", [])
        if insight_parts:
            insight_lines = [text for text, _ in insight_parts]
            ins_status = pd_item.get("insight_status","—")
            ins_color  = (_C["red"] if pd_item.get("insight_level","medium")=="low"
                          else _C["yellow"] if pd_item.get("insight_level")=="medium"
                          else _C["green"])
            _insight_slide(
                prs,
                title=f"AI Insight · {prog_label}",
                subtitle=month_label_str,
                status_label=ins_status,
                status_color=ins_color,
                headline_lines=insight_lines[:3],
                rekomendasi_lines=insight_lines[3:],
            )

    buf = io.BytesIO(); prs.save(buf); buf.seek(0)
    return buf.getvalue()


def build_ppt_program_by_title(program_name, by_title_df, selected_filters,
                                actual, tgt, achievement, deviation, metric,
                                insight_lines=None, rekomendasi_lines=None,
                                weekly_label=""):
    """PPT Program by Sales Person Title dengan insight."""
    if not _PPTX_OK: return None
    import datetime
    prs = Presentation(); prs.slide_width = Inches(13.33); prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]
    now = datetime.datetime.now().strftime("%d %b %Y %H:%M")
    subtitle = f"By Sales Person Title · {weekly_label}" if weekly_label else "By Sales Person Title"

    # Cover
    _cover_slide(prs, f"PROGRAM PUSH · {program_name}", subtitle, now)

    # KPI + tabel
    sl2 = prs.slides.add_slide(blank); _slide_bg(sl2)
    _header_bar(sl2, f"PROGRAM PUSH · {program_name}", subtitle)
    ach_v = achievement or 0
    color = _C["red"] if ach_v<60 else (_C["yellow"] if ach_v<90 else _C["green"])
    act_txt = fmt(actual, metric) if actual is not None else "—"
    tgt_txt = fmt(tgt, metric) if tgt else "—"
    dev_txt = fmt(deviation, metric) if deviation is not None else "—"
    kpi = [
        ("Actual",  act_txt,             _C["blue"]),
        ("Target",  tgt_txt,             _C["navy"]),
        ("A/T",     f"{ach_v:.1f}%",     color),
        ("Deviasi", dev_txt,             _C["navy"]),
    ]
    _metric_row(sl2, kpi, top=Inches(1.2), card_h=Inches(1.3))
    if by_title_df is not None and not by_title_df.empty:
        disp = by_title_df[["Sales Person Title","Actual","Target","A/T (%)","Deviation"]].copy()
        disp["A/T (%)"]   = disp["A/T (%)"].apply(lambda v: f"{v:.1f}%" if pd.notna(v) else "—")
        disp["Actual"]    = disp["Actual"].apply(lambda v: f"{v:,.0f}" if pd.notna(v) else "—")
        disp["Target"]    = disp["Target"].apply(lambda v: f"{v:,.0f}" if pd.notna(v) else "—")
        disp["Deviation"] = disp["Deviation"].apply(lambda v: f"{v:,.0f}" if pd.notna(v) else "—")
        _table(sl2, disp, top=Inches(2.65), left=Inches(0.3), width=Inches(12.73))

    # Insight slide
    if insight_lines or rekomendasi_lines:
        status = "BAGUS" if ach_v>=90 else ("SEDANG" if ach_v>=60 else "RENDAH")
        status_color = _C["green"] if ach_v>=90 else (_C["yellow"] if ach_v>=60 else _C["red"])
        _insight_slide(
            prs,
            title=f"AI Insight · {program_name}",
            subtitle=subtitle,
            status_label=status,
            status_color=status_color,
            headline_lines=insight_lines or [],
            rekomendasi_lines=rekomendasi_lines or [],
        )

    buf = io.BytesIO(); prs.save(buf); buf.seek(0)
    return buf.getvalue()


def build_ppt_sasaran_program(programs_list, program_labels_fn,
                               potential_summary_df, detail_df, visit_table_df,
                               ai_insight_lines=None, rekomendasi_lines=None):
    """PPT Sasaran Program dengan insight dan rencana kunjungan."""
    if not _PPTX_OK: return None
    import datetime
    prs = Presentation(); prs.slide_width = Inches(13.33); prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]
    now = datetime.datetime.now().strftime("%d %b %Y %H:%M")

    # Cover
    _cover_slide(prs, "SASARAN PROGRAM PUSH", "Potensi & Rencana Kunjungan", now)

    # Slide 2: Ringkasan potensi
    if potential_summary_df is not None and not potential_summary_df.empty:
        sl2 = prs.slides.add_slide(blank); _slide_bg(sl2)
        _header_bar(sl2, "Ringkasan Potensi Belum Transaksi", "Per Program")
        _table(sl2, potential_summary_df, top=Inches(1.2), left=Inches(0.3), width=Inches(12.73))

    # Slide 3: Detail per salesperson
    if detail_df is not None and not detail_df.empty:
        # Pagination: 25 baris per slide
        chunk_size = 25
        chunks = [detail_df.iloc[i:i+chunk_size] for i in range(0, min(len(detail_df),75), chunk_size)]
        for ci, chunk in enumerate(chunks):
            slD = prs.slides.add_slide(blank); _slide_bg(slD)
            pg = f" ({ci+1}/{len(chunks)})" if len(chunks)>1 else ""
            _header_bar(slD, f"Detail Belum Transaksi per Salesperson{pg}", "")
            _table(slD, chunk, top=Inches(1.2), left=Inches(0.3), width=Inches(12.73))

    # Slide 4: Rencana kunjungan (visit_table ringkasan — top 30)
    if visit_table_df is not None and not visit_table_df.empty:
        # Tampilkan kolom kunci saja
        disp_cols = ["Sales Person Title","Relasi","Fokus kunjungan"]
        disp_cols = [c for c in disp_cols if c in visit_table_df.columns]
        if disp_cols:
            visit_disp = visit_table_df[disp_cols].head(30)
        else:
            visit_disp = visit_table_df.head(30)
        slV = prs.slides.add_slide(blank); _slide_bg(slV)
        _header_bar(slV, "Rencana Kunjungan", f"{len(visit_table_df):,} relasi dalam cakupan")
        _table(slV, visit_disp, top=Inches(1.2), left=Inches(0.3),
               width=Inches(12.73), first_col_pct=0.22)

    # Slide 5: AI Insight & Rekomendasi
    if ai_insight_lines or rekomendasi_lines:
        _insight_slide(
            prs,
            title="AI Insight · Sasaran & Potensi Program",
            subtitle="Rekomendasi berdasarkan data transaksi aktual",
            status_label="PRIORITAS KUNJUNGAN",
            status_color=_C["blue"],
            headline_lines=ai_insight_lines or [],
            rekomendasi_lines=rekomendasi_lines or [],
        )

    buf = io.BytesIO(); prs.save(buf); buf.seek(0)
    return buf.getvalue()

# ─────────────────────────────────────────────────────────────────────────────

st.set_page_config(page_title="CHP AAM | Push Performance", page_icon="📊", layout="wide")
RED, NAVY, BLUE = "#C8102E", "#000737", "#22396F"
MIDNIGHT_BLUE, CREAM = "#0C1D42", "#FCF1CF"
PROGRAMS = ["RB CHD", "NOO PDB", "PB NOZA", "PB HYC", "PB LUMA"]
PROGRAM_NUMBERS = {"RB CHD":"PROGRAM 1", "NOO PDB":"PROGRAM 2", "PB NOZA":"PROGRAM 3", "PB HYC":"PROGRAM 4", "PB LUMA":"PROGRAM 5"}
PROGRAM_TARGET_DEFAULTS = {"RB CHD":370, "NOO PDB":51, "PB NOZA":119, "PB HYC":285, "PB LUMA":46}
PROGRAM_FILTER_DEFAULTS = {
    "RB CHD":{"year":"All","month":"September","principal_name":"All","principal":"CHD","product_code":"All","product":"All","metric":"Total Relasi Beli Transaksi"},
    "NOO PDB":{"year":"All","month":"All","principal_name":"All","principal":"PDB","product_code":"All","product":"All","metric":"Total Relasi Beli Transaksi"},
    "PB NOZA":{"year":"All","month":"All","principal_name":"All","principal":"All","product_code":"NOZA001","product":"All","metric":"Total Relasi Beli Transaksi"},
    "PB HYC":{"year":"All","month":"All","principal_name":"All","principal":"HYC","product_code":"All","product":"All","metric":"Total Relasi Beli Transaksi"},
    "PB LUMA":{"year":"All","month":"All","principal_name":"All","principal":"All","product_code":"All","product":"LASERIN UNTUK MASUK ANGIN (LUMA)","metric":"Total Relasi Beli Transaksi"},
}
for _program in PROGRAMS:
    PROGRAM_FILTER_DEFAULTS[_program].setdefault("channel_group","All")
PROGRAM_SUMMARIES = {
    "RB CHD": "Monitoring program RB CHD sesuai principal ID dan periode pada kartu.",
    "NOO PDB": "Monitoring program NOO PDB sesuai principal ID dan periode pada kartu.",
    "PB NOZA": "Monitoring program PB NOZA dengan fokus produk NOZA001.",
    "PB HYC": "Monitoring program PB HYC sesuai principal ID dan periode pada kartu.",
    "PB LUMA": "Monitoring program PB LUMA dengan fokus LASERIN UNTUK MASUK ANGIN (LUMA).",
}
MONTHS = {1:"January",2:"February",3:"March",4:"April",5:"May",6:"June",7:"July",8:"August",9:"September",10:"October",11:"November",12:"December"}
ALIASES = {
 "year": ["year","tahun"], "month": ["month","bulan"], "date": ["full date","date","transaction date","tanggal"], "business_group": ["bisnis group","business group","business_group","group business"], "principal_name": ["principal"], "principal": ["principal id","principal_id","principle id"],
 "product_code": ["product code","product_code","kode produk"], "channel_group": ["channel group","channel_group","group channel"], "product": ["product description","product_description","product","description produk"],
 "sales": ["sales value gross","sales_value_gross","sales","total sales","sales value","gross sales"], "target_value": ["target value","target_value","sales target","target sales"],
 "customer": ["customer name","customer_name","customer","relasi","customer id"],
 "salesperson": ["sales person title","salesperson title","sales_person_title","sales person","salesperson","sales title"],
 "program": ["program","push program","program push"],
}
APP_DIR=Path(__file__).parent
ASSET_DIR=APP_DIR/"assets"
if not ASSET_DIR.exists():
    ASSET_DIR=APP_DIR
BRANCH_OPTIONS={"BDG":"Bandung","PTK":"Pontianak","MDN":"Medan"}
# Folder link provided by the user. An explicit file ID in secrets overrides it.
DEFAULT_DRIVE_DATABASE_FOLDER_ID="1Ef0DsWdqJ-fy7ahuQ0NR3H1dO2QkHk4b"
def _drive_service_credentials():
    """Build credentials with write access limited by the service account's Drive sharing."""
    from google.oauth2.service_account import Credentials
    return Credentials.from_service_account_info(
        dict(st.secrets["gcp_service_account"]),
        scopes=["https://www.googleapis.com/auth/drive"],
    )

def _drive_folder_id(folder_name, parent_id):
    from google.auth.transport.requests import AuthorizedSession
    session=AuthorizedSession(_drive_service_credentials())
    response=session.get("https://www.googleapis.com/drive/v3/files",params={
        "q":f"'{parent_id}' in parents and name = '{folder_name}' and mimeType = 'application/vnd.google-apps.folder' and trashed = false",
        "fields":"files(id,name)","pageSize":100,
        "supportsAllDrives":"true","includeItemsFromAllDrives":"true",
    },timeout=30)
    response.raise_for_status()
    folders=response.json().get("files",[])
    if folders:
        return folders[0]["id"]
    response=session.post("https://www.googleapis.com/drive/v3/files",params={
        "fields":"id","supportsAllDrives":"true",
    },json={"name":folder_name,"mimeType":"application/vnd.google-apps.folder","parents":[parent_id]},timeout=30)
    response.raise_for_status()
    return response.json()["id"]

def _drive_storage_folder(category, branch=None):
    """Return/create a folder inside the user-selected shared dashboard folder."""
    drive_settings=st.secrets.get("google_drive",{})
    root_id=str(drive_settings.get("storage_folder_id",drive_settings.get("database_folder_id",DEFAULT_DRIVE_DATABASE_FOLDER_ID))).strip()
    if not root_id:
        raise RuntimeError("Folder penyimpanan Google Drive belum dikonfigurasi.")
    folder_id=_drive_folder_id("Dashboard_Data" if category!="exports" else "Dashboard_Exports",root_id)
    if branch:
        folder_id=_drive_folder_id(str(branch),folder_id)
    if category not in ("", "exports"):
        folder_id=_drive_folder_id(str(category),folder_id)
    return folder_id

def _drive_upload_bytes(file_name,payload,mime_type="application/octet-stream",category="exports",branch=None):
    """Create or replace a file in the shared Drive folder."""
    from google.auth.transport.requests import AuthorizedSession
    import uuid
    folder_id=_drive_storage_folder(category,branch)
    session=AuthorizedSession(_drive_service_credentials())
    safe_name=str(file_name).replace("'","\\'")
    response=session.get("https://www.googleapis.com/drive/v3/files",params={
        "q":f"'{folder_id}' in parents and name = '{safe_name}' and trashed = false",
        "fields":"files(id,name)","pageSize":100,
        "supportsAllDrives":"true","includeItemsFromAllDrives":"true",
    },timeout=30)
    response.raise_for_status()
    boundary="codex_"+uuid.uuid4().hex
    body=(f"--{boundary}\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n"+
          json.dumps({"name":str(file_name),"mimeType":mime_type},ensure_ascii=False)+
          f"\r\n--{boundary}\r\nContent-Type: {mime_type}\r\n\r\n").encode("utf-8")+payload+f"\r\n--{boundary}--".encode("utf-8")
    matches=response.json().get("files",[])
    if matches:
        url=f"https://www.googleapis.com/upload/drive/v3/files/{matches[0]['id']}"
        response=session.patch(url,params={"uploadType":"multipart","supportsAllDrives":"true"},
            data=body,headers={"Content-Type":f"multipart/related; boundary={boundary}"},timeout=120)
    else:
        url="https://www.googleapis.com/upload/drive/v3/files"
        response=session.post(url,params={"uploadType":"multipart","supportsAllDrives":"true"},
            data=body,headers={"Content-Type":f"multipart/related; boundary={boundary}"},timeout=120)
    response.raise_for_status()
    return response.json().get("id")

def _drive_find_file(file_name,category,branch=None):
    from google.auth.transport.requests import AuthorizedSession
    folder_id=_drive_storage_folder(category,branch)
    response=AuthorizedSession(_drive_service_credentials()).get("https://www.googleapis.com/drive/v3/files",params={
        "q":f"'{folder_id}' in parents and name = '{str(file_name).replace(chr(39), chr(92)+chr(39))}' and trashed = false",
        "fields":"files(id,name)","pageSize":100,
        "supportsAllDrives":"true","includeItemsFromAllDrives":"true",
    },timeout=30)
    response.raise_for_status()
    files=response.json().get("files",[])
    return files[0]["id"] if files else None

def _drive_download_to_path(file_name,path,category,branch=None):
    from google.auth.transport.requests import AuthorizedSession
    file_id=_drive_find_file(file_name,category,branch)
    if not file_id:
        return False
    response=AuthorizedSession(_drive_service_credentials()).get(
        f"https://www.googleapis.com/drive/v3/files/{file_id}",params={"alt":"media","supportsAllDrives":"true"},timeout=120)
    response.raise_for_status()
    path=Path(path); path.parent.mkdir(parents=True,exist_ok=True); path.write_bytes(response.content)
    return True

def _drive_delete_file(file_name,category,branch=None):
    from google.auth.transport.requests import AuthorizedSession
    file_id=_drive_find_file(file_name,category,branch)
    if not file_id: return False
    response=AuthorizedSession(_drive_service_credentials()).patch(
        f"https://www.googleapis.com/drive/v3/files/{file_id}",params={"supportsAllDrives":"true"},json={"trashed":True},timeout=30)
    response.raise_for_status()
    return True

def _as_download_bytes(data):
    if isinstance(data,bytes): return data
    if isinstance(data,(bytearray,memoryview)): return bytes(data)
    if isinstance(data,str): return data.encode("utf-8")
    if hasattr(data,"getvalue"): return data.getvalue()
    if hasattr(data,"read"):
        position=data.tell() if hasattr(data,"tell") else None
        payload=data.read()
        if position is not None and hasattr(data,"seek"): data.seek(position)
        return payload
    return bytes(data)

# Export controls stay as one download action. Operational data and
# configuration are synchronized automatically through their own save paths.

def _reset_branch_session():
    _keep={"active_branch_code","db"}
    for _state_key in list(st.session_state.keys()):
        if _state_key not in _keep:
            del st.session_state[_state_key]
if "active_branch_code" not in st.session_state:
    st.session_state["active_branch_code"]="BDG"
_branch_choice=st.selectbox("Pilih Cabang",list(BRANCH_OPTIONS),format_func=lambda code:f"{BRANCH_OPTIONS[code]} ({code})",key="active_branch_code",on_change=_reset_branch_session)
ACTIVE_BRANCH_CODE=_branch_choice
ACTIVE_BRANCH_NAME=BRANCH_OPTIONS[ACTIVE_BRANCH_CODE]
BRANCH_THEMES={
    "BDG":{"red":"#C8102E","navy":"#000737","blue":"#22396F","midnight":"#0C1D42","cream":"#FCF1CF","landmarks":"GEDUNG SATE · PARIS VAN JAVA · KOTA KEMBANG"},
    "PTK":{"red":"#E36B55","navy":"#052B35","blue":"#087E83","midnight":"#0B414B","cream":"#E7FAF4","landmarks":"TUGU KHATULISTIWA · SUNGAI KAPUAS · KOTA KHATULISTIWA"},
    "MDN":{"red":"#D66A45","navy":"#21142F","blue":"#79415F","midnight":"#35203E","cream":"#FFF0D2","landmarks":"ISTANA MAIMUN · MASJID RAYA · KOTA MEDAN"},
}
_theme=BRANCH_THEMES[ACTIVE_BRANCH_CODE]
RED,NAVY,BLUE=_theme["red"],_theme["navy"],_theme["blue"]
MIDNIGHT_BLUE,CREAM=_theme["midnight"],_theme["cream"]
BRANCH_DATA_DIR=APP_DIR/"branch_data"/ACTIVE_BRANCH_CODE
BRANCH_DATA_DIR.mkdir(parents=True,exist_ok=True)
# Bandung continues using the existing paths; new branches get fully isolated files.
SETTINGS_PATH=APP_DIR/"dashboard_settings.json" if ACTIVE_BRANCH_CODE=="BDG" else BRANCH_DATA_DIR/"dashboard_settings.json"
PROGRAM_TARGET_DIR=APP_DIR/"saved_program_targets" if ACTIVE_BRANCH_CODE=="BDG" else BRANCH_DATA_DIR/"saved_program_targets"
SAVED_TARGET_DIR=APP_DIR if ACTIVE_BRANCH_CODE=="BDG" else BRANCH_DATA_DIR
OVERVIEW_DATA_DIR=APP_DIR if ACTIVE_BRANCH_CODE=="BDG" else BRANCH_DATA_DIR
DAILY_DISCIPLINE_PATH=APP_DIR/"daily_discipline_latest.csv" if ACTIVE_BRANCH_CODE=="BDG" else BRANCH_DATA_DIR/"daily_discipline_latest.csv"
VISIT_LOG_PATH=APP_DIR/"visit_log.csv" if ACTIVE_BRANCH_CODE=="BDG" else BRANCH_DATA_DIR/"visit_log.csv"
def _restore_drive_state():
    """Restore persistent operational files from Drive before loading app state."""
    if not st.secrets.get("gcp_service_account") or st.session_state.get("_drive_state_restored_branch")==ACTIVE_BRANCH_CODE:
        return
    files=[
        (SETTINGS_PATH,"dashboard_settings.json","settings"),
        (VISIT_LOG_PATH,"visit_log.csv","visit_logs"),
        (DAILY_DISCIPLINE_PATH,"daily_discipline_latest.csv","reports"),
        (OVERVIEW_DATA_DIR/"overview_calendar.xlsx","overview_calendar.xlsx","overview"),
        (OVERVIEW_DATA_DIR/"overview_fixed_call.xlsx","overview_fixed_call.xlsx","overview"),
    ]
    try:
        for path,name,category in files:
            _drive_download_to_path(name,path,category,ACTIVE_BRANCH_CODE)
        PROGRAM_TARGET_DIR.mkdir(parents=True,exist_ok=True)
        folder_id=_drive_storage_folder("program_targets",ACTIVE_BRANCH_CODE)
        from google.auth.transport.requests import AuthorizedSession
        session=AuthorizedSession(_drive_service_credentials())
        response=session.get("https://www.googleapis.com/drive/v3/files",params={
            "q":f"'{folder_id}' in parents and trashed = false","fields":"files(id,name,mimeType)","pageSize":1000,
            "supportsAllDrives":"true","includeItemsFromAllDrives":"true",
        },timeout=30)
        response.raise_for_status()
        for item in response.json().get("files",[]):
            if item.get("mimeType")=="application/vnd.google-apps.folder": continue
            payload=session.get(f"https://www.googleapis.com/drive/v3/files/{item['id']}",params={"alt":"media","supportsAllDrives":"true"},timeout=120)
            payload.raise_for_status()
            (PROGRAM_TARGET_DIR/item["name"]).write_bytes(payload.content)
        st.session_state["_drive_state_restored_branch"]=ACTIVE_BRANCH_CODE
    except Exception as exc:
        st.warning(f"Tidak dapat memuat file simpanan dari Google Drive: {exc}")
_restore_drive_state()
def load_settings():
    try: return json.loads(SETTINGS_PATH.read_text(encoding="utf-8"))
    except (OSError, ValueError): return {}
def save_settings(settings):
    SETTINGS_PATH.parent.mkdir(parents=True,exist_ok=True)
    payload=json.dumps(settings,ensure_ascii=False,indent=2).encode("utf-8")
    SETTINGS_PATH.write_bytes(payload)
    if st.secrets.get("gcp_service_account"):
        try: _drive_upload_bytes(SETTINGS_PATH.name,payload,"application/json","settings",ACTIVE_BRANCH_CODE)
        except Exception as exc: st.error(f"Konfigurasi tersimpan sementara, tetapi gagal disinkronkan ke Google Drive: {exc}")
settings=load_settings()
def program_label(program, month=None):
    if month is None:
        month=globals().get("_active_push_month")
    if month is None:
        month_index=int(st.session_state.get("push_month_tab_idx",pd.Timestamp.now().month-1))
        month_index=min(max(month_index,0),11)
        month=list(MONTHS.values())[month_index]
    configured_name=settings.get("program_config_by_month",{}).get(month,{}).get(program,{}).get("display_name")
    return str(configured_name or program)

def key(s): return " ".join(str(s).strip().lower().replace("_"," ").split())
def col(df, kind):
    lookup={key(c):c for c in df.columns}
    return next((lookup[key(a)] for a in ALIASES[kind] if key(a) in lookup), None)
def normalize(df,copy=True):
    if copy:
        df=df.copy()
    for kind in ALIASES:
        c=col(df,kind)
        if c and kind != "sales":
            if isinstance(df[c].dtype,pd.CategoricalDtype):
                clean=df[c].cat.categories.astype("string").str.strip()
                df[c]=df[c].cat.rename_categories(clean)
            else:
                df[c]=df[c].astype("string").str.strip()
    return df
def month_label(value):
    try:
        n=int(float(str(value).strip()))
        return MONTHS.get(n,str(value))
    except (ValueError,TypeError):
        return str(value).strip()
def positive_sales_rows(df):
    sales_col=col(df,"sales")
    if not sales_col:
        return None
    positive_mask=pd.to_numeric(df[sales_col],errors="coerce").fillna(0)>0
    return df.loc[positive_mask]

def metric_actual(df, metric):
    if metric=="Total Sales":
        sales_col=col(df,"sales")
        if not sales_col:
            return None
        # Total Sales includes every transaction in the selected scope;
        # unlike RB it is not restricted to positive-value rows.
        return float(pd.to_numeric(df[sales_col],errors="coerce").fillna(0).sum())
    eligible=positive_sales_rows(df)
    if eligible is None:
        return None
    customer_col=col(eligible,"customer")
    if not customer_col:
        return int(len(eligible))
    customers=eligible[customer_col].astype("string").str.strip()
    customers=customers[customers.notna() & customers.ne("")]
    return int(customers.nunique())
def fmt(x, metric):
    if x is None: return "—"
    return f"{x:,.0f}" if metric!="Total Sales" else f"Rp {x:,.0f}"
def fmt_summary(x, metric):
    """Compact monetary values in thousands for Program Push summary cards."""
    if x is None: return "—"
    if metric!="Total Sales": return fmt(x,metric)
    thousands=f"{float(x)/1000:,.0f}".replace(",",".")
    return f"Rp {thousands} rb"
def target_template():
    return pd.DataFrame(columns=["Year","Month","Sales Person Title","Target","Program","Principal","Principal ID","Channel Group","Business Group","Product Code","Product Description"])
def filter_options(df,kind):
    c=col(df,kind)
    if not c: return []
    # Filter widgets repeatedly request the same option lists on every rerun.
    # Cache these small lists per canonical database frame instead of rescanning
    # the 963k-row workbook each time.
    is_base_frame=df is globals().get("df")
    option_cache=st.session_state.setdefault("_base_filter_option_cache",{}) if is_base_frame else None
    option_key=(id(df),kind)
    if option_cache is not None and option_key in option_cache:
        return list(option_cache[option_key])
    vals=[str(value) for value in df[c].dropna().unique().tolist()]
    if kind=="month": vals=[month_label(v) for v in vals]
    result=sorted(set(vals))
    if option_cache is not None:
        option_cache[option_key]=result
    return result
def filter_select(df, kind, label, keyname):
    c=col(df,kind)
    if not c: st.caption(f"Filter {label}: kolom tidak ditemukan di database."); return []
    opts=filter_options(df,kind)
    return st.multiselect(label,opts,key=keyname,help="Pilih beberapa nilai; kosongkan untuk mencakup semua.")
def matches_dimension(series,kind,value):
    values=value if isinstance(value,(list,tuple,set)) else [value]
    values=[v for v in values if v not in (None,"All","")]
    if not values: return pd.Series(True,index=series.index)
    normalized=series.astype(str).map(month_label) if kind=="month" else series.astype(str)
    expected=[month_label(v) if kind=="month" else str(v) for v in values]
    return normalized.isin(expected)
def target_filtered(t, params, metric, require_explicit=()):
    if t is None or t.empty: return None, pd.DataFrame()
    tc={k:col(t,k) for k in ["year","month","principal_name","principal","channel_group","business_group","product_code","product","salesperson","program"]}
    amount=next((c for c in t.columns if key(c) in ["target","target value","target sales","target relasi","target transaksi"]),None)
    if not amount: return None,pd.DataFrame()
    for mandatory in ("year","month","salesperson"):
        if not tc[mandatory]: return None,pd.DataFrame()
    x=t.copy(); x[amount]=pd.to_numeric(x[amount],errors="coerce").fillna(0)
    optional_dimensions={"principal_name","principal","channel_group","business_group","product_code","product"}
    for k,v in params.items():
        if k=="program" or k not in tc or not tc[k]:
            continue
        vals=v if isinstance(v,(list,tuple,set)) else [v]
        vals=[item for item in vals if item not in (None,"All","ALL","")]
        if not vals:
            continue
        matches=matches_dimension(x[tc[k]],k,vals)
        if k in {"year","month"}:
            # Explicit ALL in target rows applies to every selected year/month.
            wildcard=x[tc[k]].astype("string").str.strip().str.casefold().eq("all")
            matches=matches | wildcard
        if k in optional_dimensions:
            raw=x[tc[k]]
            wildcard=raw.isna() | raw.astype(str).str.strip().str.lower().isin(["","all","nan","none"])
            if k in require_explicit:
                matches=matches & ~wildcard
            else:
                matches=matches | wildcard
        x=x[matches]
    if tc["program"]:
        program_col=tc["program"]
        selected_program=params["program"]
        allowed=[selected_program,program_label(selected_program),PROGRAM_NUMBERS.get(selected_program,"")]
        if selected_program.startswith("PB "):
            allowed.append("RPB "+selected_program[3:])
        elif selected_program.startswith("RPB "):
            allowed.append("PB "+selected_program[4:])
        program_values=x[program_col].astype("string").str.strip()
        wildcard=x[program_col].isna() | program_values.str.lower().isin(["","all","nan","none"])
        allowed_keys={str(value).strip().casefold() for value in allowed}
        x=x[wildcard | program_values.str.casefold().isin(allowed_keys)]
    return float(x[amount].sum()),x

bandung_bg=ASSET_DIR/"bandung_background.svg"
bandung_header=ASSET_DIR/"bandung_header.svg"
if ACTIVE_BRANCH_CODE=="BDG":
    _branch_bg_svg=bandung_bg.read_text(encoding="utf-8") if bandung_bg.exists() else ""
    _branch_header_svg=bandung_header.read_text(encoding="utf-8") if bandung_header.exists() else ""
elif ACTIVE_BRANCH_CODE=="PTK":
    _branch_bg_svg="""<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 1600 1000'><defs><linearGradient id='g' x2='0' y2='1'><stop stop-color='#052B35'/><stop offset='1' stop-color='#087E83'/></linearGradient><radialGradient id='sun'><stop stop-color='#71D9C5' stop-opacity='.5'/><stop offset='1' stop-color='#71D9C5' stop-opacity='0'/></radialGradient></defs><rect width='1600' height='1000' fill='url(#g)'/><circle cx='1250' cy='170' r='330' fill='url(#sun)'/><path d='M0 620 Q250 560 500 620 T1000 620 T1600 620 V1000 H0Z' fill='#29A9A1' opacity='.18'/><path d='M0 700 Q250 640 500 700 T1000 700 T1600 700' fill='none' stroke='#8DE3D1' stroke-width='3' opacity='.18'/><path d='M0 770 Q250 710 500 770 T1000 770 T1600 770' fill='none' stroke='#8DE3D1' stroke-width='2' opacity='.13'/><path d='M780 575V410h40V575M760 420h80v-18h-80zM790 400v-55h20v55M750 575h100v20H750z' fill='#D4F7E8' opacity='.13'/><path d='M0 835 Q300 785 600 835 T1200 835 T1800 835' fill='none' stroke='#9FE8D6' stroke-width='2' opacity='.16'/></svg>"""
    _branch_header_svg="""<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 1600 320'><defs><linearGradient id='g' x2='1'><stop stop-color='#052B35'/><stop offset='1' stop-color='#087E83'/></linearGradient><radialGradient id='s'><stop stop-color='#71D9C5' stop-opacity='.38'/><stop offset='1' stop-color='#71D9C5' stop-opacity='0'/></radialGradient></defs><rect width='1600' height='320' fill='url(#g)'/><circle cx='1280' cy='50' r='270' fill='url(#s)'/><path d='M0 260 Q250 220 500 260 T1000 260 T1600 260 V320 H0Z' fill='#45BDB0' opacity='.3'/><path d='M0 278 Q250 238 500 278 T1000 278 T1600 278' fill='none' stroke='#B7F2E4' stroke-width='3' opacity='.55'/><path d='M810 257V133h24v124M794 146h56v-13h-56zM819 132V85h8v47M780 257h86v13h-86z' fill='#E7FAF4' opacity='.5'/></svg>"""
else:
    _branch_bg_svg="""<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 1600 1000'><defs><linearGradient id='g' x2='0' y2='1'><stop stop-color='#21142F'/><stop offset='1' stop-color='#79415F'/></linearGradient><radialGradient id='sun'><stop stop-color='#F0B969' stop-opacity='.48'/><stop offset='1' stop-color='#F0B969' stop-opacity='0'/></radialGradient></defs><rect width='1600' height='1000' fill='url(#g)'/><circle cx='1250' cy='150' r='340' fill='url(#sun)'/><path d='M0 630h280l100-65 105 65h175l95-80 95 80h210l100-90 100 90h340v370H0Z' fill='#E1B26A' opacity='.1'/><path d='M660 620V420h70v-60l95-75 95 75v60h70v200M715 420h215M750 620V460h45v160M845 620V460h45v160' fill='none' stroke='#FFE6B5' stroke-width='18' opacity='.17'/><path d='M0 720h1600M0 790h1600' stroke='#F0B969' stroke-width='2' opacity='.12'/></svg>"""
    _branch_header_svg="""<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 1600 320'><defs><linearGradient id='g' x2='1'><stop stop-color='#21142F'/><stop offset='1' stop-color='#79415F'/></linearGradient><radialGradient id='s'><stop stop-color='#F0B969' stop-opacity='.4'/><stop offset='1' stop-color='#F0B969' stop-opacity='0'/></radialGradient></defs><rect width='1600' height='320' fill='url(#g)'/><circle cx='1270' cy='50' r='270' fill='url(#s)'/><path d='M650 300V150h60v-48l90-75 90 75v48h60v150M700 155h200M748 300V185h35v115M845 300V185h35v115' fill='none' stroke='#FFE6B5' stroke-width='14' opacity='.3'/><path d='M0 272h1600M0 292h1600' stroke='#F0B969' stroke-width='3' opacity='.55'/></svg>"""
bandung_bg_data=base64.b64encode(_branch_bg_svg.encode("utf-8")).decode("ascii") if _branch_bg_svg else ""
bandung_header_data=base64.b64encode(_branch_header_svg.encode("utf-8")).decode("ascii") if _branch_header_svg else ""
logo=ASSET_DIR/"logo_aam.png"
chp_attack=ASSET_DIR/"chp_attack_logo.png"
logo_aam_img=f"<img class='brand-aam-logo' src='data:image/png;base64,{base64.b64encode(logo.read_bytes()).decode('ascii')}' alt='AAM'>" if logo.exists() else ""
attack_img=f"<img class='brand-chp-logo' src='data:image/png;base64,{base64.b64encode(chp_attack.read_bytes()).decode('ascii')}' alt='CHP ATTACK'>" if chp_attack.exists() else ""
bandung_header_img=f"<img class='brand-backdrop' src='data:image/svg+xml;base64,{bandung_header_data}' alt='' aria-hidden='true'>" if bandung_header_data else ""
st.markdown(f"<style>.stApp{{background-color:{NAVY};background-image:linear-gradient(180deg,rgba(0,7,55,.84),rgba(12,29,66,.76)),url('data:image/svg+xml;base64,{bandung_bg_data}');background-size:cover,cover;background-position:center,center bottom;background-attachment:fixed;color:{CREAM}}}.stMainBlockContainer{{background:transparent}}section[data-testid='stSidebar']{{background:rgba(0,7,55,.82)}}h1,h2,h3,h4,p,label,[data-testid='stCaptionContainer']{{color:{CREAM}}}div[data-baseweb='select']>div,input,textarea{{background-color:{MIDNIGHT_BLUE}!important;color:{CREAM}!important;border-color:{BLUE}!important}}button[kind='primary']{{background:linear-gradient(180deg,#33509B 0%,#25407F 60%,#1B3064 100%);border:1px solid #45629F;color:{CREAM};box-shadow:0 4px 12px rgba(0,7,55,.45),inset 0 1px 0 rgba(255,255,255,.22);transition:transform .15s ease,box-shadow .15s ease}}button[kind='primary']:hover{{transform:translateY(-1px);box-shadow:0 7px 16px rgba(0,7,55,.55),inset 0 1px 0 rgba(255,255,255,.28)}}button[kind='secondary']{{background:linear-gradient(180deg,#1B3064 0%,#142650 100%);border:1px solid #34508F;color:{CREAM};box-shadow:0 3px 10px rgba(0,7,55,.35),inset 0 1px 0 rgba(255,255,255,.14)}}button[kind='secondary']:hover{{border-color:#45629F;color:{CREAM}}}.brand{{background:rgba(0,7,55,.96);color:{CREAM};padding:10px 22px;border-radius:16px;border-left:7px solid {BLUE};border-right:2px solid rgba(252,241,207,.24);margin-bottom:12px;display:flex;align-items:center;gap:24px;min-height:112px;box-shadow:0 12px 34px rgba(0,7,55,.30)}}.brand-aam-logo{{width:105px;height:82px;object-fit:contain;object-position:center;flex:0 0 auto}}.brand-copy{{flex:1;min-width:0}}.brand h1{{margin:0 0 7px 0}}.brand-chp-logo{{width:150px;height:86px;object-fit:contain;object-position:center;flex:0 0 auto}}.brand small{{color:{CREAM}}}.program{{border-top:4px solid {RED};padding-top:8px}}div[data-testid='stMetric']{{background:linear-gradient(160deg,#16305F 0%,{MIDNIGHT_BLUE} 55%,#081634 100%);color:{CREAM};padding:14px;border-radius:12px;border:1px solid #34508F;box-shadow:0 8px 20px rgba(0,7,55,.38),inset 0 1px 0 rgba(255,255,255,.12)}}.scorecard{{background:linear-gradient(165deg,#16305F 0%,{MIDNIGHT_BLUE} 50%,#081634 100%);color:{CREAM};border:1px solid #34508F;border-radius:17px;padding:11px 13px;min-height:315px;box-shadow:0 12px 26px rgba(0,7,55,.45),inset 0 1px 0 rgba(255,255,255,.12);font-family:inherit;transition:transform .15s ease,box-shadow .15s ease}}.scorecard:hover{{transform:translateY(-2px);box-shadow:0 16px 32px rgba(0,7,55,.55),inset 0 1px 0 rgba(255,255,255,.16)}}.score-program{{font-size:.66rem;letter-spacing:.08em;color:{CREAM};font-weight:800;margin-bottom:5px}}.score-name{{font-size:1.03rem;font-weight:850;line-height:1.12;margin-bottom:8px}}.score-filter{{font-size:.66rem;color:rgba(252,241,207,.72);line-height:1.28;min-height:3.5em;margin-bottom:8px;overflow-wrap:anywhere}}.score-metric{{font-size:.72rem;margin:7px 0;color:{CREAM}}}.score-row{{font-size:.82rem;line-height:1.45;margin:2px 0}}.score-ach{{font-size:.72rem;margin-top:8px}}.score-ach strong{{font-size:1.26rem}}.weekly-plot{{margin:5px 0 7px;font-size:.65rem}}.weekly-plot-label{{margin-bottom:3px;color:rgba(252,241,207,.82)}}.weekly-plot-track{{height:8px;background:rgba(0,7,55,.72);border:1px solid rgba(69,98,159,.75);border-radius:999px;overflow:hidden}}.weekly-plot-fill{{height:100%;border-radius:999px;transition:width .25s ease}}.weekly-plot-value{{text-align:right;font-weight:800;margin-top:2px}}.score-gap{{font-size:.72rem;color:{CREAM};margin-top:7px;border-top:1px solid {BLUE};padding-top:7px}}.score-insight{{font-size:.64rem;color:rgba(252,241,207,.92);line-height:1.2;margin-top:5px;border-top:1px solid {BLUE};padding-top:4px}}.score-insight.insight-low{{color:#ff5368}}.score-insight.insight-medium{{color:#F5C542}}.score-insight.insight-good{{color:#39C77B}}.score-insight .negative-contribution{{color:inherit}}.score-insight .on-track-contribution{{color:#39C77B;font-weight:850}}.insight-low{{border-left:4px solid #8FA8DA;background:rgba(143,168,218,.16);padding:8px 9px;border-radius:0 9px 9px 0}}.insight-medium{{border-left:4px solid #F5C542;background:rgba(245,197,66,.14);padding:8px 9px;border-radius:0 9px 9px 0}}.insight-good{{border-left:4px solid #39C77B;background:rgba(57,199,123,.14);padding:8px 9px;border-radius:0 9px 9px 0}}.insight-status{{display:inline-block;padding:1px 5px;border-radius:999px;font-size:.54rem;font-weight:900;letter-spacing:.04em;margin-left:4px;vertical-align:1px}}.insight-status-low{{background:#8FA8DA;color:{NAVY}}}.insight-status-medium{{background:#F5C542;color:#000737}}.insight-status-good{{background:#39C77B;color:#000737}}.negative-contribution{{color:#F1C66D;font-weight:800}}.bandung-badge{{display:inline-block;background:{BLUE};color:{CREAM};padding:3px 9px;border-radius:999px;font-size:.72rem;font-weight:800;letter-spacing:.08em}}.score-heading{{font-size:1.15rem;font-weight:900;color:{CREAM};letter-spacing:.04em;padding:10px 14px;background:linear-gradient(180deg,#33509B 0%,#22396F 55%,#182A52 100%);border:1px solid #45629F;border-radius:12px;margin-bottom:8px;white-space:nowrap;text-shadow:0 1px 2px rgba(0,0,0,.45);box-shadow:0 8px 18px rgba(0,7,55,.40),inset 0 1px 0 rgba(255,255,255,.25),inset 0 -2px 0 rgba(0,0,0,.30)}}@media(max-width:700px){{.brand{{gap:10px;padding:10px}}.brand-aam-logo{{width:68px;height:60px}}.brand-chp-logo{{width:88px;height:58px}}.brand h1{{font-size:1.15rem}}.brand small{{font-size:.68rem}}}}</style>",unsafe_allow_html=True)
st.markdown(f"""
<style>
.brand {{
  position: relative;
  isolation: isolate;
  overflow: hidden;
  min-height: 138px;
  padding: 12px 24px 13px;
  border: 1px solid rgba(174, 198, 238, .32);
  border-left: 7px solid {BLUE};
  border-bottom: 3px solid rgba(232, 182, 80, .72);
  border-radius: 18px;
  background: linear-gradient(110deg, rgba(0,7,55,.98), rgba(9,26,66,.96) 55%, rgba(21,44,87,.96));
  box-shadow: 0 14px 34px rgba(0,7,55,.38), inset 0 1px 0 rgba(255,255,255,.13);
}}
.brand-backdrop {{
  position: absolute;
  z-index: 0;
  inset: 0;
  width: 100%;
  height: 100%;
  object-fit: cover;
  object-position: center 63%;
  opacity: .96;
  pointer-events: none;
}}
.brand-aam-logo, .brand-chp-logo {{ position: relative; z-index: 1; filter: drop-shadow(0 3px 7px rgba(0,0,0,.35)); }}
.brand-copy {{ position: relative; z-index: 1; align-self: center; }}
.brand h1 {{ margin: 0 0 5px !important; font-size: clamp(1.55rem, 2.4vw, 2.35rem); letter-spacing: .045em; text-shadow: 0 2px 12px rgba(0,0,0,.4); }}
.brand-caption {{ display: flex; align-items: center; gap: 11px; flex-wrap: wrap; color: {CREAM}; font-size: .82rem; line-height: 1.35; }}
.brand-caption .bandung-badge {{ background: linear-gradient(135deg, #294A89, #183469); border: 1px solid rgba(247,217,144,.55); color: #FFE9AA; box-shadow: 0 3px 10px rgba(0,0,0,.24); }}
.brand-landmark-label {{ margin-top: 7px; color: rgba(247,217,144,.9); font-size: .62rem; font-weight: 850; letter-spacing: .22em; text-shadow: 0 1px 4px rgba(0,0,0,.6); }}
.brand-landmark-label i {{ color: #D94B48; font-style: normal; padding: 0 4px; }}
@media (max-width: 700px) {{
  .brand {{ min-height: 118px; padding: 10px 12px; gap: 10px; }}
  .brand-aam-logo {{ width: 62px !important; height: 58px !important; }}
  .brand-chp-logo {{ width: 80px !important; height: 58px !important; }}
  .brand-caption {{ font-size: .68rem; gap: 6px; }}
  .brand-landmark-label {{ font-size: .52rem; letter-spacing: .1em; }}
}}
</style>
""", unsafe_allow_html=True)
st.markdown("<style>header[data-testid='stHeader']{height:1.3rem!important;min-height:1.3rem!important;background:transparent!important}header[data-testid='stHeader'] [data-testid='stToolbar']{height:1.3rem!important;min-height:1.3rem!important;transform:scale(.72);transform-origin:top right;padding:0!important}[data-testid='stDecoration']{display:none!important}.stMainBlockContainer{padding-top:.65rem!important}div[data-testid='stDataFrame'],div[data-testid='stDataEditor']{--gdg-text-dark:#1B2E46;--gdg-text-medium:#1B2E46;--gdg-text-header:#FFFFFF;--gdg-bg-cell:#F4F7FB;--gdg-bg-cell-medium:#E8EEF6;--gdg-bg-header:#335A9F;--gdg-bg-header-hover:#24477F;border-radius:14px;overflow:hidden;box-shadow:0 10px 24px rgba(0,7,55,.35)}.table-scroll-wrap table{margin:0}.table-scroll-wrap th{position:sticky;top:0;z-index:3;padding:3px 3px!important;font-size:.82rem!important;line-height:1.1!important}.table-scroll-wrap td{padding:2px 3px!important;font-size:.82rem!important;line-height:1.1!important}div[data-testid='stTable'] table{border-radius:14px;overflow:hidden;border-collapse:separate;border-spacing:0;box-shadow:0 10px 24px rgba(0,7,55,.35)}div[data-testid='stTable'] th,div[data-testid='stTable'] th p{background:linear-gradient(180deg,#335A9F 0%,#24477F 100%)!important;color:#FFFFFF!important;font-weight:800!important;letter-spacing:.03em;text-shadow:0 1px 2px rgba(0,0,0,.35);box-shadow:inset 0 1px 0 rgba(255,255,255,.22),inset 0 -2px 0 rgba(0,0,0,.24);border-bottom:2px solid #1D3B70;padding:3px 3px!important;font-size:.82rem!important;line-height:1.1!important}div[data-testid='stTable'] td,div[data-testid='stTable'] td p{color:#1B2E46!important;background:#F4F7FB!important;padding:2px 3px!important;font-size:.82rem!important;line-height:1.1!important}div[data-testid='stTable'] tbody tr:nth-child(even) td{background:#E8EEF6!important}</style>",unsafe_allow_html=True)
st.markdown(f"""
<style>
/* Safari exports the rendered document; let it see the full scroll height.
   Keep the existing app backgrounds and colors unchanged. */
@supports (-webkit-hyphens: none) {{
  html, body, #root, [data-testid="stApp"], .stApp,
  [data-testid="stAppViewContainer"], [data-testid="stMain"],
  [data-testid="stMainBlockContainer"], .stMainBlockContainer {{
    height: auto !important;
    min-height: 100vh !important;
    max-height: none !important;
    overflow: visible !important;
  }}
  body {{ background-color: {NAVY} !important; overflow-y: auto !important; }}
  [data-testid="stAppViewContainer"], [data-testid="stMain"] {{
    flex: 1 0 auto !important;
  }}
}}
</style>
""", unsafe_allow_html=True)
st.markdown("""
<style>
@media print {
  @page { margin: 8mm; }
  .brand { -webkit-print-color-adjust: exact !important; print-color-adjust: exact !important; }
  .brand-backdrop { display: block !important; visibility: visible !important; opacity: 1 !important; }
  html, body, #root, [data-testid="stApp"], .stApp,
  [data-testid="stAppViewContainer"], [data-testid="stMain"],
  [data-testid="stMainBlockContainer"], .stMainBlockContainer,
  section.main, main {
    height: auto !important;
    min-height: 0 !important;
    max-height: none !important;
    overflow: visible !important;
    position: static !important;
    contain: none !important;
  }
  body { overflow: visible !important; }
  html, body, #root, [data-testid="stApp"], .stApp,
  [data-testid="stAppViewContainer"], [data-testid="stMain"] {
    background-color: #000737 !important;
  }
  body, #root, [data-testid="stApp"], .stApp {
    background-image: linear-gradient(180deg,rgba(0,7,55,.84),rgba(12,29,66,.76)),
      url('data:image/svg+xml;base64,__BANDUNG_BG__') !important;
    background-size: cover, cover !important;
    background-position: center, center bottom !important;
    background-repeat: no-repeat !important;
    background-attachment: scroll !important;
    color: #FCF1CF !important;
  }
  [data-testid="stAppViewContainer"], [data-testid="stMain"] {
    background-image: none !important;
  }
  [data-testid="stAppViewContainer"], [data-testid="stMain"] {
    display: block !important;
  }
  [data-testid="stVerticalBlock"], [data-testid="stElementContainer"],
  [data-testid="stHorizontalBlock"], [data-testid="stMainBlockContainer"],
  .table-scroll-wrap, div[data-testid="stDataFrame"],
  div[data-testid="stTable"] {
    height: auto !important;
    max-height: none !important;
    overflow: visible !important;
    transform: none !important;
  }
  .table-scroll-wrap table { position: static !important; }
  .table-scroll-wrap th { position: static !important; }
  section[data-testid="stSidebar"], header[data-testid="stHeader"],
  [data-testid="stToolbar"], [data-testid="stDecoration"] {
    display: none !important;
  }
  * {
    -webkit-print-color-adjust: exact !important;
    print-color-adjust: exact !important;
  }
  .scorecard, .brand, img, svg, tr {
    break-inside: avoid;
    page-break-inside: avoid;
  }
}
</style>
""".replace("__BANDUNG_BG__", bandung_bg_data), unsafe_allow_html=True)
st.markdown(f"<div class='brand'>{bandung_header_img}{logo_aam_img}<div class='brand-copy'><h1>CHP AAM</h1><div class='brand-caption'><span class='bandung-badge'>CABANG {ACTIVE_BRANCH_NAME.upper()} · {ACTIVE_BRANCH_CODE}</span><span>PUSH PERFORMANCE · Monitoring actual, target, achievement, dan insight per program</span></div><div class='brand-landmark-label'>{_theme['landmarks'].replace(' · ',' <i>·</i> ')}</div></div>{attack_img}</div>",unsafe_allow_html=True)
if "dashboard_page" not in st.session_state:
    st.session_state["dashboard_page"]=settings.get("dashboard_page","Summary Program")
if st.session_state["dashboard_page"]=="Sasaran per Sales Person Title":
    st.session_state["dashboard_page"]="Sasaran Program"
if "dashboard_section" not in st.session_state:
    st.session_state["dashboard_section"]=settings.get("dashboard_section","Program Push")
if "nav_minimized" not in st.session_state:
    st.session_state["nav_minimized"]=False
if st.session_state["nav_minimized"]:
    if st.button("☰ Tampilkan navigasi",key="show_dashboard_navigation"):
        st.session_state["nav_minimized"]=False
        st.rerun()
    main_section=st.session_state.get("dashboard_section",settings.get("dashboard_section","Program Push"))
    page=st.session_state.get("dashboard_page",settings.get("dashboard_page","Summary Program")) if main_section=="Program Push" else None
else:
    nav_cols=st.columns([7,1])
    with nav_cols[0]:
        main_section=st.radio("Dashboard CHP AAM",["Performance Sales","Program Push","Menu Insentif","Report"],horizontal=True,key="dashboard_section")
    with nav_cols[1]:
        if st.button("− Minimize",key="minimize_dashboard_navigation",help="Sembunyikan navigasi utama."):
            st.session_state["nav_minimized"]=True
            st.rerun()
    if main_section=="Program Push":
        page=st.radio("Pilih tampilan Program Push",["Summary Program","Program by Sales Person Title","Sasaran Program","Sasaran Kunjungan Harian per Sales Person Title"],horizontal=True,key="dashboard_page")
    else:
        page=None
settings["dashboard_section"]=main_section

# ── TAB BULAN untuk Program Push ─────────────────────────────────────────────
if main_section=="Program Push":
    _tab_months   = list(MONTHS.values())
    _tab_abbr     = ["JAN","FEB","MAR","APR","MAY","JUN","JUL","AUG","SEP","OCT","NOV","DEC"]
    _month_to_num = {v:k for k,v in MONTHS.items()}

    # Bulan aktif tab
    _now_month = pd.Timestamp.now().month
    _saved_tab = st.session_state.get("push_month_tab_idx", settings.get("push_month_tab_idx", _now_month-1))
    try:
        _saved_tab = int(_saved_tab)
    except (TypeError, ValueError):
        _saved_tab = _now_month-1
    if not 0 <= _saved_tab < 12:
        _saved_tab = _now_month-1
    st.session_state.setdefault("push_month_tab_idx", _saved_tab)

    _active_push_month     = _tab_months[_saved_tab]
    _active_push_month_num = _saved_tab + 1
    _pbm_key = _active_push_month   # key untuk per-bulan settings

    # Minimize hides all Program Push controls while preserving the active month.
    if not st.session_state.get("nav_minimized", False):
        _tab_cols = st.columns(12, gap="small")
        for _ti, (_tcol, _tabbr) in enumerate(zip(_tab_cols, _tab_abbr)):
            with _tcol:
                _is_active = (_ti == _saved_tab)
                if st.button(_tabbr, key=f"tab_month_{_ti}", width="stretch",
                             type="primary" if _is_active else "secondary"):
                    st.session_state["push_month_tab_idx"] = _ti
                    settings["push_month_tab_idx"] = _ti
                    save_settings(settings)
                    st.rerun()

    # ── Load konfigurasi per bulan ────────────────────────────────────────
    _pbm_cfg = settings.setdefault("program_config_by_month", {})
    _this_month_cfg = _pbm_cfg.setdefault(_pbm_key, {})

    # Helper: ambil config bulan ini, fallback ke PROGRAM_FILTER_DEFAULTS
    def _prog_cfg(prog, kind, fallback=None):
        v = _this_month_cfg.get(prog, {}).get(kind, None)
        if v is None and fallback is not None: return fallback
        if v is None: return PROGRAM_FILTER_DEFAULTS.get(prog, {}).get(kind, "")
        return v

    # _cfg_save dibaca dari settings (expander ada di bawah setelah df dimuat)
    _cfg_save = {}
    for _p in PROGRAMS:
        _d = dict(_this_month_cfg.get(_p, {}))
        _d.setdefault("display_name", _p)
        _d.setdefault("principal",   PROGRAM_FILTER_DEFAULTS.get(_p,{}).get("principal","All") or "All")
        _d.setdefault("product_code",PROGRAM_FILTER_DEFAULTS.get(_p,{}).get("product_code","All") or "All")
        _d.setdefault("product",     PROGRAM_FILTER_DEFAULTS.get(_p,{}).get("product","All") or "All")
        _d.setdefault("channel_group",PROGRAM_FILTER_DEFAULTS.get(_p,{}).get("channel_group","All") or "All")
        _d.setdefault("business_group",PROGRAM_FILTER_DEFAULTS.get(_p,{}).get("business_group","All") or "All")
        _d.setdefault("metric",       PROGRAM_FILTER_DEFAULTS.get(_p,{}).get("metric","Total Relasi Beli Transaksi"))
        _saved_year = settings.get("filters",{}).get(_p,{}).get("year", "All")
        if isinstance(_saved_year, (list, tuple)):
            _saved_year = _saved_year[0] if _saved_year else "All"
        _d.setdefault("year", str(_saved_year or "All"))
        _d.setdefault("calculation_month", _active_push_month)
        _cfg_save[_p] = _d
        _this_month_cfg[_p] = _d

    # Terapkan konfigurasi bulan aktif ke session_state untuk tampilan detail.
    for _prog in PROGRAMS:
        _pkey = key(_prog).replace(" ","_")
        _month_val = _cfg_save[_prog].get("calculation_month", _active_push_month) or _active_push_month
        st.session_state[f"filter_month_{_pkey}"] = [_month_val]
        _pid_val  = _cfg_save[_prog].get("principal","All") or "All"
        st.session_state[f"filter_principal_{_pkey}"] = [] if _pid_val=="All" else [_pid_val]
        _pc_val   = _cfg_save[_prog].get("product_code","All") or "All"
        st.session_state[f"filter_product_code_{_pkey}"] = [] if _pc_val=="All" else [_pc_val]
        _prod_val = _cfg_save[_prog].get("product","All") or "All"
        st.session_state[f"filter_product_{_pkey}"] = [] if _prod_val=="All" else [_prod_val]
        _channel_val = _cfg_save[_prog].get("channel_group","All") or "All"
        st.session_state[f"filter_channel_group_{_pkey}"] = [] if _channel_val=="All" else [_channel_val]
        _business_group_val = _cfg_save[_prog].get("business_group","All") or "All"
        st.session_state[f"filter_business_group_{_pkey}"] = [] if _business_group_val=="All" else [_business_group_val]
        _year_val = _cfg_save[_prog].get("year","All") or "All"
        st.session_state[f"filter_year_{_pkey}"] = [] if _year_val=="All" else [str(_year_val)]
        st.session_state[f"filter_metric_{_pkey}"] = _cfg_save[_prog].get("metric","Total Relasi Beli Transaksi")

    if not st.session_state.get("nav_minimized", False):
        st.caption(f"📅 Konfigurasi yang sedang dibuka: **{_active_push_month}**. Bulan transaksi dihitung sesuai pilihan tiap program.")

with st.sidebar:
    st.header("Data dashboard")
    default_db=Path(__file__).parent/"database_push.xlsx"
    try:
        _drive_settings=st.secrets.get("google_drive",{})
        _gcp_account=st.secrets.get("gcp_service_account",{})
        _drive_file_id=str(_drive_settings.get("database_file_id","")).strip()
        _drive_database_folder_id=str(_drive_settings.get("database_folder_id",DEFAULT_DRIVE_DATABASE_FOLDER_ID)).strip()
        _drive_is_folder=not bool(_drive_file_id)
        _drive_source_id=_drive_file_id or _drive_database_folder_id
        _drive_configured=bool(_drive_source_id and _gcp_account)
    except Exception:
        _drive_settings={}; _gcp_account={}; _drive_configured=False
    # Drive data is divided into monthly branch spreadsheets so each file stays
    # small enough to open and update in Google Sheets.
    _data_source_mode="Google Drive · spreadsheet bulanan" if _drive_configured else "Lokal / unggah"
    source=None
    _drive_database_path=None
    if _data_source_mode.startswith("Google Drive"):
        st.caption(f"Sumber data: spreadsheet bulanan cabang {ACTIVE_BRANCH_NAME} ({ACTIVE_BRANCH_CODE}) di Google Drive.")
        st.caption("Untuk update harian, buka file bulan berjalan di folder database_monthly_split, tambahkan/perbarui data, lalu gunakan tombol periksa pembaruan.")
        _drive_database_path=APP_DIR/".remote_database"/"database_push.xlsx"
        if st.button("↻ Periksa pembaruan dari Drive",key="refresh_drive_database",width="stretch"):
            # Refresh Drive metadata/listing only. The normalized DataFrame cache is
            # keyed by each file's modifiedTime and size, so a changed workbook
            # naturally gets a new cache key. Clearing all resource caches here
            # duplicates large branch data in memory and can crash Streamlit Cloud.
            st.cache_data.clear()
            st.rerun()
    else:
        source=st.file_uploader("Ganti database untuk sesi ini (opsional)",type=["xlsx","xls"],key="db")
    if _drive_configured:
        st.caption(f"Cabang aktif: {ACTIVE_BRANCH_NAME} ({ACTIVE_BRANCH_CODE}). Data diambil dari spreadsheet per bulan agar mudah diperbarui.")
    elif default_db.exists():
        st.caption(f"Cabang aktif: {ACTIVE_BRANCH_NAME} ({ACTIVE_BRANCH_CODE}). Workbook dapat berisi sheet BDG, PTK, dan MDN; dashboard hanya membaca sheet cabang aktif.")
    else:
        st.caption("Letakkan database_push.xlsx di folder aplikasi, atau unggah file di sini.")
    if not _drive_configured and st.button("↻ Refresh database",width="stretch"):
        st.cache_data.clear()
        st.cache_resource.clear()
        st.session_state.pop("_base_filter_option_cache",None)
        st.session_state.pop("_sasaran_positive_pool_cache",None)
        st.rerun()
    calendar_path=OVERVIEW_DATA_DIR/"overview_calendar.xlsx"
    fixed_call_path=OVERVIEW_DATA_DIR/"overview_fixed_call.xlsx"
    with st.expander("File sasaran kunjungan harian"):
        st.caption("Dashboard memakai tanggal hari ini dari Overview Calendar dan rute hari/week dari Overview Fixed Call.")
        for upload_label,file_key,path in [
            ("Perbarui Overview Calendar","calendar_upload",calendar_path),
            ("Perbarui Overview Fixed Call","fixed_call_upload",fixed_call_path),
        ]:
            uploaded=st.file_uploader(upload_label,type=["xlsx"],key=file_key)
            if uploaded is not None:
                content=uploaded.getvalue()
                fingerprint=hashlib.sha256(content).hexdigest()
                state_key=f"{file_key}_fingerprint"
                if st.session_state.get(state_key)!=fingerprint:
                    path.write_bytes(content)
                    if st.secrets.get("gcp_service_account"):
                        try: _drive_upload_bytes(path.name,content,"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet","overview",ACTIVE_BRANCH_CODE)
                        except Exception as exc: st.error(f"File diperbarui sementara, tetapi gagal disinkronkan ke Google Drive: {exc}")
                    st.session_state[state_key]=fingerprint
                    st.success(f"{path.name} diperbarui.")
            st.caption(f"File aktif: {path.name}" if path.exists() else f"File belum tersedia: {path.name}")
    # Keep the active selection aligned with the persisted program list when
    # that list changes in settings (for example after restoring all programs).
    _saved_programs = [p for p in settings.get("programs", PROGRAMS) if p in PROGRAMS]
    _programs_signature = tuple(_saved_programs)
    if st.session_state.get("_programs_signature") != _programs_signature:
        st.session_state["programs"] = _saved_programs or PROGRAMS
        st.session_state["_programs_signature"] = _programs_signature
    elif "programs" not in st.session_state:
        saved_programs=settings.get("programs",PROGRAMS)
        st.session_state["programs"]=[p for p in saved_programs if p in PROGRAMS] or PROGRAMS
    programs=st.multiselect("Pilih program Push",PROGRAMS,key="programs",format_func=program_label)

@st.cache_data(show_spinner="Membaca database…")
def read_local_file(path,modified_ns,file_size,sheet_name=0):
    p=Path(path)
    return pd.read_csv(p) if p.suffix.lower()==".csv" else pd.read_excel(p,sheet_name=sheet_name)

@st.cache_data(ttl=60, show_spinner=False)
def _google_drive_metadata(source_id, account_fingerprint, is_folder=False):
    """Check the Drive version at most once per minute for normal app reruns."""
    from google.oauth2.service_account import Credentials
    from google.auth.transport.requests import AuthorizedSession
    service_account_info=dict(st.secrets["gcp_service_account"])
    credentials=Credentials.from_service_account_info(
        service_account_info,
        scopes=["https://www.googleapis.com/auth/drive.readonly"],
    )
    session=AuthorizedSession(credentials)
    if is_folder:
        response=session.get(
            "https://www.googleapis.com/drive/v3/files",
            params={
                "q":f"'{source_id}' in parents and trashed = false",
                "fields":"files(id,name,mimeType,size,modifiedTime)",
                "pageSize":100,"orderBy":"modifiedTime desc",
                "supportsAllDrives":"true","includeItemsFromAllDrives":"true",
            },timeout=30,
        )
        response.raise_for_status()
        files=response.json().get("files",[])
        excel_files=[item for item in files if str(item.get("name","")).lower().endswith((".xlsx",".xls"))]
        if not excel_files:
            raise FileNotFoundError("Tidak ada file Excel (.xlsx/.xls) di folder Google Drive tersebut.")
        return next(
            (item for item in excel_files if item.get("name","").lower()=="database_push_branches.xlsx"),
            next((item for item in excel_files if item.get("name","").lower()=="database_push.xlsx"),excel_files[0]),
        )
    response=session.get(
        f"https://www.googleapis.com/drive/v3/files/{source_id}",
        params={"fields":"id,name,mimeType,size,modifiedTime","supportsAllDrives":"true"},timeout=30,
    )
    response.raise_for_status()
    return response.json()

@st.cache_data(ttl=60, show_spinner=False)
def _google_drive_monthly_files(root_folder_id, branch_code, account_fingerprint):
    """Find the active branch's monthly files under database_monthly_split."""
    from google.oauth2.service_account import Credentials
    from google.auth.transport.requests import AuthorizedSession
    service_account_info=dict(st.secrets["gcp_service_account"])
    credentials=Credentials.from_service_account_info(
        service_account_info,scopes=["https://www.googleapis.com/auth/drive.readonly"]
    )
    session=AuthorizedSession(credentials)
    def children(parent_id):
        result=[]; page_token=None
        while True:
            params={
                "q":f"'{parent_id}' in parents and trashed = false",
                "fields":"nextPageToken,files(id,name,mimeType,size,modifiedTime)",
                "pageSize":1000,"supportsAllDrives":"true","includeItemsFromAllDrives":"true",
            }
            if page_token: params["pageToken"]=page_token
            response=session.get("https://www.googleapis.com/drive/v3/files",params=params,timeout=30)
            response.raise_for_status()
            payload=response.json(); result.extend(payload.get("files",[]))
            page_token=payload.get("nextPageToken")
            if not page_token: return result
    root_items=children(root_folder_id)
    split_folder=next((item for item in root_items if item.get("mimeType")=="application/vnd.google-apps.folder" and key(item.get("name",""))=="database monthly split"),None)
    if not split_folder:
        raise FileNotFoundError("Folder database_monthly_split belum ditemukan di folder Drive utama.")
    branch_items=children(split_folder["id"])
    branch_folder=next((item for item in branch_items if item.get("mimeType")=="application/vnd.google-apps.folder" and item.get("name","").strip().upper()==branch_code),None)
    if not branch_folder:
        raise FileNotFoundError(f"Folder cabang {branch_code} belum ditemukan di database_monthly_split.")
    files=children(branch_folder["id"])
    candidates={}
    for item in files:
        name=str(item.get("name",""))
        upper_name=name.upper()
        if upper_name.startswith(f"DATABASE_PUSH_{branch_code}_") and name.lower().endswith((".xlsx", ".csv")):
            period=Path(name).stem.upper().replace(f"DATABASE_PUSH_{branch_code}_", "")
            candidates.setdefault(period,{})[Path(name).suffix.lower()]=item
    # Read legacy months from their fast CSV copies and only open the current
    # month's editable Excel workbook. Loading dozens of large XLSX files on
    # Streamlit Cloud kept new sessions stuck on the loading screen.
    periods=sorted(candidates)
    current_period=periods[-1] if periods else None
    monthly_by_period={}
    for period,by_extension in candidates.items():
        preferred=".xlsx" if period==current_period else ".csv"
        fallback=".csv" if preferred==".xlsx" else ".xlsx"
        item=by_extension.get(preferred) or by_extension.get(fallback)
        if item:
            monthly_by_period[period]=item
    return sorted(monthly_by_period.values(),key=lambda item:item.get("name",""))

def _download_drive_monthly_files(files, destination, service_account_info):
    """Download changed month files one at a time and keep a local reusable copy."""
    from google.oauth2.service_account import Credentials
    from google.auth.transport.requests import AuthorizedSession
    credentials=Credentials.from_service_account_info(
        dict(service_account_info),scopes=["https://www.googleapis.com/auth/drive.readonly"]
    )
    session=AuthorizedSession(credentials)
    destination=Path(destination); destination.mkdir(parents=True,exist_ok=True)
    paths=[]; signatures=[]
    for item in files:
        name=Path(str(item["name"])).name
        local_path=destination/name
        version=f"{item.get('modifiedTime','')}:{item.get('size','')}"
        marker=local_path.with_suffix(local_path.suffix+".version")
        if not (local_path.exists() and marker.exists() and marker.read_text(encoding="utf-8")==version):
            response=session.get(
                f"https://www.googleapis.com/drive/v3/files/{item['id']}",
                params={"alt":"media","supportsAllDrives":"true"},timeout=(30,900),
            )
            response.raise_for_status()
            payload=response.content
            is_xlsx=name.lower().endswith(".xlsx")
            if is_xlsx and not payload.startswith(b"PK"):
                raise ValueError(f"File {name} bukan workbook Excel yang valid atau belum selesai diunggah.")
            if not is_xlsx and not payload.startswith((b"\xef\xbb\xbf",b"Product Group",b"\"Product Group")):
                raise ValueError(f"File {name} bukan spreadsheet/CSV yang valid atau belum selesai diunggah.")
            temporary=local_path.with_suffix(local_path.suffix+".tmp")
            temporary.write_bytes(payload); temporary.replace(local_path)
            marker.write_text(version,encoding="utf-8")
        paths.append(str(local_path))
        signatures.append(f"{name}:{version}")
    return paths,"|".join(signatures)

@st.cache_resource(max_entries=1,show_spinner="Menyiapkan data transaksi cabang…")
def read_normalized_monthly_branch(branch_code, paths, signature):
    """Load monthly files with shared categorical columns to stay within Cloud memory."""
    cache_dir=APP_DIR/".remote_database"/"normalized"
    cache_dir.mkdir(parents=True,exist_ok=True)
    fingerprint=hashlib.sha256(signature.encode("utf-8")).hexdigest()[:20]
    parquet_path=cache_dir/f"{branch_code}_{fingerprint}.parquet"
    if parquet_path.exists():
        try: return pd.read_parquet(parquet_path)
        except Exception: parquet_path.unlink(missing_ok=True)

    # Find text columns and their shared categories in a low-memory first pass.
    csv_paths=[Path(path) for path in paths if Path(path).suffix.lower()==".csv"]
    text_columns=[]; categories={}
    if csv_paths:
        sample=pd.read_csv(csv_paths[0],nrows=1000,low_memory=False)
        dimension_names={key(alias) for kind,aliases in ALIASES.items() if kind!="sales" for alias in aliases}
        text_columns=[name for name in sample.columns if not pd.api.types.is_numeric_dtype(sample[name].dtype) or key(name) in dimension_names]
        del sample
        categories={name:set() for name in text_columns}
        for path in csv_paths:
            for chunk in pd.read_csv(path,usecols=lambda name:name in categories,dtype="string",chunksize=50000,low_memory=False):
                for name in text_columns:
                    if name in chunk:
                        categories[name].update(chunk[name].dropna().str.strip().unique().tolist())
        for path in paths:
            if Path(path).suffix.lower()==".xlsx":
                excel_dims=pd.read_excel(path,engine="openpyxl",usecols=lambda name:name in categories)
                for name in text_columns:
                    if name in excel_dims:
                        categories[name].update(excel_dims[name].dropna().astype("string").str.strip().unique().tolist())
                del excel_dims
    shared_dtypes={name:pd.CategoricalDtype(categories=sorted(values)) for name,values in categories.items()}

    pieces=[]
    for path in paths:
        if Path(path).suffix.lower()==".xlsx":
            piece=pd.read_excel(path,engine="openpyxl")
            for name,dtype in shared_dtypes.items():
                if name in piece: piece[name]=piece[name].astype("string").str.strip().astype(dtype)
        else:
            piece=pd.read_csv(path,low_memory=False,dtype=shared_dtypes)
        if not piece.empty: pieces.append(piece)
    if not pieces:
        return pd.DataFrame()
    normalized=normalize(pd.concat(pieces,ignore_index=True,sort=False),copy=False)
    # Keep normalized text as categorical; expanding it to a Python string per row
    # was the main cause of Streamlit Cloud exceeding its memory limit.
    for name in text_columns:
        if name in normalized and not isinstance(normalized[name].dtype,pd.CategoricalDtype):
            normalized[name]=normalized[name].astype("category")
    temporary=parquet_path.with_suffix(".parquet.tmp")
    try:
        normalized.to_parquet(temporary,index=False); temporary.replace(parquet_path)
    except Exception:
        temporary.unlink(missing_ok=True)
    return normalized

def _google_drive_database_path(source_id, destination, service_account_info, is_folder=False):
    """Download the configured shared Excel workbook only when Drive says it changed."""
    try:
        from google.oauth2.service_account import Credentials
        from google.auth.transport.requests import AuthorizedSession
        credentials=Credentials.from_service_account_info(
            dict(service_account_info),
            scopes=["https://www.googleapis.com/auth/drive.readonly"],
        )
        session=AuthorizedSession(credentials)
        _account_fingerprint=hashlib.sha256(json.dumps(dict(service_account_info),sort_keys=True).encode("utf-8")).hexdigest()
        metadata=_google_drive_metadata(str(source_id),_account_fingerprint,is_folder)
        file_id=metadata["id"]
        metadata_url=f"https://www.googleapis.com/drive/v3/files/{file_id}"
        if metadata.get("mimeType")=="application/vnd.google-apps.spreadsheet":
            raise ValueError("File di Drive berupa Google Sheets. Simpan/unggah database sebagai Excel (.xlsx) agar seluruh data besar dapat dibaca tanpa batas sel spreadsheet.")
        version=f"{file_id}:{metadata.get('modifiedTime','')}:{metadata.get('size','')}"
        marker=destination.with_suffix(destination.suffix+".version")
        if destination.exists() and marker.exists() and marker.read_text(encoding="utf-8")==version:
            return destination
        response=session.get(
            metadata_url,params={"alt":"media","supportsAllDrives":"true"},timeout=(30,900)
        )
        if response.status_code==403:
            raise PermissionError(
                "Google Drive menolak akses ke workbook database. Pastikan file database_push_branches.xlsx "
                "dibagikan kepada service account dashboard dengan akses Viewer, lalu klik 'Periksa pembaruan dari Drive'."
            )
        response.raise_for_status()
        content=response.content
        if not content.startswith(b"PK"):
            raise ValueError("Isi file Drive bukan workbook Excel yang valid.")
        destination.parent.mkdir(parents=True,exist_ok=True)
        temporary=destination.with_suffix(destination.suffix+".tmp")
        temporary.write_bytes(content)
        temporary.replace(destination)
        marker.write_text(version,encoding="utf-8")
        return destination
    except Exception as exc:
        raise RuntimeError(f"Tidak dapat mengambil database dari Google Drive: {exc}") from exc

@st.cache_data(show_spinner="Membaca file…")
def read_uploaded_file(filename,payload,sheet_name=0):
    return pd.read_csv(io.BytesIO(payload)) if filename.lower().endswith(".csv") else pd.read_excel(io.BytesIO(payload),sheet_name=sheet_name)

def resolve_branch_sheet(sheet_names,branch_code):
    """Resolve a workbook tab for a branch, allowing city names and legacy BDG Sheet1."""
    normalized={key(name):name for name in sheet_names}
    candidates=[branch_code,BRANCH_OPTIONS[branch_code]]
    for candidate in candidates:
        if key(candidate) in normalized:
            return normalized[key(candidate)]
    if branch_code=="BDG" and "sheet1" in normalized:
        return normalized["sheet1"]
    if branch_code=="BDG" and len(sheet_names)==1:
        return sheet_names[0]
    return None

def uploaded_branch_sheet(uploaded_file,branch_code):
    if uploaded_file.name.lower().endswith(".csv"):
        return None
    try:
        _sheet_names=pd.ExcelFile(io.BytesIO(uploaded_file.getvalue())).sheet_names
        return resolve_branch_sheet(_sheet_names,branch_code)
    except Exception:
        return None

@st.cache_resource(show_spinner="Menyiapkan database transaksi…")
def read_normalized_dashboard_file(path,modified_ns,file_size,sheet_name=0):
    # Read large Excel sheets in batches. Loading the whole workbook with
    # pandas first can exceed the memory limit of the online Streamlit app.
    source_path=Path(path)
    _sheet_cache=key(sheet_name).replace(" ","_") if isinstance(sheet_name,str) else f"sheet_{sheet_name}"
    parquet_path=source_path.with_name(f".{source_path.stem}_{_sheet_cache}_dashboard_cache.parquet")
    metadata_path=parquet_path.with_suffix(".meta.json")
    expected={"modified_ns":int(modified_ns),"file_size":int(file_size),"sheet_name":sheet_name}
    if parquet_path.exists() and metadata_path.exists():
        try:
            if json.loads(metadata_path.read_text(encoding="utf-8"))==expected:
                return pd.read_parquet(parquet_path)
        except Exception:
            pass
    if source_path.suffix.lower()!=".xlsx":
        normalized=normalize(pd.read_excel(source_path,sheet_name=sheet_name),copy=False)
    else:
        from openpyxl import load_workbook
        import pyarrow as pa
        import pyarrow.parquet as pq
        workbook=load_workbook(source_path,read_only=True,data_only=True)
        writer=None
        temporary_parquet=parquet_path.with_name(parquet_path.name+".tmp")
        try:
            worksheet=workbook[sheet_name] if isinstance(sheet_name,str) else workbook.worksheets[int(sheet_name)]
            rows=worksheet.iter_rows(values_only=True)
            raw_headers=next(rows,None)
            if raw_headers is None:
                return pd.DataFrame()
            headers=[]; counts={}
            for index,value in enumerate(raw_headers):
                base=str(value).strip() if value is not None else f"Unnamed: {index}"
                count=counts.get(base,0); counts[base]=count+1
                headers.append(base if count==0 else f"{base}.{count}")
            sales_column=next((headers[i] for i in range(len(headers)) if key(headers[i]) in {key(alias) for alias in ALIASES["sales"]}),None)

            def write_batch(batch):
                nonlocal writer
                if not batch:
                    return
                frame=normalize(pd.DataFrame.from_records(batch,columns=headers),copy=False)
                actual_sales_column=col(frame,"sales")
                for column in frame.columns:
                    if column==actual_sales_column:
                        frame[column]=pd.to_numeric(frame[column],errors="coerce").astype("float64")
                    else:
                        frame[column]=frame[column].astype("string")
                table=pa.Table.from_pandas(frame,preserve_index=False)
                if writer is None:
                    writer=pq.ParquetWriter(temporary_parquet,table.schema)
                writer.write_table(table)

            batch=[]
            for row in rows:
                if not row or all(value is None for value in row):
                    continue
                batch.append(row)
                if len(batch)>=50000:
                    write_batch(batch)
                    batch=[]
            write_batch(batch)
            if writer is not None:
                writer.close(); writer=None
            else:
                empty=pd.DataFrame(columns=headers)
                for column in empty.columns:
                    empty[column]=empty[column].astype("float64" if column==sales_column else "string")
                empty.to_parquet(temporary_parquet,index=False)
            temporary_parquet.replace(parquet_path)
        finally:
            if writer is not None:
                writer.close()
            workbook.close()
            temporary_parquet.unlink(missing_ok=True)
        normalized=pd.read_parquet(parquet_path)
    try:
        temporary_metadata=metadata_path.with_name(metadata_path.name+".tmp")
        temporary_metadata.write_text(json.dumps(expected),encoding="utf-8")
        temporary_metadata.replace(metadata_path)
    except Exception:
        try: temporary_metadata.unlink(missing_ok=True)
        except OSError: pass
    return normalized

@st.cache_resource(show_spinner="Menyiapkan database transaksi…")
def read_normalized_dashboard_upload(filename,payload,sheet_name=0):
    raw=read_uploaded_file(filename,payload,sheet_name)
    return normalize(raw,copy=False)

def read_file(file):
    if file is None:
        file=Path(__file__).parent/"database_push.xlsx"
    if isinstance(file,(str,Path)):
        p=Path(file)
        if not p.exists(): return None
        stat=p.stat()
        try: return read_local_file(str(p),stat.st_mtime_ns,stat.st_size)
        except Exception as e: st.error(f"File tidak dapat dibaca: {e}"); return None
    if file is None:return None
    try: return read_uploaded_file(file.name,file.getvalue())
    except Exception as e: st.error(f"File tidak dapat dibaca: {e}"); return None

def calendar_weeks_in_period(path, year, month):
    """Count distinct company calendar weeks for a configured year and month."""
    calendar=read_file(path) if path and Path(path).exists() else None
    if calendar is None or calendar.empty:
        return 0
    year_col=next((c for c in calendar.columns if key(c) in {"yearid","year"}),None)
    month_col=next((c for c in calendar.columns if key(c) in {"monthid","month"}),None)
    week_col=next((c for c in calendar.columns if key(c)=="week"),None)
    if not all([year_col,month_col,week_col]):
        return 0
    scoped=calendar[
        matches_dimension(calendar[year_col],"year",[year]) &
        matches_dimension(calendar[month_col],"month",[month])
    ].copy()
    if scoped.empty:
        return 0
    scoped["_year"]=pd.to_numeric(scoped[year_col],errors="coerce")
    scoped["_week"]=pd.to_numeric(scoped[week_col],errors="coerce")
    return int(scoped.dropna(subset=["_year","_week"])[["_year","_week"]].drop_duplicates().shape[0])

def calendar_week_date_range(path, year, month, week):
    """Return first and last calendar dates for a configured company week."""
    calendar=read_file(path) if path and Path(path).exists() else None
    if calendar is None or calendar.empty:
        return None,None
    date_col=next((c for c in calendar.columns if key(c)=="date"),None)
    week_col=next((c for c in calendar.columns if key(c)=="week"),None)
    year_col=next((c for c in calendar.columns if key(c) in {"yearid","year"}),None)
    month_col=next((c for c in calendar.columns if key(c) in {"monthid","month"}),None)
    if not all([date_col,week_col,year_col,month_col]):
        return None,None
    scoped=calendar[matches_dimension(calendar[month_col],"month",[month])].copy()
    if year not in (None,"All",""):
        scoped=scoped[matches_dimension(scoped[year_col],"year",[year])]
    scoped["_date"]=pd.to_datetime(scoped[date_col],errors="coerce").dt.normalize()
    scoped["_week"]=pd.to_numeric(scoped[week_col],errors="coerce")
    scoped=scoped.dropna(subset=["_date","_week"])
    scoped=scoped[scoped["_week"].eq(int(week))]
    if scoped.empty:
        return None,None
    return scoped["_date"].min().date(),scoped["_date"].max().date()

def load_program_targets():
    """Load legacy targets and overlay files explicitly assigned to each program."""
    base_path=next(iter(SAVED_TARGET_DIR.glob("saved_target.*")),None)
    base=read_file(base_path)
    pieces=[]
    target_dir=PROGRAM_TARGET_DIR
    uploaded_programs=set()
    for program in PROGRAMS:
        path=target_dir/f"{key(program).replace(' ','_')}.csv"
        if path.exists():
            own=read_file(path)
            if own is not None and not own.empty:
                own=own.copy(); own["Program"]=program
                pieces.append(own); uploaded_programs.add(program)
    if base is not None and not base.empty:
        program_col=col(base,"program")
        for program in PROGRAMS:
            if program in uploaded_programs:
                continue
            fallback=base.copy()
            if program_col:
                vals=fallback[program_col].astype("string").str.strip().str.casefold()
                allowed={program.casefold(),program_label(program).casefold(),"all",""}
                fallback=fallback[vals.isna() | vals.isin(allowed)]
            if not fallback.empty:
                fallback["Program"]=program
                pieces.append(fallback)
    return pd.concat(pieces,ignore_index=True,sort=False) if pieces else base

if _data_source_mode=="Google Drive · spreadsheet bulanan":
    try:
        _account_fingerprint=hashlib.sha256(json.dumps(dict(_gcp_account),sort_keys=True).encode("utf-8")).hexdigest()
        _monthly_files=_google_drive_monthly_files(_drive_database_folder_id,ACTIVE_BRANCH_CODE,_account_fingerprint)
        _monthly_dir=APP_DIR/".remote_database"/"monthly"/ACTIVE_BRANCH_CODE
        _monthly_paths,_monthly_signature=_download_drive_monthly_files(_monthly_files,_monthly_dir,_gcp_account)
        if not _monthly_paths:
            st.error(f"Belum ada spreadsheet bulanan untuk cabang {ACTIVE_BRANCH_CODE} di Google Drive.")
            df=None
        else:
            df=read_normalized_monthly_branch(ACTIVE_BRANCH_CODE,tuple(_monthly_paths),_monthly_signature)
            if df.empty:
                st.warning(f"Folder {ACTIVE_BRANCH_CODE} belum berisi transaksi. Tambahkan data pada spreadsheet bulan berjalan di Drive.")
                st.stop()
    except Exception as _drive_error:
        st.error(f"Tidak dapat membaca spreadsheet bulanan dari Google Drive: {_drive_error}")
        df=None
    _default_database=None
elif _data_source_mode.startswith("Google Drive"):
    try:
        _default_database=_google_drive_database_path(
            _drive_source_id,
            _drive_database_path,
            _gcp_account,
            _drive_is_folder,
        )
    except Exception as _drive_error:
        st.error(str(_drive_error))
        _default_database=None
elif source is None:
    _default_database=APP_DIR/"database_push.xlsx"
elif source is not None:
    _default_database=None
if source is None and _data_source_mode!="Google Drive · spreadsheet bulanan":
    if _default_database is not None and _default_database.exists():
        _db_stat=_default_database.stat()
        try:
            _default_sheets=pd.ExcelFile(_default_database).sheet_names
            _branch_sheet=resolve_branch_sheet(_default_sheets,ACTIVE_BRANCH_CODE)
            if _branch_sheet is None:
                st.error(f"Sheet untuk cabang {ACTIVE_BRANCH_NAME} ({ACTIVE_BRANCH_CODE}) belum ada di {_default_database.name}. Tambahkan sheet dengan nama {ACTIVE_BRANCH_CODE}.")
                df=None
            else:
                df=read_normalized_dashboard_file(str(_default_database),_db_stat.st_mtime_ns,_db_stat.st_size,_branch_sheet)
        except Exception as e:
            st.error(f"File tidak dapat dibaca: {e}")
            df=None
    else:
        df=None
elif source is not None:
    try:
        _uploaded_sheet=uploaded_branch_sheet(source,ACTIVE_BRANCH_CODE)
        if _uploaded_sheet is None:
            _available_sheets=pd.ExcelFile(io.BytesIO(source.getvalue())).sheet_names
            st.error(f"Sheet cabang {ACTIVE_BRANCH_NAME} ({ACTIVE_BRANCH_CODE}) tidak ditemukan. Sheet yang tersedia: {', '.join(_available_sheets)}. Tambahkan sheet {ACTIVE_BRANCH_CODE} ke workbook.")
            df=None
        else:
            df=read_normalized_dashboard_upload(source.name,source.getvalue(),_uploaded_sheet)
    except Exception as e:
        st.error(f"File tidak dapat dibaca: {e}")
        df=None
target=load_program_targets()
if df is None:
    if _drive_configured:
        st.info("Database belum dapat dimuat dari Google Drive. Pastikan file spreadsheet bulan cabang aktif dapat diakses oleh service account dashboard, lalu klik 'Periksa pembaruan dari Drive'.")
    else:
        st.info("Unggah workbook Excel transaksi dari sidebar untuk melihat dashboard. Target dapat disiapkan terpisah melalui template sasaran.")
    st.stop()

def target_preview_for_config(program,params,metric):
    """Find the monthly target that should seed a program's weekly setting."""
    specific_path=PROGRAM_TARGET_DIR/f"{key(program).replace(' ','_')}.csv"
    target_source=read_file(specific_path) if specific_path.exists() else target
    if target_source is not None and not target_source.empty:
        target_source=target_source.copy()
        if specific_path.exists(): target_source["Program"]=program
        amount,rows=target_filtered(target_source,{**params,"program":program},metric)
        if rows is not None and not rows.empty:
            return float(amount or 0)
    if metric=="Total Sales":
        scope=df.copy()
        for kind in ("year","month","principal_name","principal","channel_group","business_group","product_code","product"):
            c=col(scope,kind); values=params.get(kind,[])
            if c and values:
                scope=scope[matches_dimension(scope[c],kind,values)]
        target_col=next((c for c in scope.columns if key(c) in ["target value","target sales","target value gross"]),None)
        dims=[col(scope,k) for k in ["year","month","salesperson","principal_name","principal","channel_group","business_group","product_code","product"]]
        dims=[c for c in dims if c]
        if target_col and dims:
            uniq=scope[dims+[target_col]].drop_duplicates()
            return float(pd.to_numeric(uniq[target_col],errors="coerce").fillna(0).sum())
    if metric!="Total Sales":
        return float(settings.get("program_targets",PROGRAM_TARGET_DEFAULTS).get(program,0))
    return 0.0

# ── Expander konfigurasi program per bulan (butuh df, disisipkan setelah df dimuat) ──
_metric_options = ["Total Relasi Beli Transaksi", "Total Sales"]
_metric_labels = {
    "Total Relasi Beli Transaksi": "RB (Relasi yang transaksi > 0)",
    "Total Sales": "SALES (total seluruh transaksi)",
}
if main_section=="Program Push" and "_active_push_month" in dir() and not st.session_state.get("nav_minimized", False):
    _tab_months_cfg = list(MONTHS.values())
    _pbm_cfg2 = settings.setdefault("program_config_by_month", {})
    _this_month_cfg2 = _pbm_cfg2.setdefault(_active_push_month, {})

    def _prog_cfg2(prog, kind, fallback=None):
        v = _this_month_cfg2.get(prog, {}).get(kind, None)
        if v is None and fallback is not None: return fallback
        if v is None: return PROGRAM_FILTER_DEFAULTS.get(prog, {}).get(kind, "")
        return v

    # Dropdown options dari df (sekarang sudah tersedia)
    _pid_c   = col(df,"principal");      _pid_all   = sorted(df[_pid_c].dropna().astype(str).unique()) if _pid_c else []
    _pc_c    = col(df,"product_code");   _pc_all    = sorted(df[_pc_c].dropna().astype(str).unique())  if _pc_c else []
    _prod_c  = col(df,"product");        _prod_all  = sorted(df[_prod_c].dropna().astype(str).unique()) if _prod_c else []
    _channel_c = col(df,"channel_group"); _channel_all = sorted(df[_channel_c].dropna().astype(str).unique()) if _channel_c else []
    _business_group_c = col(df,"business_group"); _business_group_all = sorted(df[_business_group_c].dropna().astype(str).unique()) if _business_group_c else []
    _year_c = col(df,"year"); _year_all = sorted(df[_year_c].dropna().astype(str).unique()) if _year_c else []
    def _sync_config_widget(widget_key, value):
        """Refresh persisted config changes without discarding unsaved edits."""
        source_key = f"{widget_key}__config_source"
        if st.session_state.get(source_key) != value:
            st.session_state[widget_key] = value
            st.session_state[source_key] = value

    with st.expander(f"⚙️ Konfigurasi Program · {_active_push_month}", expanded=False):
        st.caption("Pilih bulan transaksi, tahun, metric, dan filter masing-masing program untuk perhitungan Summary.")
        _template=target_template().to_csv(index=False).encode("utf-8-sig")
        st.download_button("Unduh template target CSV",_template,"template_target_chp_aam.csv","text/csv",key=f"target_template_{_active_push_month}")
        st.caption("Upload target di dalam setiap program. Kolom wajib: Year, Month, Sales Person Title, Target. Kolom filter opsional mengikuti pilihan konfigurasi program.")
        _cfg_new = {}
        for _p in PROGRAMS:
            _sn = str(_prog_cfg2(_p,"display_name", _p) or _p).strip() or _p
            st.markdown(f"**{PROGRAM_NUMBERS.get(_p,'PROGRAM')} — {_sn}**")
            _c1,_c2,_c3,_c4 = st.columns([2,1,1,2], gap="small")
            with _c1:
                _name_key=f"pbm2_name_{_active_push_month}_{key(_p).replace(' ','_')}"
                _sync_config_widget(_name_key,_sn)
                _nn = st.text_input("Nama",value=_sn,
                    key=_name_key,
                    label_visibility="collapsed", placeholder=f"Nama program {PROGRAM_NUMBERS.get(_p,'')}")
            with _c2:
                _pid_opts2 = ["All"]+_pid_all
                _sp = _prog_cfg2(_p,"principal", PROGRAM_FILTER_DEFAULTS.get(_p,{}).get("principal","All") or "All")
                if _sp not in _pid_opts2: _sp="All"
                _pid_key=f"pbm2_pid_{_active_push_month}_{key(_p).replace(' ','_')}"
                _sync_config_widget(_pid_key,_sp)
                _np = st.selectbox("Principal ID",_pid_opts2,
                    index=_pid_opts2.index(_sp),
                    key=_pid_key)
            with _c3:
                _pc_opts2 = ["All"]+_pc_all
                _spc = _prog_cfg2(_p,"product_code", PROGRAM_FILTER_DEFAULTS.get(_p,{}).get("product_code","All") or "All")
                if _spc not in _pc_opts2: _spc="All"
                _pc_key=f"pbm2_pc_{_active_push_month}_{key(_p).replace(' ','_')}"
                _sync_config_widget(_pc_key,_spc)
                _npc = st.selectbox("Product Code",_pc_opts2,
                    index=_pc_opts2.index(_spc),
                    key=_pc_key)
            with _c4:
                _prod_opts2 = ["All"]+_prod_all
                _sprod = _prog_cfg2(_p,"product", PROGRAM_FILTER_DEFAULTS.get(_p,{}).get("product","All") or "All")
                if _sprod not in _prod_opts2: _sprod="All"
                _prod_key=f"pbm2_prod_{_active_push_month}_{key(_p).replace(' ','_')}"
                _sync_config_widget(_prod_key,_sprod)
                _nprod = st.selectbox("Produk",_prod_opts2,
                    index=_prod_opts2.index(_sprod),
                    key=_prod_key)
            _c6,_c7,_c8,_c11 = st.columns([2,2,2,1], gap="small")
            with _c6:
                _channel_opts2 = ["All"]+_channel_all
                _schannel = _prog_cfg2(_p,"channel_group", PROGRAM_FILTER_DEFAULTS.get(_p,{}).get("channel_group","All") or "All")
                if _schannel not in _channel_opts2: _schannel="All"
                _channel_key=f"pbm2_channel_{_active_push_month}_{key(_p).replace(' ','_')}"
                _sync_config_widget(_channel_key,_schannel)
                _nchannel = st.selectbox("Channel Group",_channel_opts2,
                    index=_channel_opts2.index(_schannel),
                    key=_channel_key)
            with _c7:
                _business_group_opts2 = ["All"]+_business_group_all
                _sbusiness_group = _prog_cfg2(_p,"business_group", PROGRAM_FILTER_DEFAULTS.get(_p,{}).get("business_group","All") or "All")
                if _sbusiness_group not in _business_group_opts2: _sbusiness_group="All"
                _business_group_key=f"pbm2_business_group_{_active_push_month}_{key(_p).replace(' ','_')}"
                _sync_config_widget(_business_group_key,_sbusiness_group)
                _nbusiness_group = st.selectbox("Bisnis Group",_business_group_opts2,
                    index=_business_group_opts2.index(_sbusiness_group),key=_business_group_key)
            with _c8:
                _smetric = _prog_cfg2(_p,"metric", PROGRAM_FILTER_DEFAULTS.get(_p,{}).get("metric","Total Relasi Beli Transaksi"))
                if _smetric not in _metric_options: _smetric = "Total Relasi Beli Transaksi"
                _metric_widget_key=f"pbm2_metric_{_active_push_month}_{key(_p).replace(' ','_')}"
                _sync_config_widget(_metric_widget_key,_smetric)
                _nmetric = st.selectbox("Metric",_metric_options,
                    index=_metric_options.index(_smetric),
                    format_func=lambda value: _metric_labels.get(value,value),
                    key=_metric_widget_key)
            with _c11:
                _year_opts2 = ["All"]+_year_all
                _syear = str(_prog_cfg2(_p,"year", "All") or "All")
                if _syear not in _year_opts2: _syear="All"
                _year_key=f"pbm2_year_{_active_push_month}_{key(_p).replace(' ','_')}"
                _sync_config_widget(_year_key,_syear)
                _nyear = st.selectbox("Year",_year_opts2,
                    index=_year_opts2.index(_syear),
                    key=_year_key)
            _c9,_c10 = st.columns([1,1], gap="small")
            with _c9:
                _month_opts2 = _tab_months_cfg
                _scalc_month = _prog_cfg2(_p,"calculation_month",_active_push_month) or _active_push_month
                if _scalc_month not in _month_opts2: _scalc_month = _active_push_month
                _calc_month_key=f"pbm2_calc_month_{_active_push_month}_{key(_p).replace(' ','_')}"
                _sync_config_widget(_calc_month_key,_scalc_month)
                _ncalc_month = st.selectbox("Bulan Transaksi",_month_opts2,
                    index=_month_opts2.index(_scalc_month),key=_calc_month_key)
            _weekly_target_params={"year":[] if _nyear=="All" else [_nyear],
                                   "month":[_ncalc_month],
                                   "principal":[] if _np=="All" else [_np],
                                   "channel_group":[] if _nchannel=="All" else [_nchannel],
                                   "business_group":[] if _nbusiness_group=="All" else [_nbusiness_group],
                                   "product_code":[] if _npc=="All" else [_npc],
                                   "product":[] if _nprod=="All" else [_nprod],
                                   "program":_p}
            _monthly_config_target=target_preview_for_config(_p,_weekly_target_params,_nmetric)
            _allocation_options=["Manual / upload per Sales Person Title","Average 2026 (YTD)","Average 3 bulan terakhir"]
            _allocation_key=f"pbm2_target_basis_{_active_push_month}_{key(_p).replace(' ','_')}"
            _saved_allocation=_prog_cfg2(_p,"target_allocation_basis",_allocation_options[0])
            if _saved_allocation not in _allocation_options: _saved_allocation=_allocation_options[0]
            _sync_config_widget(_allocation_key,_saved_allocation)
            _target_basis=st.selectbox("Basis target per Sales Person Title",_allocation_options,
                                       index=_allocation_options.index(_saved_allocation),key=_allocation_key)
            _saved_auto_total=_prog_cfg2(_p,"target_total",_monthly_config_target or 0)
            try: _saved_auto_total=float(_saved_auto_total or 0)
            except (TypeError,ValueError): _saved_auto_total=float(_monthly_config_target or 0)
            _auto_total_key=f"pbm2_target_total_{_active_push_month}_{key(_p).replace(' ','_')}"
            _sync_config_widget(_auto_total_key,_saved_auto_total)
            _target_total=st.number_input("Target total program",min_value=0.0,value=max(0.0,_saved_auto_total),
                                          step=100000.0 if _nmetric=="Total Sales" else 1.0,format="%.0f",
                                          key=_auto_total_key,disabled=_target_basis==_allocation_options[0],
                                          help="Saat memilih Average, target total ini dialokasikan otomatis ke setiap Sales Person Title sesuai porsi actual historis.")
            st.caption("Average 2026 dihitung dari rata-rata per bulan sebelum bulan target; Average 3 bulan terakhir memakai tiga bulan penuh sebelumnya.")
            _config_week_count=calendar_weeks_in_period(calendar_path,_nyear,_ncalc_month) if _nyear!="All" else 0
            _weekly_week_key=f"pbm2_weekly_week_{_active_push_month}_{key(_p).replace(' ','_')}"
            _saved_weekly_week=int(_prog_cfg2(_p,"weekly_week",1) or 1)
            _sync_config_widget(_weekly_week_key,_saved_weekly_week)
            _nweekly_week=st.selectbox("Week actual acuan",[1,2,3,4],
                                       index=[1,2,3,4].index(_saved_weekly_week if _saved_weekly_week in [1,2,3,4] else 1),
                                       key=_weekly_week_key,
                                       format_func=lambda value:f"Week {value}",
                                       help="Actual mingguan dan analisa Summary dihitung dari tanggal yang ditandai pada minggu ini di Overview Calendar.")
            _default_week_start,_default_week_end=calendar_week_date_range(calendar_path,_nyear,_ncalc_month,_nweekly_week)
            _saved_start=_prog_cfg2(_p,"weekly_start_date",None)
            _saved_end=_prog_cfg2(_p,"weekly_end_date",None)
            _saved_week_for_dates=int(_prog_cfg2(_p,"weekly_week",1) or 1)
            if _saved_week_for_dates!=int(_nweekly_week):
                _saved_start=_saved_end=None
            try: _saved_start=pd.to_datetime(_saved_start).date() if _saved_start else None
            except Exception: _saved_start=None
            try: _saved_end=pd.to_datetime(_saved_end).date() if _saved_end else None
            except Exception: _saved_end=None
            _range_start=_saved_start or _default_week_start or pd.Timestamp.today().date()
            _range_end=_saved_end or _default_week_end or _range_start
            _date_range_key=f"pbm2_weekly_dates_{_active_push_month}_{key(_p).replace(' ','_')}_{int(_nweekly_week)}"
            _sync_config_widget(_date_range_key,(_range_start,_range_end) if _range_start and _range_end else None)
            _weekly_date_range=st.date_input("Rentang tanggal actual · dari — sampai",value=(_range_start,_range_end) if _range_start and _range_end else None,
                                             key=_date_range_key,
                                             help="Tanggal actual, plot progres, dan analisa mingguan dihitung hanya pada rentang ini.")
            if isinstance(_weekly_date_range,(tuple,list)) and len(_weekly_date_range)==2:
                _weekly_start_date,_weekly_end_date=_weekly_date_range
            else:
                _weekly_start_date=_weekly_end_date=None
            _default_weekly_percent=(100.0/_config_week_count if _config_week_count else 0.0)
            _saved_weekly_percent=_prog_cfg2(_p,"weekly_target_percent",None)
            if _saved_weekly_percent is None:
                _legacy_weekly_target=_prog_cfg2(_p,"weekly_target",None)
                _saved_weekly_percent=(float(_legacy_weekly_target or 0)/float(_monthly_config_target)*100
                                       if _legacy_weekly_target is not None and float(_monthly_config_target or 0)>0
                                       else _default_weekly_percent)
            _weekly_target_key=f"pbm2_weekly_target_pct_{_active_push_month}_{key(_p).replace(' ','_')}"
            _sync_config_widget(_weekly_target_key,float(_saved_weekly_percent or 0))
            _nweekly_target_percent=st.number_input(
                "Target Mingguan (% dari target bulanan)",
                min_value=0.0,max_value=100.0,value=min(100.0,float(_saved_weekly_percent or 0)),step=1.0,format="%.1f",
                key=_weekly_target_key,
                help="Atur porsi target bulanan untuk minggu berjalan. Summary menampilkan progres mingguan dalam persentase."
            )
            _nweekly_target=float(_monthly_config_target or 0)*float(_nweekly_target_percent or 0)/100.0
            with st.expander(f"🎯 Target khusus · {program_label(_p)}",expanded=False):
                _program_target_path=PROGRAM_TARGET_DIR/f"{key(_p).replace(' ','_')}.csv"
                _program_target_upload=st.file_uploader("Upload target untuk program ini (.xlsx / .csv)",type=["xlsx","xls","csv"],key=f"program_target_upload_{key(_p).replace(' ','_')}")
                _target_col1,_target_col2=st.columns([1,1])
                with _target_col1:
                    if st.button("Simpan target program",key=f"save_program_target_{key(_p).replace(' ','_')}",width="stretch"):
                        if _program_target_upload is None:
                            st.error("Pilih file target terlebih dahulu.")
                        else:
                            try:
                                _target_upload_df=read_uploaded_file(_program_target_upload.name,_program_target_upload.getvalue())
                                _required_target_kinds={"year":"Year","month":"Month","salesperson":"Sales Person Title"}
                                _target_rename={}
                                _missing_target=[]
                                for _kind,_label in _required_target_kinds.items():
                                    _source_col=col(_target_upload_df,_kind)
                                    if _source_col: _target_rename[_source_col]=_label
                                    else: _missing_target.append(_label)
                                _amount_col=next((c for c in _target_upload_df.columns if key(c) in ["target","target value","target sales","target relasi","target transaksi"]),None)
                                if _amount_col: _target_rename[_amount_col]="Target"
                                else: _missing_target.append("Target")
                                if _missing_target:
                                    st.error("Kolom wajib belum ditemukan: "+", ".join(_missing_target))
                                else:
                                    _target_upload_df=_target_upload_df.rename(columns=_target_rename)
                                    _target_upload_df["Program"]=_p
                                    PROGRAM_TARGET_DIR.mkdir(parents=True,exist_ok=True)
                                    _target_csv=_target_upload_df.to_csv(index=False).encode("utf-8-sig")
                                    _program_target_path.write_bytes(_target_csv)
                                    if st.secrets.get("gcp_service_account"):
                                        _drive_upload_bytes(_program_target_path.name,_target_csv,"text/csv","program_targets",ACTIVE_BRANCH_CODE)
                                    st.success(f"Target {_p} berhasil disimpan dan hanya digunakan untuk program ini.")
                                    st.rerun()
                            except Exception as _target_error:
                                st.error(f"File target tidak dapat diproses: {_target_error}")
                with _target_col2:
                    if st.button("Hapus target khusus",key=f"remove_program_target_{key(_p).replace(' ','_')}",width="stretch",disabled=not _program_target_path.exists()):
                        _program_target_path.unlink(missing_ok=True)
                        if st.secrets.get("gcp_service_account"):
                            try: _drive_delete_file(_program_target_path.name,"program_targets",ACTIVE_BRANCH_CODE)
                            except Exception as exc: st.error(f"Target di Drive gagal dihapus: {exc}")
                        st.success(f"Target khusus {_p} dihapus; target dasar kembali digunakan.")
                        st.rerun()
                st.caption(f"Target khusus tersimpan: {_program_target_path.name}" if _program_target_path.exists() else "Belum ada target khusus. Akan digunakan target dasar jika tersedia.")
            _cfg_new[_p] = {"display_name":_nn.strip() or _p,
                            "principal":_np,"product_code":_npc,"product":_nprod,
                            "channel_group":_nchannel,"business_group":_nbusiness_group,
                            "metric":_nmetric,"year":_nyear,
                            "calculation_month":_ncalc_month,
                            "weekly_target":float(_nweekly_target),
                            "weekly_target_percent":float(_nweekly_target_percent),
                            "weekly_week":int(_nweekly_week),
                            "weekly_start_date":_weekly_start_date.isoformat() if _weekly_start_date else None,
                            "weekly_end_date":_weekly_end_date.isoformat() if _weekly_end_date else None,
                            "target_allocation_basis":_target_basis,
                            "target_total":float(_target_total)}
            # Keep the Summary in sync with the current week/date widgets on the
            # same rerun; the Save button persists these settings for later.
            _cfg_save.setdefault(_p,{}).update({
                "weekly_target":float(_nweekly_target),
                "weekly_target_percent":float(_nweekly_target_percent),
                "weekly_week":int(_nweekly_week),
                "weekly_start_date":_weekly_start_date.isoformat() if _weekly_start_date else None,
                "weekly_end_date":_weekly_end_date.isoformat() if _weekly_end_date else None,
                "target_allocation_basis":_target_basis,
                "target_total":float(_target_total),
            })

        if st.button(f"💾 Simpan konfigurasi {_active_push_month}",
                     key=f"save_pbm2_{_active_push_month}", type="primary"):
            _pbm_cfg2[_active_push_month] = _cfg_new
            settings["program_config_by_month"] = _pbm_cfg2
            # Update session_state filter langsung; names remain scoped to this month's config.
            for _prog in PROGRAMS:
                _pkey2 = key(_prog).replace(" ","_")
                st.session_state[f"filter_month_{_pkey2}"]=[_cfg_new[_prog]["calculation_month"]]
                _pv=_cfg_new[_prog]["principal"];    st.session_state[f"filter_principal_{_pkey2}"]=[] if _pv=="All" else [_pv]
                _cv=_cfg_new[_prog]["product_code"]; st.session_state[f"filter_product_code_{_pkey2}"]=[] if _cv=="All" else [_cv]
                _rv=_cfg_new[_prog]["product"];      st.session_state[f"filter_product_{_pkey2}"]=[] if _rv=="All" else [_rv]
                _chv=_cfg_new[_prog]["channel_group"]; st.session_state[f"filter_channel_group_{_pkey2}"]=[] if _chv=="All" else [_chv]
                _bgv=_cfg_new[_prog]["business_group"]; st.session_state[f"filter_business_group_{_pkey2}"]=[] if _bgv=="All" else [_bgv]
                _yv=_cfg_new[_prog]["year"]; st.session_state[f"filter_year_{_pkey2}"]=[] if _yv=="All" else [str(_yv)]
                st.session_state[f"filter_metric_{_pkey2}"]=_cfg_new[_prog]["metric"]
            save_settings(settings)
            st.success(f"Konfigurasi {_active_push_month} tersimpan!")
            st.rerun()

for kind,label in [("year","Year"),("month","Month"),("principal","Principal ID"),("product","Product Description")]:
    if not col(df,kind): st.warning(f"Kolom {label} tidak ditemukan; filter ini tidak dapat diterapkan.")
if not col(df,"salesperson"): st.warning("Kolom Sales Person Title tidak ditemukan. Detail actual per title akan tersedia setelah kolom ini ditambahkan ke database.")

def get_program_filter(program, kind):
    pkey=key(program).replace(" ","_")
    widget_key=f"filter_{kind}_{pkey}"
    is_metric=kind=="metric"
    allowed=["Total Relasi Beli Transaksi","Total Sales"] if is_metric else filter_options(df,kind)
    if widget_key not in st.session_state:
        default=PROGRAM_FILTER_DEFAULTS.get(program,{}).get(kind,"Total Relasi Beli Transaksi" if is_metric else "All")
        saved=settings.get("filters",{}).get(program,{}).get(kind,default)
        st.session_state[widget_key]=saved if is_metric else ([] if saved in (None,"All","") else (saved if isinstance(saved,list) else [saved]))
    current=st.session_state[widget_key]
    if is_metric:
        if current not in allowed: st.session_state[widget_key]=allowed[0]
    else:
        if isinstance(current,str): current=[] if current in ("All","") else [current]
        st.session_state[widget_key]=[v for v in current if v in allowed]
    return st.session_state[widget_key]

def render_program_filters(program):
    rows=[]
    pkey=key(program).replace(" ","_")
    for row_kinds in [[("year","Year"),("month","Month"),("principal_name","Principal")],[("principal","Principal ID"),("channel_group","Channel Group"),("product_code","Product Code")],[("product","Product")]]:
        cols=st.columns(3)
        for box,(kind,label) in zip(cols,row_kinds):
            with box:
                get_program_filter(program,kind)
                rows.append((kind,filter_select(df,kind,label,f"filter_{kind}_{pkey}")))
    with st.columns([1,2,2])[0]:
        metric_key=f"filter_metric_{pkey}"
        get_program_filter(program,"metric")
        metric=st.selectbox("Metric",["Total Relasi Beli Transaksi","Total Sales"],key=metric_key)
    selected=dict(rows); selected["metric"]=metric
    _active_month=list(MONTHS.values())[min(max(int(st.session_state.get("push_month_tab_idx",pd.Timestamp.now().month-1)),0),11)]
    _program_config=settings.get("program_config_by_month",{}).get(_active_month,{}).get(program,{})
    selected["target_allocation_basis"]=_program_config.get("target_allocation_basis","Manual / upload per Sales Person Title")
    selected["target_total"]=_program_config.get("target_total",0)
    return selected

def apply_active_program_config(program, selected, config_month=None):
    """Make the active-month program configuration authoritative for Summary.

    The filter widgets keep their own Streamlit session state. If that state was
    created before switching month/configuration, it can be stale for one or
    more reruns. Derive the complete program scope directly from the active
    config so scorecards and exports always use the selected month and saved
    dimensions.
    """
    # Use the rendered config when Program Push is active. Other dashboard
    # sections (including Report Harian) resolve the persisted month config
    # directly because the Program Push-only _cfg_save variable is not created.
    active_month = config_month
    if active_month is None:
        _month_index = int(st.session_state.get("push_month_tab_idx", pd.Timestamp.now().month-1))
        _month_index = min(max(_month_index,0),11)
        active_month = list(MONTHS.values())[_month_index]
    rendered_cfg = globals().get("_cfg_save", {})
    cfg = rendered_cfg.get(program, {}) if isinstance(rendered_cfg,dict) and config_month is None else {}
    if not cfg:
        cfg = settings.get("program_config_by_month",{}).get(active_month,{}).get(program,{})
    if not cfg:
        cfg = PROGRAM_FILTER_DEFAULTS.get(program,{})
    selected = dict(selected)
    year_value = str(cfg.get("year", "All") or "All")
    selected["year"] = [] if year_value == "All" else [year_value]
    selected["principal_name"] = []
    calculation_month = cfg.get("calculation_month", active_month) or active_month
    selected["month"] = [calculation_month]
    for kind in ("principal", "product_code", "product"):
        value = cfg.get(kind, PROGRAM_FILTER_DEFAULTS.get(program, {}).get(kind, "All")) or "All"
        selected[kind] = [] if value == "All" else [value]
    for kind in ("channel_group", "business_group"):
        value = cfg.get(kind, PROGRAM_FILTER_DEFAULTS.get(program, {}).get(kind, "All")) or "All"
        selected[kind] = [] if value == "All" else [value]
    selected["metric"] = cfg.get("metric", PROGRAM_FILTER_DEFAULTS.get(program, {}).get("metric", "Total Relasi Beli Transaksi"))
    try:
        selected["weekly_week"] = int(cfg.get("weekly_week", 1) or 1)
    except (TypeError,ValueError):
        selected["weekly_week"] = 1
    selected["weekly_start_date"] = cfg.get("weekly_start_date")
    selected["weekly_end_date"] = cfg.get("weekly_end_date")
    selected["target_allocation_basis"] = cfg.get("target_allocation_basis","Manual / upload per Sales Person Title")
    selected["target_total"] = cfg.get("target_total",0)
    return selected

def filtered_data(selected):
    # Start from the cached canonical frame. Boolean filters below create the
    # scoped frames; copying the entire workbook here made each program scan
    # duplicate the full (often >100 MB) transaction database.
    scoped=df
    for kind in ["year","month","principal_name","principal","channel_group","business_group","product_code","product"]:
        c=col(scoped,kind); value=selected.get(kind,[])
        constrained = not (value is None or value == "All" or value == [] or value == "")
        # A missing configured dimension must never silently broaden a program
        # to the whole database. An empty scope is safer and exposes the schema
        # mismatch in the Summary result.
        if constrained and not c:
            return scoped.iloc[0:0].copy()
        if c and constrained:
            scoped=scoped[matches_dimension(scoped[c],kind,value)]
    return scoped

_AUTO_TARGET_ALLOCATION_CACHE={}
def auto_target_allocation(program,selected,basis,total_target):
    """Allocate a program total to titles by historical monthly actual shares."""
    try: total_target=float(total_target or 0)
    except (TypeError,ValueError): total_target=0.0
    if total_target<=0:
        return {}
    dimensions=[]
    for kind in ("principal_name","principal","channel_group","business_group","product_code","product"):
        values=selected.get(kind,[])
        if isinstance(values,(list,tuple,set)): values=tuple(sorted(map(str,values)))
        else: values=(str(values),)
        dimensions.append((kind,values))
    cache_key=(program,basis,selected.get("metric"),total_target,tuple(dimensions),
               str((selected.get("month") or [""])[0]),str((selected.get("year") or [""])[0]))
    if cache_key in _AUTO_TARGET_ALLOCATION_CACHE:
        return _AUTO_TARGET_ALLOCATION_CACHE[cache_key]
    date_col=col(df,"date"); title_col=col(df,"salesperson"); sales_col=col(df,"sales")
    if not date_col or not title_col or not sales_col:
        _AUTO_TARGET_ALLOCATION_CACHE[cache_key]={}
        return {}
    historic_filter=dict(selected)
    historic_filter["year"]=[]; historic_filter["month"]=[]
    historic=filtered_data(historic_filter).copy()
    historic["_auto_date"]=pd.to_datetime(historic[date_col],errors="coerce").dt.normalize()
    historic=historic.dropna(subset=["_auto_date"])
    target_year_values=selected.get("year",[])
    try: target_year=int(float(target_year_values[0])) if isinstance(target_year_values,(list,tuple)) and target_year_values else int(float(target_year_values))
    except (TypeError,ValueError): target_year=2026
    target_month_value=selected.get("month",[])
    target_month_value=target_month_value[0] if isinstance(target_month_value,(list,tuple)) and target_month_value else target_month_value
    target_month_number=next((n for n,name in MONTHS.items() if key(name)==key(target_month_value)),None)
    try:
        if target_month_number is None: target_month_number=int(float(str(target_month_value)))
    except (TypeError,ValueError): target_month_number=12
    if basis=="Average 3 bulan terakhir":
        target_month_start=pd.Timestamp(target_year,target_month_number,1)
        end_date=target_month_start-pd.Timedelta(days=1)
        start_date=end_date.to_period("M").start_time-pd.DateOffset(months=2)
        expected_months=3
    else:
        baseline_year=2026
        start_date=pd.Timestamp(baseline_year,1,1)
        if target_year==baseline_year and target_month_number>1:
            end_date=pd.Timestamp(baseline_year,target_month_number,1)-pd.Timedelta(days=1)
            expected_months=target_month_number-1
        else:
            end_date=pd.Timestamp(baseline_year,12,31)
            expected_months=12
    historic=historic[historic["_auto_date"].between(start_date,end_date)]
    if historic.empty:
        _AUTO_TARGET_ALLOCATION_CACHE[cache_key]={}
        return {}
    metric=selected.get("metric")
    if metric!="Total Sales":
        historic=positive_sales_rows(historic)
        if historic is None or historic.empty:
            _AUTO_TARGET_ALLOCATION_CACHE[cache_key]={}
            return {}
    historic["_auto_month"]=historic["_auto_date"].dt.to_period("M")
    historic["_auto_title"]=historic[title_col].astype("string").str.strip()
    if metric=="Total Sales":
        monthly=historic.groupby(["_auto_month","_auto_title"],dropna=False)[sales_col].sum()
    else:
        customer_col=col(historic,"customer")
        if customer_col:
            historic["_auto_customer"]=historic[customer_col].astype("string").str.strip().replace("",pd.NA)
            historic=historic.dropna(subset=["_auto_customer"])
            monthly=historic.groupby(["_auto_month","_auto_title"],dropna=False)["_auto_customer"].nunique()
        else:
            monthly=historic.groupby(["_auto_month","_auto_title"],dropna=False).size()
    if monthly.empty:
        _AUTO_TARGET_ALLOCATION_CACHE[cache_key]={}
        return {}
    average=monthly.groupby(level=1).sum()/max(1,expected_months)
    history_by_key={}
    for raw_title,value in average.items():
        title_key=key(raw_title)
        if title_key not in {"","nan","none","<na>"}:
            history_by_key[title_key]=history_by_key.get(title_key,0.0)+max(0.0,float(value))
    denominator=float(sum(history_by_key.values()))
    if denominator>0:
        allocation={title:total_target*value/denominator for title,value in history_by_key.items()}
    else:
        allocation={}
    _AUTO_TARGET_ALLOCATION_CACHE[cache_key]=allocation
    return allocation

def target_for(scoped,program,selected,title=None,require_explicit=()):
    params={**selected,"program":program}
    if title is not None: params["salesperson"]=str(title)
    allocation_basis=selected.get("target_allocation_basis","Manual / upload per Sales Person Title")
    if allocation_basis in {"Average 2026 (YTD)","Average 3 bulan terakhir"}:
        try: allocation_total=float(selected.get("target_total",0) or 0)
        except (TypeError,ValueError): allocation_total=0.0
        if title is None:
            return allocation_total,"Target total dialokasikan dari average historis"
        allocation=auto_target_allocation(program,selected,allocation_basis,allocation_total)
        return float(allocation.get(key(title),0.0)),"Alokasi otomatis berdasarkan average historis"
    specific_target_path=PROGRAM_TARGET_DIR/f"{key(program).replace(' ','_')}.csv"
    if specific_target_path.exists():
        specific_target=read_file(specific_target_path)
        if specific_target is not None and not specific_target.empty:
            # Read the selected program's own upload directly so legacy target
            # files or another program's rows cannot hide its matching targets.
            specific_target=specific_target.copy()
            specific_target["Program"]=program
            specific_total,specific_rows=target_filtered(specific_target,params,selected["metric"],require_explicit=require_explicit)
            if specific_rows is not None and not specific_rows.empty:
                return float(specific_total or 0),"File target khusus"
            return 0.0,"File target khusus tidak memiliki baris sesuai filter"
    tgt,matching=target_filtered(target,params,selected["metric"],require_explicit=require_explicit)
    source="file target"
    if matching is not None and not matching.empty:
        return float(tgt or 0),source
    # A program-specific upload is authoritative for that program, but targets
    # uploaded for a different program must not suppress the configured
    # baseline/database target for this one.
    if selected["metric"]=="Total Sales":
        target_col=next((c for c in scoped.columns if key(c) in ["target value","target sales","target value gross"]),None)
        dims=[col(scoped,k) for k in ["year","month","salesperson","principal_name","principal","channel_group","business_group","product_code","product"]]
        dims=[c for c in dims if c]
        if target_col and dims:
            uniq=scoped[dims+[target_col]].drop_duplicates()
            return float(pd.to_numeric(uniq[target_col],errors="coerce").fillna(0).sum()),"Target Value bawaan database"
    if selected["metric"]!="Total Sales" and title is None and not require_explicit:
        program_targets=settings.get("program_targets",PROGRAM_TARGET_DEFAULTS)
        return float(program_targets.get(program,0)),"Target program pada scorecard"
    return 0.0,"Belum ada target per Sales Person Title"

def calculate_program_kpi(program,selected):
    scoped=filtered_data(selected)
    actual=metric_actual(scoped,selected["metric"])
    tgt,source=target_for(scoped,program,selected)
    achievement=actual/tgt*100 if actual is not None and tgt else None
    deviation=actual-tgt if actual is not None else None
    return scoped,actual,tgt,achievement,deviation,source

def title_performance(scoped,program,selected):
    title_col=col(scoped,"salesperson")
    rows=[]
    seen=set()
    if title_col:
        for raw_title,grp in scoped.groupby(title_col,dropna=False):
            title=str(raw_title).strip()
            actual=metric_actual(grp,selected["metric"])
            tgt,_=target_for(grp,program,selected,title)
            rows.append({"Sales Person Title":title,"Actual":actual,"Target":tgt,
                         "A/T (%)":actual/tgt*100 if actual is not None and tgt else None,
                         "Deviation":actual-tgt if actual is not None else None})
            seen.add(title.casefold())
    # Add title-only targets so target holders remain visible even without a transaction row.
    if target is not None and not target.empty:
        params={**selected,"program":program}
        _,matching=target_filtered(target,params,selected["metric"])
        target_title_col=col(matching,"salesperson")
        if target_title_col:
            for raw_title in matching[target_title_col].dropna().astype(str).unique():
                title=raw_title.strip()
                if not title or title.casefold() in {"all","nan","none","<na>"} or title.casefold() in seen:
                    continue
                title_scoped=(scoped[scoped[title_col].astype(str).str.strip().str.casefold()==title.casefold()]
                              if title_col else scoped.iloc[0:0])
                actual=metric_actual(title_scoped,selected["metric"])
                if actual is None: actual=0
                tgt,_=target_for(title_scoped,program,selected,title)
                rows.append({"Sales Person Title":title,"Actual":actual,"Target":tgt,
                             "A/T (%)":actual/tgt*100 if actual is not None and tgt else None,
                             "Deviation":actual-tgt if actual is not None else None})
                seen.add(title.casefold())
    allocation_basis=selected.get("target_allocation_basis","Manual / upload per Sales Person Title")
    if title_col and allocation_basis in {"Average 2026 (YTD)","Average 3 bulan terakhir"}:
        allocated=auto_target_allocation(program,selected,allocation_basis,selected.get("target_total",0))
        historical_filter=dict(selected); historical_filter["year"]=[]; historical_filter["month"]=[]
        historical_titles=filtered_data(historical_filter)
        history_title_col=col(historical_titles,"salesperson")
        if history_title_col and allocated:
            for raw_title in historical_titles[history_title_col].dropna().astype(str).unique():
                title=raw_title.strip()
                if key(title) not in allocated or title.casefold() in seen:
                    continue
                title_scoped=scoped[scoped[title_col].astype(str).str.strip().str.casefold()==title.casefold()]
                actual=metric_actual(title_scoped,selected["metric"])
                if actual is None: actual=0
                tgt,_=target_for(title_scoped,program,selected,title)
                rows.append({"Sales Person Title":title,"Actual":actual,"Target":tgt,
                             "A/T (%)":actual/tgt*100 if actual is not None and tgt else None,
                             "Deviation":actual-tgt if actual is not None else None})
                seen.add(title.casefold())
    return pd.DataFrame(rows).sort_values("Actual",ascending=False) if rows else pd.DataFrame()

def append_daily_additions(scoped,by_title,metric):
    date_col=col(scoped,"date")
    title_col=col(scoped,"salesperson")
    sales_col=col(scoped,"sales")
    if not date_col or not title_col or not sales_col or by_title.empty:
        return by_title
    customer_col=col(scoped,"customer")
    wanted=[title_col,date_col]+([customer_col] if customer_col else [])+[sales_col]
    daily=scoped[list(dict.fromkeys(wanted))].copy()
    daily["_date"]=pd.to_datetime(daily[date_col],errors="coerce").dt.normalize()
    daily=daily.dropna(subset=["_date"])
    dates=sorted(daily["_date"].dropna().unique())
    daily["_title"]=daily[title_col].fillna("<NA>").astype(str)
    daily["_sales"]=pd.to_numeric(daily[sales_col],errors="coerce").fillna(0)
    eligible=daily.copy() if metric=="Total Sales" else daily[daily["_sales"]>0].copy()
    if metric=="Total Sales":
        increments=eligible.groupby(["_date","_title"])["_sales"].sum()
    elif customer_col:
        daily["_customer_key"]=daily[customer_col].astype("string").str.strip().replace("",pd.NA)
        eligible=daily[daily["_sales"]>0].copy()
        eligible=eligible.dropna(subset=["_customer_key"])
        first=eligible.groupby(["_title","_customer_key"],dropna=False)["_date"].min().reset_index()
        increments=first.groupby(["_date","_title"]).size()
    else:
        increments=eligible.groupby(["_date","_title"]).size()
    months={1:"JAN",2:"FEB",3:"MAR",4:"APR",5:"MAY",6:"JUN",7:"JUL",8:"AUG",9:"SEPT",10:"OCT",11:"NOV",12:"DEC"}
    include_year=len({pd.Timestamp(d).year for d in dates})>1
    for date_value in dates:
        date=pd.Timestamp(date_value).normalize()
        label=f"{date.day} {months[date.month]}"+(f" {date.year}" if include_year else "")
        by_title[label]=by_title["Sales Person Title"].map(lambda title: increments.get((date,str(title)),0)).fillna(0)
    return by_title

def weekly_program_performance(scoped,metric,selected,calendar_path):
    """Return title/program actuals for the current company week and weeks in scope."""
    calendar=read_file(calendar_path) if calendar_path and Path(calendar_path).exists() else None
    if calendar is None or calendar.empty:
        calendar=pd.DataFrame({"DATE":pd.date_range(pd.Timestamp(selected.get("weekly_start_date") or pd.Timestamp.today()).normalize(),pd.Timestamp(selected.get("weekly_end_date") or selected.get("weekly_start_date") or pd.Timestamp.today()).normalize()),"WEEK":int(selected.get("weekly_week") or 1),"YEARID":pd.Timestamp(selected.get("weekly_start_date") or pd.Timestamp.today()).year})
    date_key=next((c for c in calendar.columns if key(c)=="date"),None)
    week_col=next((c for c in calendar.columns if key(c)=="week"),None)
    year_col=next((c for c in calendar.columns if key(c) in {"yearid","year"}),None)
    month_col=next((c for c in calendar.columns if key(c) in {"monthid","month"}),None)
    data_date_col=col(scoped,"date")
    if not all([date_key,week_col,year_col,data_date_col]):
        return {},None,0,""
    cal=calendar[[date_key,week_col,year_col]+([month_col] if month_col else [])].copy()
    cal["_date"]=pd.to_datetime(cal[date_key],errors="coerce").dt.normalize()
    cal["_year"]=pd.to_numeric(cal[year_col],errors="coerce")
    cal["_week"]=pd.to_numeric(cal[week_col],errors="coerce")
    cal=cal.dropna(subset=["_date","_year","_week"]).drop_duplicates("_date")
    if cal.empty:
        return {},None,0,""
    period=cal.copy()
    year_filter=selected.get("year",[])
    if year_filter not in (None,"All",[],""):
        period=period[matches_dimension(period[year_col],"year",year_filter)]
    month_filter=selected.get("month",[])
    if month_col and month_filter not in (None,"All",[],""):
        period=period[matches_dimension(period[month_col],"month",month_filter)]
    week_count=len(period[["_year","_week"]].drop_duplicates())

    requested_start=selected.get("weekly_start_date")
    requested_end=selected.get("weekly_end_date")
    try: requested_start=pd.Timestamp(requested_start).normalize() if requested_start else None
    except Exception: requested_start=None
    try: requested_end=pd.Timestamp(requested_end).normalize() if requested_end else None
    except Exception: requested_end=None
    if requested_start is not None and requested_end is not None:
        if requested_end<requested_start:
            requested_start,requested_end=requested_end,requested_start
    requested_week=selected.get("weekly_week")
    if requested_week not in (None,"",0):
        week_number=int(requested_week)
    else:
        today=pd.Timestamp.now().normalize()
        today_row=cal[cal["_date"].eq(today)]
        if today_row.empty:
            latest=cal[cal["_date"].le(today)]
            today_row=latest.nlargest(1,"_date") if not latest.empty else cal.nsmallest(1,"_date")
        week_number=int(today_row.iloc[0]["_week"])
    if requested_start is None or requested_end is None:
        period=period[period["_week"].eq(week_number)]
        if period.empty:
            year_label=str((selected.get("year") or [""])[0])
            return {},0,week_count,f"Week {week_number}"+(f" · {year_label}" if year_label else "")
    else:
        # When an explicit date range is configured, it is the source of truth
        # for actual rows. The selected Week remains the label/target context;
        # intersecting it with the calendar week code can incorrectly zero the
        # actual when the configured range intentionally spans calendar labels.
        period=period[period["_date"].between(requested_start,requested_end)]
        if period.empty:
            week_year=int(selected.get("year",[pd.Timestamp.now().year])[0] or pd.Timestamp.now().year)
            return {},0,week_count,f"Week {week_number} · {week_year}"
    week_year=int(period.iloc[0]["_year"])
    week_key=(week_year,week_number)
    week_label=f"Week {week_number} · {week_year}"

    date_to_week=dict(zip(period["_date"],zip(period["_year"].astype(int),period["_week"].astype(int))))
    work=scoped.copy() if metric=="Total Sales" and col(scoped,"sales") else positive_sales_rows(scoped)
    if work is None:
        return {},None,week_count,week_label
    work=work.copy()
    work["_scope_date"]=pd.to_datetime(work[data_date_col],errors="coerce").dt.normalize()
    work["_week_key"]=work["_scope_date"].map(date_to_week)
    if requested_start is not None and requested_end is not None:
        weekly=work[work["_scope_date"].between(requested_start,requested_end)].copy()
    else:
        weekly=work[work["_week_key"].map(lambda value:isinstance(value,tuple) and value==week_key)].copy()
    title_col=col(weekly,"salesperson")
    customer_col=col(weekly,"customer")
    sales_col=col(weekly,"sales")
    by_title={}
    if title_col:
        for raw_title,group in weekly.groupby(title_col,dropna=False):
            title_key=key(raw_title)
            if metric=="Total Sales" and sales_col:
                value=float(pd.to_numeric(group[sales_col],errors="coerce").fillna(0).sum())
            elif customer_col:
                value=metric_actual(group,metric)
            else:
                value=int(len(group))
            by_title[title_key]=value
    weekly_total=metric_actual(weekly,metric)
    return by_title,weekly_total,week_count,week_label

def style_performance_table(table,compact=False,light=False,fit_width=False):
    midnight="#0C1D42"
    navy="#000737"
    blue="#22396F"
    cream="#FCF1CF"
    negative="#8F2638"
    if light:
        body_bg,body_text,base_bg="#F4F6F9","#26364A","#E4EAF3"
    else:
        body_bg,body_text,base_bg=navy,cream,midnight
    fixed_columns={"Salesperson_Title","ACT","TGT","A/T","DEV","TGT WEEKLY","DEV WEEKLY"}
    daily_columns=[column for column in table.columns if column not in fixed_columns]
    cell_padding="0px 1px" if compact else "7px 9px"
    head_padding="0px 2px" if compact else "9px 8px"
    spacing="0 0" if compact else "0 3px"
    table_props=[("border-collapse","separate"),("border-spacing",spacing),("border","1px solid "+blue),("background-color",body_bg),("font-size",".52rem" if compact else ".9rem"),("line-height",".72" if compact else "1.25"),("width","100%")]
    extra_styles=[]
    if fit_width:
        extra_styles=[
            {"selector":"","props":[("table-layout","fixed"),("width","100%"),("font-size",".82rem")]},
            {"selector":"th","props":[("white-space","normal"),("overflow-wrap","anywhere"),("padding","5px 3px")]},
            {"selector":"td","props":[("white-space","normal"),("overflow-wrap","anywhere"),("padding","5px 3px")]},
            {"selector":"th:first-child","props":[("width","24%"),("text-align","left")]},
            {"selector":"th:nth-child(2)","props":[("width","24%"),("text-align","left")]},
            {"selector":"td:first-child","props":[("width","24%"),("text-align","left")]},
        ]
    styled=table.style.set_table_attributes('class="compact-performance"' if compact else 'class="performance-table"').set_table_styles([
        {"selector":"table","props":table_props},
        {"selector":"th","props":[("background","linear-gradient(180deg, "+blue+" 0%, "+navy+" 100%)"),("color",cream),("font-weight","bold"),("border","1px solid "+blue),("text-align","center"),("box-shadow","inset 0 1px 0 rgba(252,241,207,.28), inset 0 -3px 0 rgba(0,7,55,.8)"),("padding",head_padding)]},
        {"selector":"td","props":[("border","1px solid "+blue),("background-color",body_bg),("color",body_text),("padding",cell_padding),("box-shadow","inset 0 1px 0 rgba(252,241,207,.08), inset 0 -2px 2px rgba(0,7,55,.65)")]},
        {"selector":"th.row_heading","props":[("display","none")]},
        {"selector":"th.blank","props":[("display","none")]},
    ]+extra_styles,overwrite=False)
    base_columns=[c for c in ["Salesperson_Title","ACT","TGT","TGT WEEKLY","DEV WEEKLY"] if c in table.columns]
    if base_columns:
        if light:
            styled=styled.set_properties(subset=base_columns,**{"background-color":base_bg,"color":body_text,"font-weight":"700"})
        else:
            styled=styled.set_properties(subset=base_columns,**{"background":"linear-gradient(180deg, #0C1D42 0%, #000737 100%)","color":cream,"box-shadow":"inset 0 1px 0 rgba(252,241,207,.10), inset 0 -2px 2px rgba(0,7,55,.75)"})
    if "A/T" in table.columns:
        def achievement_style(value):
            if pd.isna(value): return ""
            if value<70: return f"background-color:{negative};color:{cream};font-weight:bold"
            if value<100: return f"background-color:{blue};color:{cream};font-weight:bold"
            return "background-color:#D7F2E2;color:#146C3E;font-weight:bold"
        styled=styled.map(achievement_style,subset=["A/T"])
    if "DEV" in table.columns:
        def deviation_style(value):
            if pd.isna(value): return ""
            if value<0: return "background-color:#FCE1E4;color:#9E1B32;font-weight:bold"
            if value>0: return "background-color:#D7F2E2;color:#146C3E;font-weight:bold"
            return ""
        styled=styled.map(deviation_style,subset=["DEV"])
        if "DEV WEEKLY" in table.columns:
            styled=styled.map(deviation_style,subset=["DEV WEEKLY"])
    if daily_columns:
        daily_values=table[daily_columns].apply(pd.to_numeric,errors="coerce")
        maximum=float(daily_values.max().max()) if not daily_values.empty else 0.0
        def daily_style(value):
            if pd.isna(value): return ""
            ratio=max(0.0,min(1.0,float(value)/maximum)) if maximum>0 else 0.0
            if ratio<0.5:
                # Ease from midnight navy to Bandung blue.
                t=ratio*2
                start=(12,29,66); end=(34,57,111)
                rgb=tuple(round(a+(b-a)*t) for a,b in zip(start,end))
                color="#%02X%02X%02X"%rgb
                text=cream
            else:
                # Higher daily additions move toward the warm cream palette color.
                t=(ratio-.5)*2
                start=(34,57,111); end=(252,241,207)
                rgb=tuple(round(a+(b-a)*t) for a,b in zip(start,end))
                color="#%02X%02X%02X"%rgb
                text=navy if ratio>=.78 else cream
            weight="bold" if float(value)>0 else "normal"
            return f"background-color:{color};color:{text};font-weight:{weight}"
        styled=styled.map(daily_style,subset=daily_columns)
    if "Salesperson_Title" in table.columns:
        total_mask=table["Salesperson_Title"].astype(str).eq("TOTAL PROGRAM")
        if total_mask.any():
            if light:
                total_props={"font-weight":"900","background-color":"#D8E1EE","color":body_text,"border-top":"2px solid #22396F"}
            else:
                total_props={"font-weight":"900","border-top":"2px solid #FCF1CF"}
            styled=styled.set_properties(subset=pd.IndexSlice[total_mask,:],**total_props)
    return styled

def style_unified_table(data):
    n=len(data)
    row_colors=["background-color: #F4F6F9" if i%2==0 else "background-color: #E9EEF5" for i in range(n)]
    def _stripe(col_series):
        return [row_colors[i] for i in range(len(col_series))]
    return data.style.set_table_styles([
        {"selector":"th","props":[("background","linear-gradient(180deg,#33509B 0%,#22396F 55%,#182A52 100%)"),("color","#FCF1CF"),("font-weight","800"),("letter-spacing",".03em"),("border","1px solid #182A52"),("padding","6px 8px"),("text-align","center"),("text-shadow","0 1px 2px rgba(0,0,0,.45)"),("box-shadow","inset 0 1px 0 rgba(255,255,255,.22), inset 0 -2px 0 rgba(0,0,0,.30)")]},
        {"selector":"td","props":[("border","1px solid #DDE4EC"),("padding","5px 8px")]},
        {"selector":"table","props":[("border-collapse","collapse"),("background-color","#F4F6F9"),("width","100%")]},
    ]).set_properties(**{"background-color":"#F4F6F9","color":"#26364A"})\
      .apply(_stripe,axis=0)

def table_scroll_wrap(html,max_height=None):
    return f"<div class='table-scroll-wrap' style='width:100%;overflow:visible;max-height:none;border-radius:14px;box-shadow:0 10px 24px rgba(0,7,55,.35)'>{html}</div>"

def channel_group_performance(scoped,program,selected):
    channel_col=col(scoped,"channel_group")
    if not channel_col:
        return pd.DataFrame()
    rows=[]
    for group,grp in scoped.groupby(channel_col,dropna=False):
        label=str(group).strip() if pd.notna(group) and str(group).strip() else "(kosong)"
        actual=metric_actual(grp,selected["metric"])
        group_filter=selected.copy()
        group_filter["channel_group"]=[label] if label!="(kosong)" else []
        explicit=("channel_group",) if label!="(kosong)" else ()
        if label=="(kosong)" or (target is not None and col(target,"channel_group") is None):
            tgt=0.0
        else:
            tgt,_=target_for(grp,program,group_filter,require_explicit=explicit)
        rows.append({"Channel Group":label,"Actual":actual,"Target":tgt,
                     "A/T (%)":actual/tgt*100 if actual is not None and tgt else None,
                     "Deviation":actual-tgt if actual is not None and tgt else None})
    return pd.DataFrame(rows)

def latest_daily_progress(scoped,metric):
    eligible=scoped.copy() if metric=="Total Sales" and col(scoped,"sales") else positive_sales_rows(scoped)
    date_col=col(eligible,"date") if eligible is not None else None
    if eligible is None or not date_col or eligible.empty:
        return "Belum ada transaksi positif untuk membaca progres harian."
    eligible=eligible.copy()
    eligible["_date"]=pd.to_datetime(eligible[date_col],errors="coerce").dt.normalize()
    eligible=eligible.dropna(subset=["_date"])
    if eligible.empty:
        return "Kolom tanggal belum cukup untuk membaca progres harian."
    if metric=="Total Sales":
        sales_col=col(eligible,"sales")
        daily=eligible.groupby("_date")[sales_col].sum()
    else:
        customer_col=col(eligible,"customer")
        if customer_col:
            first=eligible.dropna(subset=[customer_col]).groupby(customer_col)["_date"].min()
            daily=first.value_counts().sort_index()
        else:
            daily=eligible.groupby("_date").size()
    latest=pd.Timestamp(eligible["_date"].max()).normalize()
    dates=pd.date_range(latest-pd.Timedelta(days=7),latest,freq="D")
    values=daily.reindex(dates,fill_value=0).astype(float)
    current=float(values.iloc[-1]); baseline=float(values.iloc[:-1].mean())
    latest_label=f"{latest.day} {MONTHS[latest.month][:3].upper()}"
    unit="sales" if metric=="Total Sales" else "relasi"
    current_text=fmt(current,metric)
    if baseline<=0:
        if current<=0:
            return f"Progres {unit} pada {latest_label} belum bergerak; belum ada basis pembanding tujuh hari sebelumnya."
        return f"Progres {latest_label} sebesar {current_text}; belum ada basis pembanding tujuh hari sebelumnya."
    change=(current/baseline-1)*100
    if current<baseline*0.7:
        return f"Progres harian melemah: {latest_label} {current_text}, {abs(change):.0f}% di bawah rata-rata tujuh hari sebelumnya ({fmt(baseline,metric)})."
    if current<baseline:
        return f"Progres {latest_label} {current_text}, masih {abs(change):.0f}% di bawah rata-rata tujuh hari sebelumnya ({fmt(baseline,metric)})."
    return f"Progres {latest_label} {current_text}, setara atau di atas rata-rata tujuh hari sebelumnya ({fmt(baseline,metric)})."

def summary_program_insight(scoped,program,selected,actual,target_value,achievement,deviation,title_detail):
    metric=selected["metric"]
    has_title_detail=(
        title_detail is not None
        and not title_detail.empty
        and {"Sales Person Title","Target","Actual","Deviation"}.issubset(title_detail.columns)
    )
    if has_title_detail:
        target_values=pd.to_numeric(title_detail["Target"],errors="coerce").fillna(0)
        deviations=pd.to_numeric(title_detail["Deviation"],errors="coerce")
        below_titles=title_detail[(target_values>0)&(deviations<0)].copy()
        below_titles["_gap"]=pd.to_numeric(below_titles["Target"],errors="coerce").fillna(0)-pd.to_numeric(below_titles["Actual"],errors="coerce").fillna(0)
        below_titles=below_titles.sort_values("_gap",ascending=False)
    else:
        # Keep summary rendering resilient when title detail is unavailable or
        # a target merge returns a frame without the expected title columns.
        below_titles=pd.DataFrame(columns=["Sales Person Title","Target","Actual","Deviation","_gap"])

    if actual is None:
        return [("Prioritas eksekusi: periksa kelengkapan transaksi dan pemetaan Sales Person Title; actual belum dapat dihitung.",True),
                ("Closing: validasi target tiap title, tetapkan owner peluang, nilai/relasi yang dikomit, dan tanggal close.",False)]

    if not below_titles.empty:
        # Name the largest shortfalls and give each owner a measurable next move.
        # Keep the summary compact so all five program cards remain shareable.
        priority=below_titles.head(3)
        actions=[]
        for _,row in priority.iterrows():
            title=str(row["Sales Person Title"])
            gap=fmt(row["_gap"],metric)
            row_actual=float(pd.to_numeric(pd.Series([row.get("Actual",0)]),errors="coerce").fillna(0).iloc[0])
            if metric=="Total Sales":
                action=(f"{title} (gap {gap}): susun daftar pelanggan repeat/upsell, ajukan nilai order senilai gap, "
                        "dan kunci tanggal PO.")
            elif row_actual<=0:
                action=(f"{title} (gap {gap}): prioritaskan outlet sasaran tanpa transaksi positif, lakukan kontak/kunjungan, "
                        "dan amankan order pertama.")
            else:
                action=(f"{title} (gap {gap}): follow-up outlet yang belum order ulang, minta komitmen jumlah outlet "
                        "dan tanggal transaksi.")
            actions.append(action)
        extra=max(0,len(below_titles)-len(priority))
        tail=f"; lanjutkan ke {extra} title berikutnya sesuai gap" if extra else ""
        improvement="Prioritas perbaikan: " + " ".join(actions) + tail
        parts=[(improvement,True)]
    elif not has_title_detail:
        parts=[("Detail Sales Person Title belum tersedia untuk filter ini; periksa kolom Sales Person Title dan target per title.",False)]
    else:
        parts=[("Prioritas perbaikan: semua Sales Person Title bertarget sudah mencapai target; jaga follow-up peluang yang siap order.",False)]

    if target_value is None or target_value<=0:
        closing="Closing: lengkapi target per Sales Person Title agar sisa gap dan komitmen closing dapat dihitung tepat."
        parts.append((closing,False))
    else:
        remaining=max(0.0,float(target_value)-(float(actual) if actual is not None else 0.0))
        unit="sales" if metric=="Total Sales" else "relasi"
        if remaining>0:
            focus=", ".join(str(title) for title in below_titles["Sales Person Title"].head(3)) if "Sales Person Title" in below_titles.columns else ""
            focus_text=f" Owner utama: {focus}." if focus else ""
            closing=(f"Closing: masih perlu {fmt(remaining,metric)} {unit}. Bagi gap ke owner sesuai peluang, "
                     f"catat komitmen dan tanggal close, cek order pending setiap hari, lalu pastikan realisasi tercatat.{focus_text}")
            parts.append((closing,True))
        else:
            closing="Closing: target tercapai. Konfirmasi order pending, pastikan realisasi tercatat, dan jaga repeat order sampai periode ditutup."
            parts.append((closing,False))
    return parts

def summary_weekly_insight(program,metric,week_label,weekly_actual,weekly_target,
                           weekly_by_title,title_detail):
    """Explain weekly progress against the configured weekly target."""
    unit="sales" if metric=="Total Sales" else "relasi"
    if weekly_actual is None:
        return [(f"Actual {week_label or 'mingguan'} belum dapat dihitung; periksa kolom tanggal dan kalender perusahaan.",True)]
    if weekly_target is None or float(weekly_target)<=0:
        return [(f"Target mingguan {program} belum disetel. Isi target pada Konfigurasi Program agar analisa minggu berjalan aktif.",True)]
    target=float(weekly_target)
    actual=float(weekly_actual)
    remaining=max(0.0,target-actual)
    title_gaps=[]
    title_on_track=[]
    if title_detail is not None and not title_detail.empty:
        detail_targets=pd.to_numeric(title_detail.get("Target"),errors="coerce").fillna(0)
        total_detail_target=float(detail_targets.sum())
        if total_detail_target>0:
            for (_,row),monthly_title_target in zip(title_detail.iterrows(),detail_targets):
                if monthly_title_target<=0: continue
                title=str(row.get("Sales Person Title",""))
                weekly_title_target=target*float(monthly_title_target)/total_detail_target
                title_actual=float(weekly_by_title.get(key(title),0) or 0)
                gap=max(0.0,weekly_title_target-title_actual)
                if gap>0:
                    title_gaps.append((gap,title,title_actual,weekly_title_target))
                else:
                    title_on_track.append((title_actual,title,weekly_title_target))
    title_gaps.sort(reverse=True)
    title_on_track.sort(reverse=True)
    parts=[]
    if remaining>0:
        parts.append((f"{week_label}: actual {fmt_summary(actual,metric)} dari target {fmt_summary(target,metric)}; gap {fmt_summary(remaining,metric)} {unit}.",True))
    else:
        parts.append((f"{week_label}: target mingguan tercapai ({fmt_summary(actual,metric)} / {fmt_summary(target,metric)}). Pertahankan ritme.",False))
    for gap,title,title_actual,title_target in title_gaps[:3]:
        if metric=="Total Sales":
            note="follow-up PO dan pastikan tanggal order"
        elif title_actual<=0:
            note="prioritaskan relasi sasaran yang belum transaksi"
        else:
            note="kejar relasi yang belum transaksi ulang"
        parts.append((f"Belum on track: {title} ({fmt_summary(title_actual,metric)} / {fmt_summary(title_target,metric)}; gap {fmt_summary(gap,metric)}) — {note}.",True))
    for title_actual,title,title_target in title_on_track[:3]:
        parts.append((f"On track: {title} ({fmt_summary(title_actual,metric)} / {fmt_summary(title_target,metric)}). Pertahankan pencapaian.",False))
    return parts

def insight_for(program,actual,target_value,achievement,deviation,by_title,metric):
    if actual is None:
        return "Kolom yang diperlukan untuk menghitung actual tidak tersedia pada database."
    if not target_value:
        return f"Actual {fmt(actual,metric)} tercatat, tetapi target belum tersedia untuk filter program ini. Unggah sasaran yang memuat Year, Month, dan Sales Person Title."
    lead=by_title.iloc[0]["Sales Person Title"] if not by_title.empty else "belum ada detail title"
    below=by_title[by_title["Deviation"].notna() & (by_title["Deviation"]<0)] if not by_title.empty else pd.DataFrame()
    lag=below.sort_values("Deviation").iloc[0]["Sales Person Title"] if not below.empty else None
    if deviation<0:
        gap=fmt(abs(deviation),metric)
        extra=f" Title dengan actual tertinggi: {lead}." if lead else ""
        if lag: extra+=f" Prioritaskan tindak lanjut pada {lag}, yang memiliki gap terbesar per title."
        return f"{program} mencapai {achievement:.1f}% target dengan kekurangan {gap}.{extra}"
    return f"{program} telah mencapai {achievement:.1f}% target dengan surplus {fmt(deviation,metric)}. Actual tertinggi berasal dari {lead}; pertahankan aktivitas yang menghasilkan capaian tersebut."

def canonical_targets(source_df):
    if source_df is None:
        tcol=next((c for c in df.columns if key(c) in ["target value","target sales","target value gross"]),None)
        if not tcol: return target_template()
        dims=[col(df,k) for k in ["year","month","salesperson","principal_name","principal","channel_group","business_group","product_code","product"]]
        dims=[c for c in dims if c]
        if not dims: return target_template()
        out=df[dims+[tcol]].drop_duplicates().copy()
        rename={}
        for kind,label in [("year","Year"),("month","Month"),("salesperson","Sales Person Title"),("principal_name","Principal"),("principal","Principal ID"),("channel_group","Channel Group"),("business_group","Business Group"),("product_code","Product Code"),("product","Product Description")]:
            c=col(out,kind)
            if c: rename[c]=label
        rename[tcol]="Target"
        out=out.rename(columns=rename)
        out["Program"]=""
    else:
        out=source_df.copy()
        rename={}
        for kind,label in [("year","Year"),("month","Month"),("salesperson","Sales Person Title"),("principal_name","Principal"),("principal","Principal ID"),("channel_group","Channel Group"),("business_group","Business Group"),("product_code","Product Code"),("product","Product Description"),("program","Program")]:
            c=col(out,kind)
            if c: rename[c]=label
        target_col=next((c for c in out.columns if key(c) in ["target","target value","target sales","target relasi","target transaksi"]),None)
        if target_col: rename[target_col]="Target"
        out=out.rename(columns=rename)
    for c in ["Year","Month","Sales Person Title","Target","Program","Principal","Principal ID","Channel Group","Business Group","Product Code","Product Description"]:
        if c not in out: out[c]=""
    out=out[["Year","Month","Sales Person Title","Target","Program","Principal","Principal ID","Channel Group","Business Group","Product Code","Product Description"]]
    return out.reset_index(drop=True)

if main_section=="Menu Insentif":
    st.markdown("<div class='score-heading'>◇ MENU INSENTIF · BULAN BERJALAN</div>",unsafe_allow_html=True)

    # ── Filter Tahun & Bulan — default bulan berjalan, align ke Performance Sales ──
    _now = pd.Timestamp.now()
    _perf_saved = settings.get("performance_filters",{})
    _def_yr = int(_perf_saved.get("year", _now.year) or _now.year)
    _def_mn = int(_perf_saved.get("month", _now.month) or _now.month)

    # ambil opsi tahun & bulan dari database
    _yr_c_ins = next((c for c in df.columns if c.lower()=="year"), None)
    _mn_c_ins = next((c for c in df.columns if c.lower()=="month"), None)
    _yr_opts = sorted(df[_yr_c_ins].dropna().astype(int).unique().tolist()) if _yr_c_ins else [_def_yr]
    _mn_opts = sorted(df[_mn_c_ins].dropna().astype(int).unique().tolist()) if _mn_c_ins else list(range(1,13))

    _ins_filt_cols = st.columns([1,1,3])
    with _ins_filt_cols[0]:
        _ins_yr = st.selectbox("Tahun Insentif",_yr_opts,
            index=_yr_opts.index(_def_yr) if _def_yr in _yr_opts else len(_yr_opts)-1,
            key="ins_year")
    with _ins_filt_cols[1]:
        from calendar import month_abbr as _mabb
        _ins_mn = st.selectbox("Bulan Insentif",_mn_opts,
            index=_mn_opts.index(_def_mn) if _def_mn in _mn_opts else len(_mn_opts)-1,
            format_func=lambda m: f"{m} – {_mabb[m]}",
            key="ins_month")
    with _ins_filt_cols[2]:
        st.caption(f"📅 Menghitung insentif bulan **{_mabb[_ins_mn]} {_ins_yr}** · Scope aligned dengan filter Performance Sales. Ubah bulan di sini untuk simulasi bulan lain.")

    insentif_role=st.radio("Pilih Role",["Supervisor","Salesman CHP"],horizontal=True,key="insentif_role")

    # ── helper: deteksi tipe CHP dari nama title ─────────────────────────
    def detect_sp_type(title):
        t=str(title).upper()
        if "NHS" in t:  return "NHS"
        if "MOD" in t:  return "MOD"
        return "ALL_REG"

    # ── helper: hitung semua parameter insentif dari database ───────────────
    def calc_insentif_prefill(sp_title, sel_year=None, sel_month=None, table_type="SUPV"):
        """
        Hitung parameter insentif dari database untuk sp_title.
        Returns dict {param_key: nilai_pct, ...} + "_notes" dict untuk keterangan.
        table_type: "SUPV" | "ALL_REG" | "MOD" | "NHS"
        """
        if df is None or not sp_title: return {}, {}
        sc = col(df,"salesperson"); sales_c = col(df,"sales"); tgt_c = col(df,"target_value")
        pid_c = col(df,"principal"); cust_c = col(df,"customer"); dt_c = col(df,"date")
        yr_c  = next((c for c in df.columns if c.lower()=="year"), None)
        mn_c  = next((c for c in df.columns if c.lower()=="month"), None)
        if not sc: return {}, {}

        # scope ke title yang dipilih
        base = df[df[sc].astype(str).str.strip()==str(sp_title).strip()].copy()
        if base.empty: return {}, {}

        # scope ke bulan berjalan jika dipilih
        if yr_c and mn_c and sel_year:
            base = base[pd.to_numeric(base[yr_c],errors="coerce")==int(sel_year)]
        bm = base.copy()  # base untuk bulan berjalan
        if mn_c and sel_month:
            bm = base[pd.to_numeric(base[mn_c],errors="coerce")==int(sel_month)]

        out = {}; notes = {}

        # ── PARAM 1: % A/T Gross Sales (Non Toserba = exclude CHW) ──────────
        if sales_c and tgt_c and pid_c:
            non_chw = bm[~bm[pid_c].astype(str).str.upper().str.strip().eq("CHW")].copy()
            non_chw[sales_c] = pd.to_numeric(non_chw[sales_c],errors="coerce").fillna(0)
            non_chw[tgt_c]   = pd.to_numeric(non_chw[tgt_c],  errors="coerce").fillna(0)
            act_nc = float(non_chw[non_chw[sales_c]>0][sales_c].sum())
            tgt_nc = float(non_chw[tgt_c].sum())
            at_nc  = round(act_nc/tgt_nc*100,1) if tgt_nc>0 else 0.0
            param1_key = {
                "SUPV":    "1. % A/T Gross Sales Cabang (Non Toserba)",
                "ALL_REG": "1. % A/T Gross Sales Rayon (Non Toserba)",
                "MOD":     "1. % A/T Gross Sales Rayon",
                "NHS":     "1. % A/T Gross Sales Rayon",
            }.get(table_type,"1. % A/T Gross Sales Rayon (Non Toserba)")
            out[param1_key] = at_nc
            notes[param1_key] = f"Actual Rp {act_nc:,.0f} / Target Rp {tgt_nc:,.0f} (exclude CHW)"

        # ── PARAM 2: % A/T Sales Toserba (CHW only) ─────────────────────────
        if sales_c and tgt_c and pid_c:
            chw = bm[bm[pid_c].astype(str).str.upper().str.strip().eq("CHW")].copy()
            chw[sales_c] = pd.to_numeric(chw[sales_c],errors="coerce").fillna(0)
            chw[tgt_c]   = pd.to_numeric(chw[tgt_c],  errors="coerce").fillna(0)
            act_chw = float(chw[chw[sales_c]>0][sales_c].sum())
            tgt_chw = float(chw[tgt_c].sum())
            at_chw  = round(act_chw/tgt_chw*100,1) if tgt_chw>0 else 0.0
            param2_key_map = {
                "SUPV":    "2. % A/T Sales Toserba",
                "ALL_REG": "4. % A/T Sales Toserba (CHW)",
            }
            if table_type in param2_key_map:
                pk2 = param2_key_map[table_type]
                out[pk2] = at_chw
                notes[pk2] = f"CHW — Actual Rp {act_chw:,.0f} / Target Rp {tgt_chw:,.0f}"

        # ── PARAM 3: POT (% Principal On Target > 100%, dari total 20 POT) ──
        if sales_c and tgt_c and pid_c:
            POT_TOTAL = 20
            pid_grp = bm.copy()
            pid_grp[sales_c] = pd.to_numeric(pid_grp[sales_c],errors="coerce").fillna(0)
            pid_grp[tgt_c]   = pd.to_numeric(pid_grp[tgt_c],  errors="coerce").fillna(0)
            pid_agg = (pid_grp[pid_grp[sales_c]>0]
                       .groupby(pid_c, dropna=False)
                       .agg(act_sum=(sales_c,"sum"),tgt_sum=(tgt_c,"sum"))
                       .reset_index())
            pot_on  = int((pid_agg["act_sum"]>=pid_agg["tgt_sum"]).sum())
            pot_pct = round(pot_on/POT_TOTAL*100,1)
            param3_key_map = {
                "SUPV":    "3. Principal On Target (POT) · Target A/T Wajib 100%",
                "ALL_REG": "3. Principal On Target (POT) · A/T Wajib 100%",
                "MOD":     "2. Principal On Target (POT) · A/T Wajib 100%",
            }
            if table_type in param3_key_map:
                pk3 = param3_key_map[table_type]
                out[pk3] = pot_pct
                notes[pk3] = f"{pot_on} dari {POT_TOTAL} Principal On Target (≥100%)"

        # ── PARAM 4: % EC (Effective Call) — relasi unik transaksi per hari ─
        # Hitung: jumlah hari-dengan-transaksi / jumlah hari-kerja-di-bulan
        # Proxy: jumlah tanggal unik dengan transaksi positif / Working Day di bulan
        if dt_c and cust_c and sales_c and mn_c and yr_c:
            bm2 = bm.copy()
            bm2[sales_c] = pd.to_numeric(bm2[sales_c],errors="coerce").fillna(0)
            bm2 = bm2[bm2[sales_c]>0]
            bm2["_dt"] = pd.to_datetime(bm2[dt_c],errors="coerce").dt.normalize()
            bm2 = bm2.dropna(subset=["_dt"])
            actual_days = bm2["_dt"].nunique()
            # ambil Working Day dari kolom database jika ada
            wd_c = next((c for c in df.columns if "working day" in c.lower() and "left" not in c.lower()), None)
            if wd_c and not bm.empty:
                wd_val = pd.to_numeric(bm[wd_c],errors="coerce").dropna()
                working_days = int(wd_val.max()) if not wd_val.empty else max(actual_days,1)
            else:
                # fallback: hari bisnis dari calendar
                if sel_month and sel_year:
                    import calendar as _cal
                    _, last = _cal.monthrange(int(sel_year), int(sel_month))
                    m_dates = pd.date_range(f"{int(sel_year)}-{int(sel_month):02d}-01", periods=last)
                    working_days = int(pd.Series(m_dates).map(lambda d: d.weekday()<5).sum())
                else:
                    working_days = max(actual_days,1)
            ec_pct = round(actual_days/working_days*100,1) if working_days>0 else 0.0
            param4_key_map = {
                "SUPV":    "4. Actual EC (Effective Call) Vs Target EC Per Salesman",
                "ALL_REG": "2. Actual Effective Call vs Target Call per hari (HK kalender FC 2026)",
                "MOD":     "4. Actual Effective Call vs Target Call per hari (HK kalender FC 2026)",
            }
            if table_type in param4_key_map:
                pk4 = param4_key_map[table_type]
                out[pk4] = ec_pct
                notes[pk4] = f"{actual_days} hari transaksi / {working_days} hari kerja"

        # ── PARAM 5: % Program PUSH (program ≥100% / total program) ─────────
        # Align dengan dashboard: hitung dari calculate_program_kpi per program
        try:
            # ── Align dengan Summary Program dashboard ──────────────────────
            # Hitung total actual / total target lintas semua PROGRAMS,
            # menggunakan filter yang sudah tersimpan di session_state
            # (read-only — tidak modifikasi state untuk menghindari widget error).
            def _rf(prog, kind):
                wk = f"filter_{kind}_{key(prog).replace(' ','_')}"
                return st.session_state.get(wk, PROGRAM_FILTER_DEFAULTS.get(prog,{}).get(kind,"All"))

            total_act = 0.0; total_tgt = 0.0
            prog_detail = []
            for prog in PROGRAMS:
                psel = {k: _rf(prog, k) for k in ["year","month","principal_name","principal",
                                                    "channel_group","product_code","product","metric"]}
                # scope ke bulan insentif jika filter belum diset per bulan
                if sel_year and psel.get("year") in (None,"All",""):
                    psel["year"] = str(sel_year)
                if sel_month and psel.get("month") in (None,"All",""):
                    psel["month"] = MONTHS.get(int(sel_month), str(sel_month))
                _, pact, ptgt, pach, _, _ = calculate_program_kpi(prog, psel)
                pact = pact or 0.0; ptgt = ptgt or 0.0
                total_act += pact; total_tgt += ptgt
                prog_detail.append(f"{program_label(prog)}: A/T {pach:.1f}%" if pach else f"{program_label(prog)}: —")

            push_pct = round(total_act / total_tgt * 100, 1) if total_tgt > 0 else 0.0
            param5_key_map = {
                "SUPV":    "5. % Program PUSH (Sesuai bulan terkait)",
                "ALL_REG": "6. % Program PUSH (Sesuai bulan terkait)",
                "MOD":     "3. % Program PUSH (Sesuai bulan terkait)",
                "NHS":     "4. % Program PUSH",
            }
            if table_type in param5_key_map:
                pk5 = param5_key_map[table_type]
                out[pk5] = push_pct
                notes[pk5] = (f"Total ACT: {total_act:,.0f} / Total TGT: {total_tgt:,.0f} "
                              f"({push_pct:.1f}%) — " + " | ".join(prog_detail))
        except Exception:
            pass

        # ── PARAM 6: % Growth RB CHP (bulan berjalan vs AVG RB 2025) ────────
        # RB = Customer Name unik dengan Sales Value Gross > 0, bisnis CHP
        if cust_c and sales_c and yr_c and mn_c and pid_c:
            chp_base = base.copy()
            chp_base[sales_c] = pd.to_numeric(chp_base[sales_c],errors="coerce").fillna(0)
            chp_base = chp_base[chp_base[sales_c]>0]

            # RB bulan berjalan 2026
            rb_cur = 0
            if sel_year and sel_month:
                cur_rows = chp_base[
                    (pd.to_numeric(chp_base[yr_c],errors="coerce")==int(sel_year)) &
                    (pd.to_numeric(chp_base[mn_c],errors="coerce")==int(sel_month))
                ]
                rb_cur = int(cur_rows[cust_c].dropna().astype(str).str.strip().nunique())

            # AVG RB 2025 per bulan
            rows_2025 = chp_base[pd.to_numeric(chp_base[yr_c],errors="coerce")==2025]
            if not rows_2025.empty and mn_c:
                rb_by_month_2025 = rows_2025.groupby(
                    pd.to_numeric(rows_2025[mn_c],errors="coerce")
                )[cust_c].nunique()
                avg_rb_2025 = float(rb_by_month_2025.mean()) if not rb_by_month_2025.empty else 0.0
            else:
                avg_rb_2025 = 0.0

            growth_rb = round((rb_cur/avg_rb_2025-1)*100,1) if avg_rb_2025>0 else 0.0
            param6_key_map = {
                "SUPV":    "6. % Growth RB CHP (dihitung dari AVG RB 2025)",
                "ALL_REG": "5. % Growth RB CHP (dihitung dari AVG RB 2025)",
                "NHS":     "3. % Growth RB (Actual RB Bulan berjalan 2026 vs AVG RB 2025)",
            }
            if table_type in param6_key_map:
                pk6 = param6_key_map[table_type]
                out[pk6] = max(0.0, growth_rb)
                notes[pk6] = f"RB bulan ini: {rb_cur} | AVG RB 2025: {avg_rb_2025:.1f} | Growth: {growth_rb:+.1f}%"

        return out, notes

    # ── helper: render kalkulator otomatis ───────────────────────────────
    def render_kalkulator(table_def, role_key, title_suffix="", prefill=None, notes=None):
        """
        Auto-render: nilai dari prefill langsung dipakai sebagai hasil kalkulasi.
        Parameter yang tidak ada di prefill masih bisa diisi manual.
        Hasil insentif dihitung realtime tanpa tombol.
        """
        prefill = prefill or {}; notes = notes or {}
        from calendar import month_abbr as _mabb2
        mn_label = _mabb2[int(_ins_mn)] if _ins_mn else "—"
        st.caption(f"Monitoring insentif **{mn_label} {_ins_yr}** · 🔵 = dihitung otomatis dari database · ⬜ = input manual (Collection dll)")
        nama = st.text_input("Nama Salesman",value="",placeholder="Contoh: Budi Santoso",key=f"ins_nama_{role_key}")

        # Bangun nilai input: auto dari prefill, sisanya manual
        input_vals = {}
        manual_params = [p for p in table_def if p not in prefill]
        if manual_params:
            st.markdown("**Parameter manual (tidak tersedia di database):**")
            cl2, cr2 = st.columns(2, gap="large")
            for i, param in enumerate(manual_params):
                with (cl2 if i%2==0 else cr2):
                    input_vals[param] = st.number_input(
                        param, min_value=0.0, max_value=999.0, value=0.0,
                        step=0.1, format="%.1f",
                        key=f"ins_manual_{role_key}_{i}",
                        help="Masukkan nilai aktual (%)")
        # Gabungkan: auto prefill + manual
        for param in table_def:
            if param not in input_vals:
                input_vals[param] = float(prefill.get(param, 0.0))

        st.markdown("---")
        st.markdown("### 📊 Hasil Kalkulasi Insentif Otomatis")

        rows = []; total = 0
        for param, _ranges in table_def.items():
            val = input_vals[param]; mr = "—"; mp = 0
            for rl, cond, rp in _ranges:
                try:
                    if cond(val): mr = rl; mp = rp; break
                except Exception: pass
            total += mp
            src = "🔵 Auto" if param in prefill else "⬜ Manual"
            note_str = notes.get(param,"")
            rows.append({
                "Sumber": src,
                "Parameter": param,
                "Nilai (%)": val,
                "Range": mr,
                "Rupiah Insentif": mp,
                "Keterangan": note_str,
            })

        rdf = pd.DataFrame(rows)
        mxp = sum(r[0][2] for r in table_def.values())
        pct = total/mxp*100 if mxp > 0 else 0
        col_t = "#39C77B" if pct>=80 else ("#F5C542" if pct>=60 else "#C8102E")
        nl = f" — {html.escape(nama)}" if nama.strip() else ""
        from calendar import month_abbr as _mabb3
        period_label = f"{_mabb3[int(_ins_mn)]} {_ins_yr}" if _ins_mn else "—"

        # Kartu total
        st.markdown(f"""
<div style='background:#0C1D42;border:2px solid {col_t};border-radius:14px;padding:20px 28px;margin-bottom:16px'>
  <div style='color:#FCF1CF;font-size:.85rem;opacity:.8'>TOTAL INSENTIF · {html.escape(period_label)}{html.escape(nl)}</div>
  <div style='color:{col_t};font-size:2.4rem;font-weight:900;margin-top:4px'>Rp {total:,.0f}</div>
  <div style='color:#FCF1CF;font-size:.8rem;opacity:.7;margin-top:6px'>{pct:.1f}% dari maksimum (Rp {mxp:,.0f})</div>
  <div style='margin-top:10px;display:flex;gap:16px;flex-wrap:wrap'>
    <div style='color:#FCF1CF;font-size:.75rem'>🔵 Auto: {sum(1 for p in table_def if p in prefill)} parameter</div>
    <div style='color:#FCF1CF;font-size:.75rem'>⬜ Manual: {len(manual_params)} parameter</div>
  </div>
</div>""",unsafe_allow_html=True)

        # Tabel detail
        def _sty(dff):
            n = len(dff); rc = ["background-color:#F4F6F9" if i%2==0 else "background-color:#E9EEF5" for i in range(n)]
            def _s(s): return [rc[i] for i in range(len(s))]
            styled = (dff.style
                .set_table_styles([
                    {"selector":"th","props":[("background","linear-gradient(180deg,#33509B 0%,#22396F 55%,#182A52 100%)"),("color","#FCF1CF"),("font-weight","800"),("border","1px solid #182A52"),("padding","7px 10px"),("text-align","center")]},
                    {"selector":"td","props":[("border","1px solid #DDE4EC"),("padding","6px 10px")]},
                    {"selector":"table","props":[("border-collapse","collapse"),("width","100%")]},
                ])
                .set_properties(**{"color":"#26364A"}).apply(_s,axis=0)
                .format({"Nilai (%)":"{:.1f}%","Rupiah Insentif":"Rp {:,.0f}"})
                .map(lambda v:"color:#2563EB;font-weight:800" if v=="🔵 Auto" else ("color:#6B7280" if v=="⬜ Manual" else ""),subset=["Sumber"])
                .map(lambda v:"color:#39C77B;font-weight:800" if isinstance(v,int) and v>0 else ("color:#C8102E;font-weight:700" if v==0 else ""),subset=["Rupiah Insentif"])
            )
            return styled
        h = 38+len(rdf)*40
        st.dataframe(_sty(rdf),width="stretch",hide_index=True,height=h)
        st.markdown(f"<div style='text-align:right;font-size:1.1rem;font-weight:900;color:#FCF1CF;background:#22396F;padding:8px 16px;border-radius:8px;margin-top:-8px'>TOTAL → Rp {total:,.0f}</div>",unsafe_allow_html=True)

        # Export
        exp = rdf.copy(); exp.insert(0,"Periode",period_label); exp.insert(1,"Salesman",nama if nama.strip() else "—")
        exp["Total Insentif"] = ""; exp.loc[len(exp)-1,"Total Insentif"] = str(total)
        st.download_button(
            f"⬇ Export Hasil Insentif{title_suffix} (.csv)",
            data=exp.to_csv(index=False).encode("utf-8-sig"),
            file_name=f"insentif_{role_key}_{period_label.replace(' ','_')}_{nama.strip().replace(' ','_') or 'salesman'}.csv",
            mime="text/csv", key=f"export_ins_{role_key}")

        with st.expander("📋 Lihat Tabel Referensi Insentif Lengkap"):
            ref = []
            for p, rngs in table_def.items():
                for rl, _, rp in rngs: ref.append({"Parameter":p,"Range":rl,"Rupiah Insentif":rp})
            rref = pd.DataFrame(ref)
            st.dataframe(rref.style.format({"Rupiah Insentif":"Rp {:,.0f}"}),width="stretch",hide_index=True,height=38+len(rref)*35)

    # ── DEFINISI TABEL ────────────────────────────────────────────────────
    SUPV_TABLE={
        "1. % A/T Gross Sales Cabang (Non Toserba)":[
            ("N ≥ 100%",        lambda n: n>=100,   986_125),("90% ≤ N < 100%",  lambda n: 90<=n<100,778_813),
            ("80% ≤ N < 90%",   lambda n: 80<=n<90, 508_250),("N < 80%",         lambda n: n<80,     374_500),],
        "2. % A/T Sales Toserba":[
            ("N > 100%",        lambda n: n>100,    358_450),("95% < N ≤ 100%",  lambda n: 95<n<=100,275_525),
            ("85% ≤ N < 95%",   lambda n: 85<=n<95, 203_300),("N < 85%",         lambda n: n<85,     149_800),],
        "3. Principal On Target (POT) · Target A/T Wajib 100%":[
            ("N > 60%",         lambda n: n>60,     537_675),("50% ≤ N < 60%",   lambda n: 50<=n<60, 413_288),
            ("40% ≤ N < 50%",   lambda n: 40<=n<50, 304_950),("30% ≤ N < 40%",   lambda n: 30<=n<40, 224_700),],
        "4. Actual EC (Effective Call) Vs Target EC Per Salesman":[
            ("N > 75%",         lambda n: n>75,     537_675),("70% ≤ N ≤ 75%",   lambda n: 70<=n<=75,413_288),
            ("65% ≤ N < 70%",   lambda n: 65<=n<70, 304_950),("N < 65%",         lambda n: n<65,     224_700),],
        "5. % Program PUSH (Sesuai bulan terkait)":[
            ("N ≥ 100%",        lambda n: n>=100,   716_900),("90% ≤ N < 100%",  lambda n: 90<=n<100,551_050),
            ("80% ≤ N < 90%",   lambda n: 80<=n<90, 406_600),("70% ≤ N < 80%",   lambda n: 70<=n<80, 299_600),],
        "6. % Growth RB CHP (dihitung dari AVG RB 2025)":[
            ("N ≥ 15%",         lambda n: n>=15,    179_225),("12.5% ≤ N < 15%", lambda n: 12.5<=n<15,137_763),
            ("10% ≤ N < 12.5%", lambda n: 10<=n<12.5,101_650),("N < 10%",        lambda n: n<10,      74_900),],
        "7. Collection":[
            ("N ≥ 100%",        lambda n: n>=100,   358_450),("98% ≤ N < 100%",  lambda n: 98<=n<100,275_525),
            ("95% ≤ N < 98%",   lambda n: 95<=n<98, 203_300),("N < 95%",         lambda n: n<95,     149_800),],
    }
    ALL_TABLE={
        "1. % A/T Gross Sales Rayon (Non Toserba)":[
            ("N ≥ 100%",        lambda n: n>=100,   570_200),("90% ≤ N < 100%",  lambda n: 90<=n<100,435_000),
            ("80% ≤ N < 90%",   lambda n: 80<=n<90, 262_600),("N < 80%",         lambda n: n<80,     188_000),],
        "2. Actual Effective Call vs Target Call per hari (HK kalender FC 2026)":[
            ("N > 75%",         lambda n: n>75,     247_600),("70% ≤ N < 75%",   lambda n: 70<=n<75, 180_000),
            ("65% ≤ N < 70%",   lambda n: 65<=n<70, 131_300),("N < 65%",         lambda n: n<65,      94_000),],
        "3. Principal On Target (POT) · A/T Wajib 100%":[
            ("N > 60%",         lambda n: n>60,     495_200),("50% ≤ N < 60%",   lambda n: 50<=n<60, 360_000),
            ("40% ≤ N < 50%",   lambda n: 40<=n<50, 262_600),("30% ≤ N < 40%",   lambda n: 30<=n<40, 188_000),],
        "4. % A/T Sales Toserba (CHW)":[
            ("N > 100%",        lambda n: n>100,    247_600),("95% < N ≤ 100%",  lambda n: 95<n<=100,180_000),
            ("90% < N ≤ 95%",   lambda n: 90<n<=95, 131_300),("N < 90%",         lambda n: n<90,      94_000),],
        "5. % Growth RB CHP (dihitung dari AVG RB 2025)":[
            ("N ≥ 15%",         lambda n: n>=15,    123_800),("12.5% ≤ N < 15%", lambda n: 12.5<=n<15, 90_000),
            ("10% ≤ N < 12.5%", lambda n: 10<=n<12.5, 65_650),("N < 10%",        lambda n: n<10,       47_000),],
        "6. % Program PUSH (Sesuai bulan terkait)":[
            ("N ≥ 100%",        lambda n: n>=100,   619_000),("90% ≤ N < 100%",  lambda n: 90<=n<100,450_000),
            ("80% ≤ N < 90%",   lambda n: 80<=n<90, 328_250),("70% ≤ N < 80%",   lambda n: 70<=n<80, 235_000),],
        "7. Collection":[
            ("N ≥ 100%",        lambda n: n>=100,   247_600),("98% ≤ N < 100%",  lambda n: 98<=n<100,180_000),
            ("95% ≤ N < 98%",   lambda n: 95<=n<98, 131_300),("N < 95%",         lambda n: n<95,      94_000),],
    }
    MOD_TABLE={
        "1. % A/T Gross Sales Rayon":[
            ("N ≥ 100%",        lambda n: n>=100,   817_800),("95% ≤ N < 100%",  lambda n: 95<=n<100,615_000),
            ("90% ≤ N < 95%",   lambda n: 90<=n<95, 393_900),("N < 85%",         lambda n: n<85,     282_000),],
        "2. Principal On Target (POT) · A/T Wajib 100%":[
            ("N > 60%",         lambda n: n>60,     371_400),("50% ≤ N < 60%",   lambda n: 50<=n<60, 270_000),
            ("40% ≤ N < 50%",   lambda n: 40<=n<50, 196_950),("30% ≤ N < 40%",   lambda n: 30<=n<40, 141_000),],
        "3. % Program PUSH (Sesuai bulan terkait)":[
            ("N ≥ 100%",        lambda n: n>=100,   371_400),("90% ≤ N < 100%",  lambda n: 90<=n<100,270_000),
            ("80% ≤ N < 90%",   lambda n: 80<=n<90, 196_950),("70% ≤ N < 80%",   lambda n: 70<=n<80, 141_000),],
        "4. Actual Effective Call vs Target Call per hari (HK kalender FC 2026)":[
            ("N > 75%",         lambda n: n>75,     247_600),("70% ≤ N < 75%",   lambda n: 70<=n<75, 180_000),
            ("65% ≤ N < 70%",   lambda n: 65<=n<70, 131_300),("N < 65%",         lambda n: n<65,      94_000),],
        "5. % Growth IBR total (semua principal CHP, vs Avg 3 bulan berjalan)":[
            ("N > 15%",         lambda n: n>15,     247_600),("12% ≤ N < 15%",   lambda n: 12<=n<15, 180_000),
            ("9% ≤ N < 12%",    lambda n: 9<=n<12,  131_300),("N ≥ 5%",          lambda n: n>=5,      94_000),],
        "6. % Growth Coverage / RB (RB bulan berjalan vs avg RB 2025)":[
            ("N > 15%",         lambda n: n>15,     247_600),("12% ≤ N < 15%",   lambda n: 12<=n<15, 180_000),
            ("9% ≤ N < 12%",    lambda n: 9<=n<12,  131_300),("N ≥ 5%",          lambda n: n>=5,      94_000),],
        "7. Collection":[
            ("N ≥ 100%",        lambda n: n>=100,   247_600),("98% ≤ N < 100%",  lambda n: 98<=n<100,180_000),
            ("95% ≤ N < 98%",   lambda n: 95<=n<98, 131_300),("N < 95%",         lambda n: n<95,      94_000),],
    }
    NHS_TABLE={
        "1. % A/T Gross Sales Rayon":[
            ("N ≥ 100%",        lambda n: n>=100,   941_600),("95% ≤ N < 100%",  lambda n: 95<=n<100,705_000),
            ("90% ≤ N < 95%",   lambda n: 90<=n<95, 459_550),("N < 85%",         lambda n: n<85,     329_000),],
        "2. % Growth YTD RPB 2026 vs 2025 Per Rayon":[
            ("N > 12.5%",       lambda n: n>12.5,   371_400),("10% < N ≤ 12.5%", lambda n: 10<n<=12.5,270_000),
            ("7.5% < N ≤ 10%",  lambda n: 7.5<n<=10,196_950),("5% < N ≤ 7.5%",   lambda n: 5<n<=7.5, 141_000),],
        "3. % Growth RB (Actual RB Bulan berjalan 2026 vs AVG RB 2025)":[
            ("N > 15%",         lambda n: n>15,     371_400),("12% ≤ N < 15%",   lambda n: 12<=n<15, 270_000),
            ("9% ≤ N < 12%",    lambda n: 9<=n<12,  196_950),("N ≥ 5%",          lambda n: n>=5,     141_000),],
        "4. % Program PUSH":[
            ("N > 75%",         lambda n: n>75,     371_400),("70% ≤ N < 75%",   lambda n: 70<=n<75, 270_000),
            ("65% ≤ N < 70%",   lambda n: 65<=n<70, 196_950),("N < 65%",         lambda n: n<65,     141_000),],
        "5. Growth Sales Produk Fokus (Produk Fokus sesuai sasaran per bulan)":[
            ("N ≥ 10%",         lambda n: n>=10,    247_600),("9% ≤ N < 10%",    lambda n: 9<=n<10,  180_000),
            ("8% ≤ N < 9%",     lambda n: 8<=n<9,   131_300),("N < 8%",          lambda n: n<8,       94_000),],
        "6. Collection":[
            ("N ≥ 100%",        lambda n: n>=100,   247_600),("98% ≤ N < 100%",  lambda n: 98<=n<100,180_000),
            ("95% ≤ N < 98%",   lambda n: 95<=n<98, 131_300),("N < 95%",         lambda n: n<95,      94_000),],
    }
    # key parameter A/T yang bisa di-prefill
    PREFILL_KEY={"SUPV":"1. % A/T Gross Sales Cabang (Non Toserba)",
                 "ALL_REG":"1. % A/T Gross Sales Rayon (Non Toserba)",
                 "MOD":"1. % A/T Gross Sales Rayon","NHS":"1. % A/T Gross Sales Rayon"}

    # ════════════════════════════════════════════════════════════════════
    # SUPERVISOR
    # ════════════════════════════════════════════════════════════════════
    if insentif_role=="Supervisor":
        st.markdown("### Insentif Supervisor")
        sc=col(df,"salesperson") if df is not None else None
        all_sp=sorted(df[sc].dropna().astype(str).unique().tolist()) if sc and df is not None else []
        fc1,fc2=st.columns([2,3])
        with fc1:
            sel_supv=st.selectbox("Pilih Sales Person Title (opsional)",["— semua / manual —"]+all_sp,key="ins_supv_sp")
        # filter tahun & bulan dari Performance Sales jika tersedia
        prefill_supv={}; notes_supv={}
        if sel_supv and sel_supv!="— semua / manual —":
            prefill_supv, notes_supv = calc_insentif_prefill(sel_supv, _ins_yr, _ins_mn, "SUPV")
            if prefill_supv:
                with fc2:
                    info_lines=[f"**{k.split('.')[0]}:** {v:.1f}%" for k,v in prefill_supv.items()]
                    st.success(f"✅ {len(prefill_supv)} parameter dihitung otomatis · " + " · ".join(info_lines))
        render_kalkulator(SUPV_TABLE,"supervisor"," Supervisor",prefill=prefill_supv,notes=notes_supv)

    # ════════════════════════════════════════════════════════════════════
    # SALESMAN CHP
    # ════════════════════════════════════════════════════════════════════
    else:
        st.markdown("### Insentif Salesman CHP")
        sc=col(df,"salesperson") if df is not None else None
        chp_titles=sorted([t for t in (df[sc].dropna().astype(str).unique().tolist() if sc and df is not None else []) if "CHP" in t.upper()])
        fc1,fc2=st.columns([2,3])
        with fc1:
            sel_chp=st.selectbox("Pilih Sales Person Title",["— pilih title —"]+chp_titles,key="ins_chp_sp")
        # auto-detect tipe
        detected="ALL_REG"
        if sel_chp and sel_chp!="— pilih title —": detected=detect_sp_type(sel_chp)
        type_options=["ALL / REG  (ALL-1, ALL-2, REG-1-A, REG-2-A)","MOD  (MOD-1-A, MOD-2-A)","NHS  (NHS-1-A)"]
        type_idx={"ALL_REG":0,"MOD":1,"NHS":2}.get(detected,0)
        sp_type=st.radio("Tipe Salesman",type_options,index=type_idx,horizontal=True,key="ins_sp_type")
        type_key="ALL_REG" if sp_type.startswith("ALL") else ("MOD" if sp_type.startswith("MOD") else "NHS")
        # prefill semua parameter yang bisa dihitung dari database
        prefill_chp={}; notes_chp={}
        if sel_chp and sel_chp!="— pilih title —":
            prefill_chp, notes_chp = calc_insentif_prefill(sel_chp, _ins_yr, _ins_mn, type_key)
            if prefill_chp:
                with fc2:
                    info_lines=[f"**{k.split('.')[0]}:** {v:.1f}%" for k,v in prefill_chp.items()]
                    st.success(f"✅ {len(prefill_chp)} parameter dihitung otomatis · " + " · ".join(info_lines))
        else:
            with fc2: st.caption("Pilih Sales Person Title untuk menghitung parameter otomatis dari database.")
        if type_key=="ALL_REG":
            st.caption("Parameter: AMBDGSCHP ALL-1, ALL-2, REG-1-A, REG-2-A")
            render_kalkulator(ALL_TABLE,"all_reg"," ALL/REG",prefill=prefill_chp,notes=notes_chp)
        elif type_key=="MOD":
            st.caption("Parameter: AMBDGSCHP MOD-1-A, MOD-2-A")
            render_kalkulator(MOD_TABLE,"mod"," MOD",prefill=prefill_chp,notes=notes_chp)
        else:
            st.caption("Parameter: AMBDGSCHP NHS-1-A")
            render_kalkulator(NHS_TABLE,"nhs"," NHS",prefill=prefill_chp,notes=notes_chp)

        # ════════════════════════════════════════════════════════════════════
        # SUMMARY ALL SALESMAN CHP
        # ════════════════════════════════════════════════════════════════════
        st.divider()
        st.markdown("### 📋 Summary Insentif All Salesman CHP")
        st.caption(f"Dihitung otomatis untuk semua Sales Person Title CHP · {_mabb[_ins_mn]} {_ins_yr}. "
                   "Parameter yang bisa dihitung dari database diisi otomatis; parameter lain (Collection, EC, dll) diasumsikan 0 untuk keperluan estimasi ranking.")

        # Tabel range untuk lookup rupiah (pakai table sesuai tipe)
        TABLE_BY_TYPE = {"ALL_REG": ALL_TABLE, "MOD": MOD_TABLE, "NHS": NHS_TABLE}
        MAX_BY_TYPE   = {t: sum(r[0][2] for r in tbl.values()) for t,tbl in TABLE_BY_TYPE.items()}

        def _match_rupiah(table_def, param_vals):
            total = 0
            breakdown = {}
            for param, ranges in table_def.items():
                val = param_vals.get(param, 0.0)
                for rl, cond, rp in ranges:
                    try:
                        if cond(val): breakdown[param]=(val,rl,rp); total+=rp; break
                    except Exception: pass
                else:
                    breakdown[param]=(val,"—",0)
            return total, breakdown

        summary_rows = []
        all_breakdowns = {}
        for title in chp_titles:
            tt = detect_sp_type(title)
            tbl = TABLE_BY_TYPE.get(tt, ALL_TABLE)
            pf, nt = calc_insentif_prefill(title, _ins_yr, _ins_mn, tt)
            total_rp, breakdown = _match_rupiah(tbl, pf)
            max_rp = MAX_BY_TYPE.get(tt, 1)
            pct_max = round(total_rp / max_rp * 100, 1) if max_rp > 0 else 0.0
            # Identifikasi parameter yang paling perlu ditingkatkan
            # = parameter auto (ada di pf) dengan rupiah terendah dibanding potensi max-nya
            gap_params = []
            for param, ranges in tbl.items():
                if param not in pf: continue          # skip manual params
                val = pf.get(param, 0.0)
                max_rp_param = ranges[0][2]           # range pertama = nilai tertinggi
                actual_rp = breakdown.get(param, (0,0,0))[2]
                gap = max_rp_param - actual_rp
                if gap > 0:
                    gap_params.append((gap, param.split(".")[0].strip(), val, actual_rp, max_rp_param))
            gap_params.sort(reverse=True)
            saran = "\n".join(
                f"P{g[1]}: nilai {g[2]:.0f}% → Rp {g[3]:,.0f} (gap Rp {g[0]:,.0f} dari maks Rp {g[4]:,.0f})"
                for g in gap_params[:4]
            ) if gap_params else "✅ Semua parameter maks"
            summary_rows.append({
                "Sales Person Title": title,
                "Tipe": tt,
                "Total Insentif": total_rp,
                "% dari Maks": pct_max,
                "Parameter Auto": len(pf),
                "Saran Fokus": saran,
            })
            all_breakdowns[title] = (tt, pf, breakdown, tbl)

        if summary_rows:
            sum_df = pd.DataFrame(summary_rows).sort_values("Total Insentif", ascending=False).reset_index(drop=True)

            # KPI cards ringkasan
            kc1,kc2,kc3,kc4 = st.columns(4)
            kc1.metric("Total Salesman CHP", f"{len(sum_df):,}")
            kc2.metric("Rata-rata Insentif", f"Rp {sum_df['Total Insentif'].mean():,.0f}")
            kc3.metric("Insentif Tertinggi",  f"Rp {sum_df['Total Insentif'].max():,.0f}")
            kc4.metric("Insentif Terendah",   f"Rp {sum_df['Total Insentif'].min():,.0f}")

            # Tabel utama — render HTML agar kolom Saran Fokus bisa wrap teks penuh
            sum_df_display = sum_df.drop(columns=["Parameter Auto"])

            def _fmt_insentif(v):
                try: return f"Rp {int(v):,}"
                except: return str(v)
            def _fmt_pct(v):
                try: return f"{float(v):.1f}%"
                except: return str(v)

            html_rows = ""
            for _, r in sum_df_display.iterrows():
                pct_v = float(r["% dari Maks"])
                bar_color = "#39C77B" if pct_v >= 80 else ("#F5C542" if pct_v >= 60 else "#C8102E")
                html_rows += (
                    f"<tr>"
                    f"<td style='width:18%;font-weight:700'>{r['Sales Person Title']}</td>"
                    f"<td style='width:7%;text-align:center'>{r['Tipe']}</td>"
                    f"<td style='width:13%;text-align:right'>{_fmt_insentif(r['Total Insentif'])}</td>"
                    f"<td style='width:9%;text-align:center'>"
                    f"  <span style='color:{bar_color};font-weight:800'>{_fmt_pct(pct_v)}</span>"
                    f"</td>"
                    f"<td style='width:53%;white-space:pre-wrap;word-break:break-word;font-size:.78rem;line-height:1.45'>"
                    f"{r['Saran Fokus']}</td>"
                    f"</tr>"
                )

            html_tbl = f"""
<style>
.ins-summary-tbl {{width:100%;border-collapse:collapse;font-size:.82rem;}}
.ins-summary-tbl th {{
    background:linear-gradient(180deg,#335A9F 0%,#24477F 100%);
    color:#FFFFFF;font-weight:800;padding:4px 7px;border:1px solid #1D3B70;
    text-align:left;white-space:nowrap;
}}
.ins-summary-tbl td {{
    border:1px solid #DDE4EC;padding:3px 7px;vertical-align:top;
    background:#F4F7FB;color:#1B2E46;
}}
.ins-summary-tbl tr:nth-child(even) td {{background:#E8EEF6;}}
</style>
<div style='overflow-x:auto;border-radius:10px;box-shadow:0 6px 18px rgba(0,7,55,.30)'>
<table class='ins-summary-tbl'>
<thead><tr>
  <th>Sales Person Title</th><th>Tipe</th><th>Total Insentif</th>
  <th>% dari Maks</th><th>Saran Fokus (parameter dengan gap terbesar)</th>
</tr></thead>
<tbody>{html_rows}</tbody>
</table></div>"""
            st.markdown(html_tbl, unsafe_allow_html=True)

            # Export
            st.download_button(
                "⬇ Export Summary All Salesman CHP (.csv)",
                data=sum_df.to_csv(index=False).encode("utf-8-sig"),
                file_name=f"summary_insentif_chp_{_mabb[_ins_mn]}_{_ins_yr}.csv",
                mime="text/csv",
                key="export_summary_chp",
            )

        else:
            st.info("Tidak ada Sales Person Title CHP yang ditemukan di database.")


elif main_section=="Report":
    st.markdown("<div class='score-heading'>◇ REPORT GENERATOR</div>",unsafe_allow_html=True)
    report_submenu=st.radio("Sub Menu Report",["Report Generator","Report Harian","Report Bulanan","Slide Meeting Kasie","AI Chat"],horizontal=True,key="report_submenu")
    if report_submenu=="AI Chat":
        st.markdown("#### 💬 AI Chat · Analisa Database & Strategi")
        st.caption("Tanyakan performa, detail bisnis, principal, salesperson, pelanggan, program PUSH, tren, atau strategi. Pilih OpenAI GPT untuk analisa lebih mendalam atau Gemini. Pertanyaan dan ringkasan agregat data dashboard diproses melalui penyedia AI yang dipilih.")
        _report_ai_date_col=col(df,"date")
        _report_ai_dates=pd.to_datetime(df[_report_ai_date_col],errors="coerce") if _report_ai_date_col else pd.Series(pd.NaT,index=df.index)
        _report_ai_years=sorted(_report_ai_dates.dropna().dt.year.astype(int).unique().tolist())
        _report_ai_year_default=int(st.session_state.get("kasie_report_year",_report_ai_years[-1] if _report_ai_years else pd.Timestamp.now().year))
        _report_ai_year_default=_report_ai_year_default if _report_ai_year_default in _report_ai_years else (_report_ai_years[-1] if _report_ai_years else pd.Timestamp.now().year)
        _report_ai_filters=st.columns([2,1,1])
        with _report_ai_filters[0]:
            _report_ai_bg_col=col(df,"business_group")
            _report_ai_bg_options=sorted(df[_report_ai_bg_col].dropna().astype(str).str.strip().unique().tolist()) if _report_ai_bg_col else []
            _report_ai_groups=st.multiselect("Bisnis Group untuk analisa",_report_ai_bg_options,key="report_ai_business_groups",placeholder="Semua Bisnis Group")
        with _report_ai_filters[1]:
            _report_ai_year=st.selectbox("Tahun",_report_ai_years or [_report_ai_year_default],index=(_report_ai_years.index(_report_ai_year_default) if _report_ai_year_default in _report_ai_years else 0),key="report_ai_year")
        with _report_ai_filters[2]:
            _report_ai_months=sorted(_report_ai_dates.loc[_report_ai_dates.dt.year.eq(int(_report_ai_year))].dropna().dt.month.astype(int).unique().tolist()) or list(MONTHS)
            _report_ai_month_default=int(st.session_state.get("kasie_report_month",_report_ai_months[-1]))
            _report_ai_month_default=_report_ai_month_default if _report_ai_month_default in _report_ai_months else _report_ai_months[-1]
            _report_ai_month=st.selectbox("Bulan",_report_ai_months,index=_report_ai_months.index(_report_ai_month_default),format_func=lambda value:MONTHS.get(int(value),str(value)),key="report_ai_month")
        _report_ai_provider=st.selectbox("Penyedia AI",["OpenAI (GPT)","Google Gemini"],key="report_ai_provider")
        if _report_ai_provider=="OpenAI (GPT)":
            _report_ai_key=st.text_input("OpenAI API key",type="password",value=st.session_state.get("openai_api_key","") or os.environ.get("OPENAI_API_KEY",""),key="report_openai_api_key",help="API key hanya disimpan di sesi browser ini dan tidak ditulis ke file pengaturan.")
            if _report_ai_key: st.session_state["openai_api_key"]=_report_ai_key
            _report_ai_model_options=["gpt-5.1"]
            _report_ai_model_key="report_openai_model"
        else:
            _report_ai_key=st.text_input("Gemini API key",type="password",value=st.session_state.get("gemini_api_key","") or os.environ.get("GEMINI_API_KEY","") or os.environ.get("GOOGLE_API_KEY",""),key="report_gemini_api_key",help="API key hanya disimpan di sesi browser ini dan tidak ditulis ke file pengaturan.")
            if _report_ai_key: st.session_state["gemini_api_key"]=_report_ai_key
            _report_ai_model_options=["gemini-3.8-flash"]
            _report_ai_model_key="report_gemini_model"
        if st.session_state.get(_report_ai_model_key) not in _report_ai_model_options:
            st.session_state[_report_ai_model_key]=_report_ai_model_options[0]
        _report_ai_model=st.selectbox("Model AI",_report_ai_model_options,key=_report_ai_model_key)

        def _report_ai_fmt(value):
            return "—" if value is None or pd.isna(value) else f"{float(value):,.0f}"

        def _report_ai_context():
            _lines=[f"Dashboard CHP AAM Cabang {ACTIVE_BRANCH_NAME} ({ACTIVE_BRANCH_CODE}). Gunakan hanya data yang tersedia; jangan mengarang. Angka sales dalam mata uang database."]
            _dcol=col(df,"date"); _scol=col(df,"sales"); _tcol=col(df,"target_value")
            _bgcol=col(df,"business_group"); _prcol=col(df,"principal"); _spcol=col(df,"salesperson"); _custcol=col(df,"customer")
            _d=pd.to_datetime(df[_dcol],errors="coerce") if _dcol else pd.Series(pd.NaT,index=df.index)
            _year=int(st.session_state.get("report_ai_year",_d.dropna().dt.year.max() if _d.notna().any() else pd.Timestamp.now().year))
            _month=int(st.session_state.get("report_ai_month",_d.dropna().dt.month.max() if _d.notna().any() else pd.Timestamp.now().month))
            _mask=_d.dt.year.eq(_year)&_d.dt.month.le(_month)
            if _bgcol and st.session_state.get("report_ai_business_groups"):
                _mask=_mask & df[_bgcol].astype(str).str.strip().isin(st.session_state["report_ai_business_groups"])
            _work=pd.DataFrame({"Actual":pd.to_numeric(df[_scol],errors="coerce").fillna(0) if _scol else 0,
                                "Target":pd.to_numeric(df[_tcol],errors="coerce").fillna(0) if _tcol else 0,
                                "Bisnis Group":df[_bgcol].astype("string").fillna("(Kosong)").str.strip() if _bgcol else "(Kolom tidak tersedia)",
                                "Principal ID":df[_prcol].astype("string").fillna("(Kosong)").str.strip() if _prcol else "(Kolom tidak tersedia)",
                                "Sales Person Title":df[_spcol].astype("string").fillna("(Kosong)").str.strip() if _spcol else "(Kolom tidak tersedia)",
                                "Customer":df[_custcol].astype("string").fillna("(Kosong)").str.strip() if _custcol else "(Kolom tidak tersedia)",
                                "Bulan":_d.dt.month})
            _period=_work[_mask].copy(); _period=_period[_period["Bulan"].eq(_month)]
            _ytd=_work[_mask]
            _month_name=MONTHS.get(_month,str(_month))
            _lines.append(f"Periode terpilih: {_month_name} {_year}; cakupan YTD Januari sampai {_month_name} {_year}.")
            for _label,_scope in [("Bulanan",_period),("YTD",_ytd)]:
                _actual=float(_scope["Actual"].sum()); _target=float(_scope["Target"].sum())
                _lines.append(f"{_label}: Actual {_report_ai_fmt(_actual)}; Target {_report_ai_fmt(_target)}; A/T {(_actual/_target*100 if _target else 0):.1f}%; Deviasi {_report_ai_fmt(_actual-_target)}.")
            for _dim in ["Bisnis Group","Principal ID","Sales Person Title"]:
                _agg=_period.groupby(_dim,dropna=False).agg(Actual=("Actual","sum"),Target=("Target","sum")).reset_index()
                _agg=_agg[(_agg["Actual"]!=0)|(_agg["Target"]!=0)].copy()
                _agg["Deviasi"]=_agg["Actual"]-_agg["Target"]
                _agg["A/T"]=_agg["Actual"].div(_agg["Target"].replace(0,float("nan")))*100
                _agg=_agg.sort_values("Actual",ascending=False).head(30)
                _lines.append(f"Rincian bulanan per {_dim} (maksimum 30 baris; Actual, Target, A/T, Deviasi):")
                for _,_row in _agg.iterrows():
                    _at_text="—" if pd.isna(_row["A/T"]) else f"{_row['A/T']:.1f}%"
                    _lines.append(f"- {_row[_dim]} | {_report_ai_fmt(_row['Actual'])} | {_report_ai_fmt(_row['Target'])} | {_at_text} | {_report_ai_fmt(_row['Deviasi'])}")
            _customers=_period.groupby("Customer",dropna=False)["Actual"].sum().sort_values(ascending=False).head(20)
            _lines.append("Pelanggan dengan actual tertinggi bulan ini:")
            _lines.extend([f"- {name}: {_report_ai_fmt(value)}" for name,value in _customers.items()])
            _lines.append("Program PUSH sesuai konfigurasi aktif:")
            for _program in programs:
                _selected=settings.get("filters",{}).get(_program,{})
                if not _selected: continue
                try:
                    _,_actual,_target,_at,_dev,_=calculate_program_kpi(_program,_selected)
                    _lines.append(f"- {program_label(_program)} | metric {_selected.get('metric','')} | Actual {_report_ai_fmt(_actual)} | Target {_report_ai_fmt(_target)} | A/T {('—' if _at is None else f'{_at:.1f}%')} | Deviasi {_report_ai_fmt(_dev)}")
                except Exception as _err:
                    _lines.append(f"- {program_label(_program)}: ringkasan tidak tersedia ({_err}).")
            return "\n".join(_lines)[:24000]

        for _msg in st.session_state.setdefault("report_ai_messages",[]):
            with st.chat_message("user" if _msg["role"]=="user" else "assistant"):
                st.markdown(_msg["text"])
        with st.form("report_ai_chat_form",clear_on_submit=True):
            _report_ai_prompt=st.text_area("Pertanyaan",placeholder="Contoh: Principal mana yang paling menekan pencapaian bulan ini dan strategi apa yang disarankan?",height=100)
            _report_ai_send=st.form_submit_button("Kirim ke AI",type="primary",width="stretch")
        if _report_ai_send and _report_ai_prompt.strip():
            if not _report_ai_key:
                st.warning("Masukkan API key untuk penyedia AI yang dipilih.")
            else:
                _messages=st.session_state["report_ai_messages"]
                _messages.append({"role":"user","text":_report_ai_prompt.strip()})
                _contents=[{"role":"user" if m["role"]=="user" else "model","parts":[{"text":m["text"]}]} for m in _messages[-20:]]
                _context=_report_ai_context()
                _instructions=f"Anda adalah analis penjualan CHP AAM Cabang {ACTIVE_BRANCH_NAME}. Jawab dalam bahasa Indonesia berdasarkan data konteks saja. Berikan angka actual, target, A/T dan deviasi jika relevan. Untuk analisa bisnis, jelaskan penyebab yang didukung data lalu berikan strategi praktis. Bedakan fakta dan rekomendasi; jika data tidak cukup, jelaskan batasannya.\n\nRingkasan database dashboard:\n"+_context
                try:
                    with st.spinner("AI sedang membaca ringkasan database dan menyusun analisa..."):
                        if _report_ai_provider=="OpenAI (GPT)":
                            _openai_input=[{"role":"user" if m["role"]=="user" else "assistant","content":m["text"]} for m in _messages[-20:]]
                            _openai_body={"model":_report_ai_model,"instructions":_instructions,"input":_openai_input,"reasoning":{"effort":"high"},"max_output_tokens":1800}
                            _response=requests.post("https://api.openai.com/v1/responses",headers={"Authorization":f"Bearer {_report_ai_key}","Content-Type":"application/json"},json=_openai_body,timeout=120)
                        else:
                            _gemini_body={"systemInstruction":{"parts":[{"text":_instructions}]},"contents":_contents}
                            _response=requests.post(f"https://generativelanguage.googleapis.com/v1beta/models/{_report_ai_model}:generateContent",headers={"x-goog-api-key":_report_ai_key},json=_gemini_body,timeout=90)
                    if _response.status_code!=200:
                        st.error(f"{_report_ai_provider} gagal merespons (HTTP {_response.status_code}): {_response.text[:400]}")
                        _messages.pop()
                    else:
                        if _report_ai_provider=="OpenAI (GPT)":
                            _openai_result=_response.json(); _reply=_openai_result.get("output_text","")
                            if not _reply:
                                _reply="\n".join(part.get("text","") for item in _openai_result.get("output",[]) for part in item.get("content",[]) if part.get("type")=="output_text")
                        else:
                            _parts=_response.json().get("candidates",[{}])[0].get("content",{}).get("parts",[])
                            _reply="".join(p.get("text","") for p in _parts)
                        _reply=_reply or "AI tidak mengembalikan teks."
                        _messages.append({"role":"assistant","text":_reply})
                        st.rerun()
                except Exception as _chat_error:
                    _messages.pop()
                    st.error(f"Tidak dapat terhubung ke Gemini: {_chat_error}")
        if st.button("Bersihkan percakapan AI",key="clear_report_ai_chat"):
            st.session_state["report_ai_messages"]=[]
            st.rerun()
        st.stop()
    elif report_submenu=="Report Harian":
        st.markdown(f"#### 🗓️ Report Harian · Cabang {ACTIVE_BRANCH_NAME}")
        st.caption("Upload kedisiplinan per posisi, lalu lengkapi target DOT dan isu/insight sebelum generate report.")
        report_date=st.date_input("Tanggal report",value=pd.Timestamp.now().date(),key="daily_report_date")
        report_date_key=report_date.isoformat()
        _daily_report_settings=settings.setdefault("daily_report_inputs",{}).setdefault(report_date_key,{})
        _daily_bg_col=col(df,"business_group")
        _daily_bg_options=sorted(df[_daily_bg_col].dropna().astype(str).str.strip().unique().tolist()) if _daily_bg_col else []
        _daily_bg_choices=["Semua Bisnis Group"]+_daily_bg_options
        _daily_bg_saved=_daily_report_settings.get("business_group","Semua Bisnis Group")
        _daily_bg_default=_daily_bg_saved if _daily_bg_saved in _daily_bg_choices else "Semua Bisnis Group"
        daily_business_group=st.selectbox("Bisnis Group report harian",_daily_bg_choices,index=_daily_bg_choices.index(_daily_bg_default),key=f"daily_report_business_group_{report_date_key}")

        st.markdown("##### 1. Kedisiplinan")
        discipline_upload=st.file_uploader("Upload data kedisiplinan (.xlsx / .csv)",type=["xlsx","xls","csv"],key="daily_discipline_upload")
        _discipline_template=pd.DataFrame(columns=["CBG","Salesman Position","TGT RC","ACT"])
        st.download_button("Unduh template kedisiplinan CSV",_discipline_template.to_csv(index=False).encode("utf-8-sig"),"template_kedisiplinan_report_harian.csv","text/csv",key="daily_discipline_template")
        discipline_df=None
        if discipline_upload is not None:
            try:
                discipline_df=read_uploaded_file(discipline_upload.name,discipline_upload.getvalue())
                discipline_lookup={key(c):c for c in discipline_df.columns}
                def _discipline_col(names):
                    return next((discipline_lookup[key(name)] for name in names if key(name) in discipline_lookup),None)
                _position_col=_discipline_col(["Salesman Position","Salesman Position Name","Sales Person Position","Position","CBG / Salesman Position","Sales Person Title"])
                _target_rc_col=_discipline_col(["TGT RC","Target RC","TGT","Target"])
                _actual_rc_col=_discipline_col(["ACT","Actual","Actual RC"])
                _cbg_col=_discipline_col(["CBG","Cabang"])
                _discipline_missing=[label for label,col_value in [("Salesman Position",_position_col),("TGT RC",_target_rc_col),("ACT",_actual_rc_col)] if not col_value]
                if _discipline_missing:
                    st.error("Kolom upload kedisiplinan belum lengkap: "+", ".join(_discipline_missing))
                    discipline_df=None
                else:
                    _discipline_rename={_position_col:"Salesman Position",_target_rc_col:"TGT RC",_actual_rc_col:"ACT"}
                    if _cbg_col: _discipline_rename[_cbg_col]="CBG"
                    discipline_df=discipline_df.rename(columns=_discipline_rename)
                    _discipline_keep=[c for c in ["CBG","Salesman Position","TGT RC","ACT"] if c in discipline_df.columns]
                    discipline_df=discipline_df[_discipline_keep].copy()
                    discipline_df["TGT RC"]=pd.to_numeric(discipline_df["TGT RC"],errors="coerce").fillna(0)
                    discipline_df["ACT"]=pd.to_numeric(discipline_df["ACT"],errors="coerce").fillna(0)
                    DAILY_DISCIPLINE_PATH.parent.mkdir(parents=True,exist_ok=True)
                    _discipline_csv=discipline_df.to_csv(index=False).encode("utf-8-sig")
                    DAILY_DISCIPLINE_PATH.write_bytes(_discipline_csv)
                    if st.secrets.get("gcp_service_account"):
                        try: _drive_upload_bytes(DAILY_DISCIPLINE_PATH.name,_discipline_csv,"text/csv","reports",ACTIVE_BRANCH_CODE)
                        except Exception as exc: st.error(f"Laporan tersimpan sementara, tetapi gagal disinkronkan ke Google Drive: {exc}")
            except Exception as _discipline_error:
                st.error(f"File kedisiplinan tidak dapat dibaca: {_discipline_error}")
        if discipline_df is None:
            _discipline_saved=DAILY_DISCIPLINE_PATH
            if _discipline_saved.exists():
                try: discipline_df=pd.read_csv(_discipline_saved)
                except Exception: discipline_df=None
        if discipline_df is None:
            discipline_df=pd.DataFrame(columns=["CBG","Salesman Position","TGT RC","ACT"])
            st.info("Upload file dengan kolom Salesman Position, TGT RC, dan ACT untuk mengisi bagian kedisiplinan.")
        if not discipline_df.empty:
            _discipline_signature=hashlib.sha256(discipline_df.to_csv(index=False).encode("utf-8")).hexdigest()[:12]
            st.caption(f"{len(discipline_df):,} posisi kedisiplinan siap dimasukkan ke report.")
            discipline_df=st.data_editor(discipline_df,width="stretch",hide_index=True,num_rows="dynamic",key=f"daily_discipline_editor_{report_date_key}_{_discipline_signature}",column_config={"TGT RC":st.column_config.NumberColumn(format="%,.0f"),"ACT":st.column_config.NumberColumn(format="%,.0f")})

        st.markdown("##### 2. Program PUSH · Actual mengikuti Summary Program")
        push_daily_rows=[]
        _report_month_name=MONTHS[report_date.month]
        for _report_program in PROGRAMS:
            _report_selected={"year":[],"month":[],"principal_name":[],"principal":[],"channel_group":[],"business_group":[],"product_code":[],"product":[],"metric":"Total Relasi Beli Transaksi"}
            _report_selected=apply_active_program_config(_report_program,_report_selected,config_month=_report_month_name)
            _report_selected["year"]=[str(report_date.year)]
            _report_selected["month"]=[_report_month_name]
            _report_scoped,_report_actual,_report_target,_report_achievement,_report_deviation,_=calculate_program_kpi(_report_program,_report_selected)
            push_daily_rows.append({"Nama program":program_label(_report_program,month=_report_month_name),"Metric":_report_selected["metric"],"Target":_report_target or 0,"Actual":_report_actual or 0,"A/T":_report_achievement})
        push_daily_df=pd.DataFrame(push_daily_rows)
        if not push_daily_df.empty:
            st.dataframe(push_daily_df.drop(columns=["Metric"]),width="stretch",hide_index=True,column_config={"Target":st.column_config.NumberColumn(format="%,.0f"),"Actual":st.column_config.NumberColumn(format="%,.0f"),"A/T":st.column_config.NumberColumn(format="%.1f%%")})

        st.markdown("##### 3. Sales KUM Gross")
        _report_sales_col=col(df,"sales"); _report_date_col=col(df,"date"); _report_target_value_col=col(df,"target_value")
        sales_kum_actual=0.0; sales_kum_target=0.0; dot_daily_sales_actual=0.0
        if _report_sales_col and _report_date_col:
            _report_sales_data=df.copy()
            _report_dates=pd.to_datetime(_report_sales_data[_report_date_col],errors="coerce")
            _report_period_mask=(_report_dates.dt.year.eq(report_date.year)&_report_dates.dt.month.eq(report_date.month)&_report_dates.dt.date.le(report_date))
            _perf_daily_filter=settings.get("performance_filters",{})
            _report_bg_col=col(_report_sales_data,"business_group")
            _report_sp_col=col(_report_sales_data,"salesperson")
            _report_sp_values=_perf_daily_filter.get("salesperson",[])
            if daily_business_group!="Semua Bisnis Group" and _report_bg_col:
                _report_period_mask &= _report_sales_data[_report_bg_col].astype(str).str.strip().eq(daily_business_group)
            if _report_sp_values and _report_sp_col: _report_period_mask &= _report_sales_data[_report_sp_col].astype(str).isin(_report_sp_values)
            _report_sales_scope=_report_sales_data.loc[_report_period_mask]
            sales_kum_actual=float(pd.to_numeric(_report_sales_scope[_report_sales_col],errors="coerce").fillna(0).sum())
            _report_day_scope=_report_sales_data.loc[_report_period_mask & _report_dates.dt.date.eq(report_date)]
            dot_daily_sales_actual=float(pd.to_numeric(_report_day_scope[_report_sales_col],errors="coerce").fillna(0).sum())
            if _report_target_value_col:
                sales_kum_target=float(pd.to_numeric(_report_sales_scope[_report_target_value_col],errors="coerce").fillna(0).sum())
        sales_kum_achievement=sales_kum_actual/sales_kum_target*100 if sales_kum_target else None
        if not _report_sales_col or not _report_date_col:
            st.warning("Kolom Sales Value Gross atau tanggal tidak ditemukan; nilai Sales KUM Gross akan menjadi 0.")
        sales_kum_cols=st.columns(3)
        sales_kum_cols[0].metric("Target Sales KUM Gross",f"{sales_kum_target:,.0f}")
        sales_kum_cols[1].metric("Actual Sales KUM Gross",f"{sales_kum_actual:,.0f}")
        sales_kum_cols[2].metric("A/T Sales KUM Gross",f"{sales_kum_achievement:,.1f}%" if sales_kum_achievement is not None else "—")

        st.markdown("##### 4. DOT")
        _dot_bg_col=col(df,"business_group")
        _dot_date_col=col(df,"date")
        _dot_sales_col=col(df,"sales")
        _dot_target_col=col(df,"target_value")
        dot_daily_sales_actual=0.0
        dot_monthly_target=0.0
        if _dot_date_col:
            _dot_dates=pd.to_datetime(df[_dot_date_col],errors="coerce")
            _dot_month_mask=_dot_dates.dt.year.eq(report_date.year)&_dot_dates.dt.month.eq(report_date.month)
            _dot_day_mask=_dot_month_mask&_dot_dates.dt.date.eq(report_date)
            if _dot_bg_col and daily_business_group!="Semua Bisnis Group":
                _dot_bg_match=df[_dot_bg_col].astype(str).str.strip().eq(daily_business_group)
                _dot_month_mask &= _dot_bg_match
                _dot_day_mask &= _dot_bg_match
            if _dot_sales_col:
                dot_daily_sales_actual=float(pd.to_numeric(df.loc[_dot_day_mask,_dot_sales_col],errors="coerce").fillna(0).sum())
            if _dot_target_col:
                dot_monthly_target=float(pd.to_numeric(df.loc[_dot_month_mask,_dot_target_col],errors="coerce").fillna(0).sum())
        dot_target=dot_monthly_target*0.05
        dot_actual=dot_daily_sales_actual
        _dot_count_default=int(_daily_report_settings.get("dot_count",0))
        _dot_cols=st.columns(4)
        with _dot_cols[0]: st.metric("Target DOT harian · 5% target bulanan",f"{dot_target:,.0f}")
        with _dot_cols[1]: st.metric("Actual sales hari berjalan",f"{dot_actual:,.0f}")
        with _dot_cols[2]: dot_achievement=dot_actual/dot_target*100 if dot_target else None; st.metric("DOT · Actual / Target",f"{dot_achievement:.1f}%" if dot_achievement is not None else "—")
        with _dot_cols[3]: dot_count=st.number_input("Jumlah DOT",min_value=0,value=_dot_count_default,step=1,key=f"daily_dot_count_{report_date_key}")
        st.caption(f"Bisnis Group report: {daily_business_group}. Target DOT = 5% dari total Target Value bulanan grup ini. Actual DOT = Sales Value Gross pada tanggal report untuk grup ini.")

        st.markdown("##### 5. Isu dan insight")
        _issues_value=_daily_report_settings.get("issues","")
        issues_text=st.text_area("Isi satu isu/insight per baris",value=_issues_value,height=150,key=f"daily_report_issues_{report_date_key}",placeholder="Contoh: RC kejar minimal 98%\nKebutuhan produk belum terpenuhi karena stok kosong")

        if st.button("✨ Generate Report Harian",type="primary",key="generate_daily_report"):
            def _daily_money(value):
                return f"{value/1_000_000:,.1f}Jt" if abs(value)>=1_000_000 else f"{value:,.0f}"
            _daily_lines=[
                "Selamat Sore, Pak Tawi dan Pak Dhiki.",
                "",
                f"Berikut update one on one dari cab. BDG, {report_date:%d %b %Y}.",
                "",
                "Format:",
                "",
                "1. Kedisiplinan : CBG/ Salesman Position / TGT RC / ACT / A/T (KUM)",
            ]
            for _,_discipline_row in discipline_df.iterrows():
                _cbg_value=_discipline_row.get("CBG","")
                _position_value=_discipline_row.get("Salesman Position","")
                _cbg_text="" if pd.isna(_cbg_value) else str(_cbg_value).strip()
                _position_text="" if pd.isna(_position_value) else str(_position_value).strip()
                _tgt_value=pd.to_numeric(_discipline_row.get("TGT RC",0),errors="coerce")
                _act_value=pd.to_numeric(_discipline_row.get("ACT",0),errors="coerce")
                _tgt_rc=0.0 if pd.isna(_tgt_value) else float(_tgt_value)
                _act_rc=0.0 if pd.isna(_act_value) else float(_act_value)
                _at_rc=_act_rc/_tgt_rc*100 if _tgt_rc else None
                _daily_lines.append(f"• {_cbg_text+' / ' if _cbg_text else ''}{_position_text} : {_tgt_rc:,.0f}/{_act_rc:,.0f}/{_at_rc:.0f}%" if _at_rc is not None else f"• {_cbg_text+' / ' if _cbg_text else ''}{_position_text} : {_tgt_rc:,.0f}/{_act_rc:,.0f}/—")
            _daily_lines.extend(["", "2. Program PUSH : Nama program / TGT / ACT / A/T"])
            for _push_row in push_daily_rows:
                _metric=_push_row["Metric"]
                _target_text=_daily_money(_push_row["Target"]) if _metric=="Total Sales" else f"{_push_row['Target']:,.0f}"
                _actual_text=_daily_money(_push_row["Actual"]) if _metric=="Total Sales" else f"{_push_row['Actual']:,.0f}"
                _at_text=f"{_push_row['A/T']:.0f}%" if _push_row["A/T"] is not None else "—"
                _daily_lines.append(f"• {_push_row['Nama program']} / {_target_text} / {_actual_text} / {_at_text}")
            _daily_lines.extend(["", "3. Sales KUM Gross : TGT / Act / A/T", f"• {_daily_money(sales_kum_target)} / {_daily_money(sales_kum_actual)} / {f'{sales_kum_achievement:.0f}%' if sales_kum_achievement is not None else '—'}", "", "4. DOT : Act / Tgt / A/T / Jml DOT", f"• {_daily_money(dot_actual)} / {_daily_money(dot_target)} / {f'{dot_achievement:.0f}%' if dot_achievement is not None else '—'} / {dot_count}x", "", "Isu dan insight pencapaian tidak on target : (RC, PUSH dan Sales)"])
            _issue_items=[line.strip().lstrip("•- ") for line in issues_text.splitlines() if line.strip()]
            _daily_lines.extend([f"{index}. {issue}" for index,issue in enumerate(_issue_items,1)] if _issue_items else ["(Belum ada isu/insight yang diisi.)"])
            _daily_lines.extend(["", "Demikian, Terima kasih."])
            _daily_report_result="\n".join(_daily_lines)
            st.session_state[f"generated_daily_report_{report_date_key}"]=_daily_report_result
            st.session_state[f"daily_report_preview_{report_date_key}"]=_daily_report_result
            _daily_report_settings.update({"business_group":daily_business_group,"dot_target":dot_target,"dot_actual":dot_actual,"dot_count":dot_count,"issues":issues_text})
            save_settings(settings)
        _daily_report_result=st.session_state.get(f"generated_daily_report_{report_date_key}")
        if _daily_report_result:
            _report_preview_key=f"daily_report_preview_{report_date_key}"
            if _report_preview_key not in st.session_state:
                st.session_state[_report_preview_key]=_daily_report_result
            st.text_area("Preview Report Harian",height=560,key=_report_preview_key)
            st.download_button("⬇ Unduh Report Harian (.txt)",data=_daily_report_result.encode("utf-8"),file_name=f"report_harian_BDG_{report_date:%Y%m%d}.txt",mime="text/plain",key=f"download_daily_report_{report_date_key}")
        save_settings(settings)
        st.stop()
    elif report_submenu=="Report Bulanan":
        st.markdown("#### 📅 Report Bulanan · Closing CHP & SPC")
        st.caption("Pilih bisnis, bulan, dan tahun. Actual/target bulanan serta POT dihitung dari database; lengkapi faktor pendukung/penghambat sebelum generate.")
        _monthly_bg_col=col(df,"business_group")
        _monthly_date_col=col(df,"date")
        _monthly_sales_col=col(df,"sales")
        _monthly_target_col=col(df,"target_value")
        _monthly_principal_col=col(df,"principal")
        if not all([_monthly_bg_col,_monthly_date_col,_monthly_sales_col,_monthly_target_col]):
            _monthly_missing=[label for label,value in [("Business Group",_monthly_bg_col),("Tanggal",_monthly_date_col),("Sales Value Gross",_monthly_sales_col),("Target Value",_monthly_target_col)] if not value]
            st.error("Kolom database belum lengkap untuk Report Bulanan: "+", ".join(_monthly_missing))
            st.stop()

        _monthly_dates=pd.to_datetime(df[_monthly_date_col],errors="coerce")
        _monthly_years=sorted(_monthly_dates.dropna().dt.year.astype(int).unique().tolist())
        _monthly_default_year=pd.Timestamp.now().year if pd.Timestamp.now().year in _monthly_years else (_monthly_years[-1] if _monthly_years else pd.Timestamp.now().year)
        _monthly_filter_cols=st.columns([2,1,1])
        with _monthly_filter_cols[0]:
            _monthly_bg_options=sorted(df[_monthly_bg_col].dropna().astype(str).str.strip().unique().tolist())
            monthly_business_groups=st.multiselect("Bisnis Group",_monthly_bg_options,default=[],key="monthly_report_business_groups",placeholder="Pilih bisnis; kosong = semua")
        with _monthly_filter_cols[1]:
            monthly_year=st.selectbox("Tahun",_monthly_years or [_monthly_default_year],index=(_monthly_years.index(_monthly_default_year) if _monthly_default_year in _monthly_years else 0),key="monthly_report_year")
        with _monthly_filter_cols[2]:
            _monthly_available_months=sorted(_monthly_dates.loc[_monthly_dates.dt.year.eq(int(monthly_year))].dropna().dt.month.astype(int).unique().tolist())
            _monthly_month_options=_monthly_available_months or list(MONTHS)
            _monthly_default_month=pd.Timestamp.now().month if pd.Timestamp.now().month in _monthly_month_options else _monthly_month_options[-1]
            monthly_month=st.selectbox("Bulan",_monthly_month_options,index=_monthly_month_options.index(_monthly_default_month),format_func=lambda number:MONTHS[int(number)],key="monthly_report_month")

        _monthly_mask=_monthly_dates.dt.year.eq(int(monthly_year))&_monthly_dates.dt.month.eq(int(monthly_month))
        if monthly_business_groups:
            _monthly_mask &= df[_monthly_bg_col].astype(str).str.strip().isin(monthly_business_groups)
        _monthly_scope=df.loc[_monthly_mask].copy()
        _monthly_scope["_monthly_actual"]=pd.to_numeric(_monthly_scope[_monthly_sales_col],errors="coerce").fillna(0)
        _monthly_scope["_monthly_target"]=pd.to_numeric(_monthly_scope[_monthly_target_col],errors="coerce").fillna(0)
        monthly_actual=float(_monthly_scope["_monthly_actual"].sum())
        monthly_target=float(_monthly_scope["_monthly_target"].sum())
        monthly_achievement=monthly_actual/monthly_target*100 if monthly_target else None
        monthly_deviation=monthly_actual-monthly_target
        _monthly_business_label=" & ".join(monthly_business_groups) if monthly_business_groups else "Semua Bisnis Group"
        _monthly_report_key=f"{int(monthly_year)}_{int(monthly_month):02d}_"+("_".join(key(v).replace(" ","_") for v in monthly_business_groups) if monthly_business_groups else "all")

        st.markdown(f"##### Bisnis {_monthly_business_label} · {MONTHS[int(monthly_month)]} {int(monthly_year)}")
        _monthly_metric_cols=st.columns(4)
        _monthly_metric_cols[0].metric("Actual Gross Sales",f"{monthly_actual:,.0f}")
        _monthly_metric_cols[1].metric("Target Gross Sales",f"{monthly_target:,.0f}")
        _monthly_metric_cols[2].metric("A/T",f"{monthly_achievement:.0f}%" if monthly_achievement is not None else "—")
        _monthly_metric_cols[3].metric("Deviasi",f"{monthly_deviation:,.0f}")

        st.markdown("##### Program PUSH")
        _monthly_push_rows=[]
        _monthly_name=MONTHS[int(monthly_month)]
        for _monthly_program in PROGRAMS:
            _monthly_selected={"year":[],"month":[],"principal_name":[],"principal":[],"channel_group":[],"business_group":[],"product_code":[],"product":[],"metric":"Total Relasi Beli Transaksi"}
            _monthly_selected=apply_active_program_config(_monthly_program,_monthly_selected,config_month=_monthly_name)
            _monthly_selected["year"]=[str(int(monthly_year))]
            _monthly_selected["month"]=[_monthly_name]
            _monthly_selected["business_group"]=monthly_business_groups
            _monthly_scoped,_monthly_actual,_monthly_program_target,_monthly_achievement,_monthly_dev,_=calculate_program_kpi(_monthly_program,_monthly_selected)
            _monthly_push_rows.append({"Nama Program":program_label(_monthly_program,month=_monthly_name),"Metric":_monthly_selected["metric"],"TGT":_monthly_program_target or 0,"ACT":_monthly_actual or 0,"A/T":_monthly_achievement,"Dev":_monthly_dev or 0})
        _monthly_push_table=pd.DataFrame([{key:value for key,value in row.items() if key!="Metric"} for row in _monthly_push_rows])
        st.dataframe(_monthly_push_table,width="stretch",hide_index=True,column_config={"TGT":st.column_config.NumberColumn(format="%,.0f"),"ACT":st.column_config.NumberColumn(format="%,.0f"),"A/T":st.column_config.NumberColumn(format="%.1f%%"),"Dev":st.column_config.NumberColumn(format="%,.0f")})

        st.markdown("##### POT · Principal On Target")
        if _monthly_principal_col:
            _monthly_pot_table=_monthly_scope.groupby(_monthly_principal_col,dropna=False).agg(ACT=("_monthly_actual","sum"),TGT=("_monthly_target","sum")).reset_index().rename(columns={_monthly_principal_col:"POT"})
            _monthly_pot_table=_monthly_pot_table[_monthly_pot_table["POT"].notna()&~_monthly_pot_table["POT"].astype(str).str.strip().isin(["","nan","<NA>"])].copy()
        else:
            _monthly_pot_table=pd.DataFrame(columns=["POT","ACT","TGT","A/T","Dev","Kontribusi Gross (%)"])
        if not _monthly_pot_table.empty:
            _monthly_pot_table["A/T"]=_monthly_pot_table["ACT"].div(_monthly_pot_table["TGT"].replace(0,float("nan")))*100
            _monthly_pot_table["Dev"]=_monthly_pot_table["ACT"]-_monthly_pot_table["TGT"]
            _monthly_pot_table["Kontribusi Gross (%)"]=_monthly_pot_table["ACT"].div(monthly_actual if monthly_actual else float("nan"))*100
            _monthly_pot_table=_monthly_pot_table.sort_values("ACT",ascending=False).reset_index(drop=True)
        else:
            _monthly_pot_table=pd.DataFrame(columns=["POT","ACT","TGT","A/T","Dev","Kontribusi Gross (%)"])
        _monthly_on_target=_monthly_pot_table[_monthly_pot_table["A/T"].ge(100)].copy() if not _monthly_pot_table.empty else _monthly_pot_table.copy()
        _monthly_below_target=_monthly_pot_table[_monthly_pot_table["A/T"].lt(100)|_monthly_pot_table["A/T"].isna()].copy() if not _monthly_pot_table.empty else _monthly_pot_table.copy()

        # Compare customer sales for each POT with the preceding calendar month.
        _monthly_customer_col=col(df,"customer")
        _monthly_previous_period=pd.Period(year=int(monthly_year),month=int(monthly_month),freq="M")-1
        _monthly_relation_changes=pd.DataFrame(columns=["POT","Relasi","Bulan Sebelumnya","Bulan Dipilih","Perubahan","Perubahan (%)"])
        _monthly_relation_summary={}
        if _monthly_principal_col and _monthly_customer_col:
            _relation_sales=pd.DataFrame({
                "POT":df[_monthly_principal_col].astype("string").str.strip(),
                "Relasi":df[_monthly_customer_col].astype("string").str.strip(),
                "Sales":pd.to_numeric(df[_monthly_sales_col],errors="coerce").fillna(0),
                "Period":_monthly_dates.dt.to_period("M"),
                "Bisnis":df[_monthly_bg_col].astype("string").str.strip(),
            })
            _relation_sales=_relation_sales[_relation_sales["POT"].notna()&_relation_sales["Relasi"].notna()&_relation_sales["POT"].ne("")&_relation_sales["Relasi"].ne("")]
            if monthly_business_groups:
                _relation_sales=_relation_sales[_relation_sales["Bisnis"].isin(monthly_business_groups)]
            _relation_current=_relation_sales[_relation_sales["Period"].eq(pd.Period(year=int(monthly_year),month=int(monthly_month),freq="M"))].groupby(["POT","Relasi"],as_index=False)["Sales"].sum().rename(columns={"Sales":"Bulan Dipilih"})
            _relation_previous=_relation_sales[_relation_sales["Period"].eq(_monthly_previous_period)].groupby(["POT","Relasi"],as_index=False)["Sales"].sum().rename(columns={"Sales":"Bulan Sebelumnya"})
            _monthly_relation_changes=_relation_previous.merge(_relation_current,on=["POT","Relasi"],how="outer").fillna({"Bulan Sebelumnya":0,"Bulan Dipilih":0})
            _monthly_relation_changes["Perubahan"]=_monthly_relation_changes["Bulan Dipilih"]-_monthly_relation_changes["Bulan Sebelumnya"]
            _monthly_relation_changes["Perubahan (%)"]=_monthly_relation_changes["Perubahan"].div(_monthly_relation_changes["Bulan Sebelumnya"].replace(0,float("nan")))*100
            _monthly_relation_changes=_monthly_relation_changes[_monthly_relation_changes["Bulan Sebelumnya"].ne(0)|_monthly_relation_changes["Bulan Dipilih"].ne(0)].copy()
            for _pot_name,_pot_relations in _monthly_relation_changes.groupby("POT"):
                _rel_increase=_pot_relations[_pot_relations["Perubahan"].gt(0)].sort_values("Perubahan",ascending=False)
                _rel_decline=_pot_relations[_pot_relations["Perubahan"].lt(0)].sort_values("Perubahan",ascending=True)
                _rel_new=_pot_relations[_pot_relations["Bulan Sebelumnya"].eq(0)&_pot_relations["Bulan Dipilih"].gt(0)]
                _rel_lost=_pot_relations[_pot_relations["Bulan Sebelumnya"].gt(0)&_pot_relations["Bulan Dipilih"].eq(0)]
                _pot_current_sales=float(_pot_relations["Bulan Dipilih"].sum())
                _pot_previous_sales=float(_pot_relations["Bulan Sebelumnya"].sum())
                _pot_change=_pot_current_sales-_pot_previous_sales
                _pot_change_pct=_pot_change/_pot_previous_sales*100 if _pot_previous_sales else None
                _best_gain=_rel_increase.iloc[0] if not _rel_increase.empty else None
                _largest_drop=_rel_decline.iloc[0] if not _rel_decline.empty else None
                _monthly_relation_summary[str(_pot_name)]={
                    "current":_pot_current_sales,"previous":_pot_previous_sales,"change":_pot_change,"change_pct":_pot_change_pct,
                    "up_count":len(_rel_increase),"down_count":len(_rel_decline),"new_count":len(_rel_new),"lost_count":len(_rel_lost),
                    "current_relation_count":int(_pot_relations["Bulan Dipilih"].gt(0).sum()),"previous_relation_count":int(_pot_relations["Bulan Sebelumnya"].gt(0).sum()),
                    "best_gain":_best_gain,"largest_drop":_largest_drop,
                }

        _previous_month_label=f"{MONTHS[_monthly_previous_period.month]} {_monthly_previous_period.year}"
        st.caption(f"Analisa relasi membandingkan {_previous_month_label} dengan {MONTHS[int(monthly_month)]} {int(monthly_year)}.")
        st.caption("Analisa tren Principal membandingkan sales bulan terpilih dengan rata-rata bulan-bulan sebelumnya pada tahun pilihan serta bulan kalender sebelumnya.")
        if not _monthly_relation_changes.empty:
            _relation_risers=_monthly_relation_changes[_monthly_relation_changes["Perubahan"].gt(0)].sort_values("Perubahan",ascending=False).groupby("POT",sort=False).head(5)
            _relation_fallers=_monthly_relation_changes[_monthly_relation_changes["Perubahan"].lt(0)].sort_values("Perubahan",ascending=True).groupby("POT",sort=False).head(5)
            _relation_display=pd.concat([_relation_risers,_relation_fallers]).sort_values(["POT","Perubahan"],ascending=[True,False]).copy()
            st.markdown("##### Perubahan Sales per Relasi · 5 kenaikan dan 5 penurunan terbesar per POT")
            st.dataframe(_relation_display,width="stretch",hide_index=True,column_config={"Bulan Sebelumnya":st.column_config.NumberColumn(format="%,.0f"),"Bulan Dipilih":st.column_config.NumberColumn(format="%,.0f"),"Perubahan":st.column_config.NumberColumn(format="%,.0f"),"Perubahan (%)":st.column_config.NumberColumn(format="%.1f%%")})
        elif _monthly_customer_col:
            st.info("Belum ada nilai transaksi relasi pada bulan ini atau bulan sebelumnya untuk dibandingkan.")
        else:
            st.info("Kolom relasi/customer tidak ditemukan; analisa per relasi tidak dapat dihitung.")

        def _monthly_relation_analysis(pot_name):
            _pot_trend=_monthly_relation_summary.get(str(pot_name))
            if _pot_trend is None:
                return f"Belum ada transaksi relasi yang dapat dibandingkan dengan {_previous_month_label}."
            _change_pct="—" if _pot_trend["change_pct"] is None else f"{_pot_trend['change_pct']:+.1f}%"
            _trend_parts=[f"sales relasi {_previous_month_label} {_pot_trend['previous']:,.0f} menjadi {_pot_trend['current']:,.0f} ({_pot_trend['change']:+,.0f}; {_change_pct})",f"jumlah relasi aktif {_pot_trend['previous_relation_count']} menjadi {_pot_trend['current_relation_count']}",f"{_pot_trend['up_count']} relasi naik",f"{_pot_trend['down_count']} relasi turun",f"{_pot_trend['new_count']} relasi baru",f"{_pot_trend['lost_count']} relasi tidak repeat"]
            if _pot_trend["best_gain"] is not None:
                _g=_pot_trend["best_gain"]
                _trend_parts.append(f"kenaikan terbesar {_g['Relasi']} ({_g['Perubahan']:+,.0f}; sales bulan ini {_g['Bulan Dipilih']:,.0f})")
            if _pot_trend["largest_drop"] is not None:
                _d=_pot_trend["largest_drop"]
                _trend_parts.append(f"penurunan terbesar {_d['Relasi']} ({_d['Bulan Sebelumnya']:,.0f} menjadi {_d['Bulan Dipilih']:,.0f}; perubahan {_d['Perubahan']:+,.0f})")
            return "; ".join(_trend_parts)+"."

        # Trend total sales per Principal: selected month vs preceding calendar
        # month and average of earlier months in the selected year.
        _monthly_principal_summary={}
        if _monthly_principal_col:
            _principal_trend_data=pd.DataFrame({
                "POT":df[_monthly_principal_col].astype("string").str.strip(),
                "Period":_monthly_dates.dt.to_period("M"),
                "Sales":pd.to_numeric(df[_monthly_sales_col],errors="coerce").fillna(0),
                "Bisnis":df[_monthly_bg_col].astype("string").str.strip(),
            })
            _principal_trend_data=_principal_trend_data[_principal_trend_data["POT"].notna()&_principal_trend_data["POT"].ne("")]
            if monthly_business_groups:
                _principal_trend_data=_principal_trend_data[_principal_trend_data["Bisnis"].isin(monthly_business_groups)]
            _principal_trend_data=_principal_trend_data[_principal_trend_data["Period"].notna()&_principal_trend_data["Period"].le(pd.Period(year=int(monthly_year),month=int(monthly_month),freq="M"))]
            _principal_monthly_sales=_principal_trend_data.groupby(["POT","Period"],as_index=False)["Sales"].sum()
            _current_principal_period=pd.Period(year=int(monthly_year),month=int(monthly_month),freq="M")
            for _principal_name in _principal_monthly_sales["POT"].dropna().unique():
                _principal_rows=_principal_monthly_sales[_principal_monthly_sales["POT"].eq(_principal_name)]
                _current_match=_principal_rows[_principal_rows["Period"].eq(_current_principal_period)]
                _previous_match=_principal_rows[_principal_rows["Period"].eq(_monthly_previous_period)]
                _year_history=_principal_rows[(_principal_rows["Period"].dt.year==int(monthly_year))&(_principal_rows["Period"].dt.month<int(monthly_month))]
                _current_sales=float(_current_match["Sales"].sum())
                _previous_sales=float(_previous_match["Sales"].sum()) if not _previous_match.empty else None
                _months_elapsed=max(0,int(monthly_month)-1)
                _average_sales=float(_year_history["Sales"].sum())/_months_elapsed if _months_elapsed else None
                _monthly_principal_summary[str(_principal_name)]={"current":_current_sales,"previous":_previous_sales,"average":_average_sales,"history_months":_months_elapsed}

        def _monthly_principal_analysis(pot_name):
            _trend=_monthly_principal_summary.get(str(pot_name))
            if _trend is None:
                return "Kolom Principal ID tidak tersedia untuk analisa tren."
            _parts=[f"sales Principal bulan ini {_trend['current']:,.0f}"]
            if _trend["average"] is not None:
                _delta=_trend["current"]-_trend["average"]
                _delta_pct=_delta/_trend["average"]*100 if _trend["average"] else None
                _direction="naik" if _delta>0 else "turun" if _delta<0 else "stabil"
                _period_span=f"Jan–{MONTHS[int(monthly_month)-1]} {int(monthly_year)}" if int(monthly_month)>2 else f"bulan sebelumnya di {int(monthly_year)}"
                _pct_text=f" ({_delta_pct:+.1f}%)" if _delta_pct is not None else ""
                _parts.append(f"dibanding rata-rata {_period_span} {_trend['average']:,.0f}, {_direction} {abs(_delta):,.0f}{_pct_text}")
            else:
                _parts.append(f"rata-rata {int(monthly_year)} belum tersedia karena belum ada bulan pembanding sebelumnya")
            if _trend["previous"] is not None:
                _prev_delta=_trend["current"]-_trend["previous"]
                _prev_direction="naik" if _prev_delta>0 else "turun" if _prev_delta<0 else "stabil"
                _parts.append(f"dibanding {_previous_month_label} {_trend['previous']:,.0f}, {_prev_direction} {abs(_prev_delta):,.0f}")
            else:
                _parts.append(f"data {_previous_month_label} tidak tersedia")
            return "; ".join(_parts)+"."

        if not _monthly_pot_table.empty:
            st.dataframe(_monthly_pot_table,width="stretch",hide_index=True,column_config={"ACT":st.column_config.NumberColumn(format="%,.0f"),"TGT":st.column_config.NumberColumn(format="%,.0f"),"A/T":st.column_config.NumberColumn(format="%.1f%%"),"Dev":st.column_config.NumberColumn(format="%,.0f"),"Kontribusi Gross (%)":st.column_config.NumberColumn(format="%.1f%%")})
        else:
            st.info("Tidak ada Principal ID/POT dalam data bulan dan bisnis yang dipilih.")

        _support_input=_monthly_on_target[["POT","A/T","Kontribusi Gross (%)"]].copy() if not _monthly_on_target.empty else pd.DataFrame(columns=["POT","A/T","Kontribusi Gross (%)"])
        _support_input["Faktor pendukung / analisa"]= [""]*len(_support_input)
        _support_input["Analisa tren Principal"]=_support_input["POT"].map(_monthly_principal_analysis) if not _support_input.empty else pd.Series(dtype="string")
        _support_input["Analisa perubahan relasi"]=_support_input["POT"].map(_monthly_relation_analysis) if not _support_input.empty else pd.Series(dtype="string")
        _barrier_input=_monthly_below_target[["POT","ACT","TGT","A/T","Dev"]].copy() if not _monthly_below_target.empty else pd.DataFrame(columns=["POT","ACT","TGT","A/T","Dev"])
        _barrier_input["Faktor penghambat / analisa"]= [""]*len(_barrier_input)
        _barrier_input["Analisa tren Principal"]=_barrier_input["POT"].map(_monthly_principal_analysis) if not _barrier_input.empty else pd.Series(dtype="string")
        _barrier_input["Analisa perubahan relasi"]=_barrier_input["POT"].map(_monthly_relation_analysis) if not _barrier_input.empty else pd.Series(dtype="string")
        _support_editor_signature=hashlib.sha256(_support_input.drop(columns=["Faktor pendukung / analisa"]).to_csv(index=False).encode("utf-8")).hexdigest()[:10]
        _barrier_editor_signature=hashlib.sha256(_barrier_input.drop(columns=["Faktor penghambat / analisa"]).to_csv(index=False).encode("utf-8")).hexdigest()[:10]
        st.markdown("##### Faktor Pendukung · POT on target")
        if not _support_input.empty:
            _support_input=st.data_editor(_support_input,width="stretch",hide_index=True,key=f"monthly_support_editor_{_monthly_report_key}_{_support_editor_signature}",column_config={"POT":st.column_config.TextColumn(disabled=True),"A/T":st.column_config.NumberColumn(disabled=True,format="%.1f%%"),"Kontribusi Gross (%)":st.column_config.NumberColumn(disabled=True,format="%.1f%%"),"Faktor pendukung / analisa":st.column_config.TextColumn(help="Tuliskan produk/relasi/program/dukungan yang mendorong pencapaian."),"Analisa tren Principal":st.column_config.TextColumn(disabled=True,width="large"),"Analisa perubahan relasi":st.column_config.TextColumn(disabled=True,width="large")})
        else:
            st.info("Belum ada POT dengan pencapaian minimal 100%.")
        st.markdown("##### Faktor Penghambat · POT belum on target")
        if not _barrier_input.empty:
            _barrier_input=st.data_editor(_barrier_input,width="stretch",hide_index=True,key=f"monthly_barrier_editor_{_monthly_report_key}_{_barrier_editor_signature}",column_config={"POT":st.column_config.TextColumn(disabled=True),"ACT":st.column_config.NumberColumn(disabled=True,format="%,.0f"),"TGT":st.column_config.NumberColumn(disabled=True,format="%,.0f"),"A/T":st.column_config.NumberColumn(disabled=True,format="%.1f%%"),"Dev":st.column_config.NumberColumn(disabled=True,format="%,.0f"),"Faktor penghambat / analisa":st.column_config.TextColumn(help="Tuliskan kendala stok, pemenuhan PO, atau penyebab lain."),"Analisa tren Principal":st.column_config.TextColumn(disabled=True,width="large"),"Analisa perubahan relasi":st.column_config.TextColumn(disabled=True,width="large")})
        else:
            st.info("Semua POT sudah on target.")

        if st.button("✨ Generate Report Bulanan",type="primary",key=f"generate_monthly_report_{_monthly_report_key}"):
            _month_title=MONTHS[int(monthly_month)]
            _report_business_title=" & ".join(monthly_business_groups) if monthly_business_groups else "Semua Bisnis Group"
            _monthly_lines=[
                "Selamat Pagi Pak Bogi dan Pak Jefta,",
                f"Berikut Report Closing Bulan {_month_title} {int(monthly_year)} CHP & SPC Cabang {ACTIVE_BRANCH_NAME}.",
                "",
                f"BISNIS {_report_business_title.upper()}",
                f"*Bisnis {_report_business_title} GROSS SALES*",
                f"Actual :   {monthly_actual:,.0f}",
                f"Target :   {monthly_target:,.0f}",
                f"A/T :  {monthly_achievement:.0f}%" if monthly_achievement is not None else "A/T :  —",
                f"Dev : {monthly_deviation:,.0f}",
                "",
                "PROGRAM PUSH :",
                "Nama Program /TGT/ACT/A/t/Dev",
            ]
            for _monthly_row in _monthly_push_rows:
                _monthly_metric=_monthly_row["Metric"]
                if _monthly_metric=="Total Sales":
                    _program_target_text=f"{_monthly_row['TGT']/1_000_000:,.1f}Jt" if abs(_monthly_row["TGT"])>=1_000_000 else f"{_monthly_row['TGT']:,.0f}"
                    _program_actual_text=f"{_monthly_row['ACT']/1_000_000:,.1f}Jt" if abs(_monthly_row["ACT"])>=1_000_000 else f"{_monthly_row['ACT']:,.0f}"
                    _program_dev_text=f"{_monthly_row['Dev']/1_000_000:,.1f}Jt" if abs(_monthly_row["Dev"])>=1_000_000 else f"{_monthly_row['Dev']:,.0f}"
                else:
                    _program_target_text=f"{_monthly_row['TGT']:,.0f}"; _program_actual_text=f"{_monthly_row['ACT']:,.0f}"; _program_dev_text=f"{_monthly_row['Dev']:,.0f}"
                _program_at_text=f"{_monthly_row['A/T']:.0f}%" if _monthly_row["A/T"] is not None else "—"
                _monthly_lines.append(f"• {_monthly_row['Nama Program']} / {_program_target_text} / {_program_actual_text} / {_program_at_text} / {_program_dev_text}")
            _pot_names=_monthly_on_target["POT"].astype(str).tolist() if not _monthly_on_target.empty else []
            _monthly_lines.extend(["", f"POT :  {len(_pot_names)} POT ({', '.join(_pot_names)})", "", "*FAKTOR PENDUKUNG*"])
            if not _monthly_on_target.empty:
                for _,_support_row in _support_input.iterrows():
                    _pot_name=str(_support_row["POT"])
                    _pot_at=float(_support_row["A/T"]) if pd.notna(_support_row["A/T"]) else 0
                    _pot_contrib=float(_support_row["Kontribusi Gross (%)"]) if pd.notna(_support_row["Kontribusi Gross (%)"]) else 0
                    _pot_note_value=_support_row.get("Faktor pendukung / analisa","")
                    _pot_note="" if pd.isna(_pot_note_value) else str(_pot_note_value).strip()
                    _auto_analysis=f"Pencapaian {_pot_at:.0f}%, kontribusi gross {_pot_contrib:.1f}% dari total bisnis. "
                    _monthly_lines.append(f"- {_pot_name} ({_pot_at:.0f}%, kontribusi {_pot_contrib:.1f}%). "+(f"{_pot_note} " if _pot_note else _auto_analysis)+f"Tren Principal: {_monthly_principal_analysis(_pot_name)} Analisa relasi: {_monthly_relation_analysis(_pot_name)}")
            else:
                _monthly_lines.append("- Tidak ada POT yang mencapai target bulan ini.")
            _monthly_lines.extend(["", "*FAKTOR PENGHAMBAT :*"])
            if not _monthly_below_target.empty:
                for _,_barrier_row in _barrier_input.iterrows():
                    _pot_name=str(_barrier_row["POT"])
                    _pot_at=float(_barrier_row["A/T"]) if pd.notna(_barrier_row["A/T"]) else 0
                    _pot_dev=float(_barrier_row["Dev"]) if pd.notna(_barrier_row["Dev"]) else 0
                    _pot_note_value=_barrier_row.get("Faktor penghambat / analisa","")
                    _pot_note="" if pd.isna(_pot_note_value) else str(_pot_note_value).strip()
                    _auto_analysis=f"Pencapaian {_pot_at:.0f}% dengan deviasi {_pot_dev:,.0f}; actual masih di bawah target bulan ini. "
                    _monthly_lines.append(f"- {_pot_name} ({_pot_at:.0f}%, deviasi {_pot_dev:,.0f}). "+(f"{_pot_note} " if _pot_note else _auto_analysis)+f"Tren Principal: {_monthly_principal_analysis(_pot_name)} Analisa relasi: {_monthly_relation_analysis(_pot_name)}")
            else:
                _monthly_lines.append("- Tidak ada POT di bawah target bulan ini.")
            _monthly_lines.extend(["", "Terimakasih"])
            _monthly_result="\n".join(_monthly_lines)
            st.session_state[f"generated_monthly_report_{_monthly_report_key}"]=_monthly_result
            st.session_state[f"monthly_report_preview_{_monthly_report_key}"]=_monthly_result
        _monthly_result=st.session_state.get(f"generated_monthly_report_{_monthly_report_key}")
        if _monthly_result:
            _monthly_preview_key=f"monthly_report_preview_{_monthly_report_key}"
            if _monthly_preview_key not in st.session_state: st.session_state[_monthly_preview_key]=_monthly_result
            st.text_area("Preview Report Bulanan",height=600,key=_monthly_preview_key)
            st.download_button("⬇ Unduh Report Bulanan (.txt)",data=_monthly_result.encode("utf-8"),file_name=f"report_closing_{key(_monthly_business_label).replace(' ','_')}_{int(monthly_year)}_{int(monthly_month):02d}.txt",mime="text/plain",key=f"download_monthly_report_{_monthly_report_key}")
        save_settings(settings)
        st.stop()
    elif report_submenu=="Slide Meeting Kasie":
        st.markdown("#### 📊 Slide Meeting Kasie · Review Bulanan Cabang")
        st.caption("Pilih Bisnis Group, bulan, dan tahun. Isi slide akan dihitung dari data aktual dashboard dan mengikuti struktur template Monthly Report Cabang.")
        _kasie_bg_col=col(df,"business_group")
        _kasie_date_col=col(df,"date")
        _kasie_sales_col=col(df,"sales")
        _kasie_target_col=col(df,"target_value")
        _kasie_principal_col=col(df,"principal")
        _kasie_customer_col=col(df,"customer")
        _kasie_missing=[label for label,column in [("Bisnis Group",_kasie_bg_col),("Tanggal",_kasie_date_col),("Sales Value Gross",_kasie_sales_col),("Target Value",_kasie_target_col),("Principal ID",_kasie_principal_col),("Customer Name",_kasie_customer_col)] if not column]
        if _kasie_missing or not _PPTX_OK or generate_meeting_kasie_pptx is None:
            if _kasie_missing: st.error("Kolom database belum lengkap untuk Slide Meeting Kasie: "+", ".join(_kasie_missing))
            else: st.error("Fitur export PPT belum tersedia. Pastikan dependency python-pptx tersedia.")
            st.stop()

        _kasie_dates=pd.to_datetime(df[_kasie_date_col],errors="coerce")
        _kasie_years=sorted(_kasie_dates.dropna().dt.year.astype(int).unique().tolist())
        _kasie_default_year=pd.Timestamp.now().year if pd.Timestamp.now().year in _kasie_years else (_kasie_years[-1] if _kasie_years else pd.Timestamp.now().year)
        _kasie_filters=st.columns([2,1,1])
        with _kasie_filters[0]:
            _kasie_bg_options=sorted(df[_kasie_bg_col].dropna().astype(str).str.strip().unique().tolist())
            _kasie_groups=st.multiselect("Bisnis Group",_kasie_bg_options,key="kasie_report_business_groups",placeholder="Kosongkan untuk semua Bisnis Group")
        with _kasie_filters[1]:
            _kasie_year=st.selectbox("Tahun",_kasie_years or [_kasie_default_year],index=(_kasie_years.index(_kasie_default_year) if _kasie_default_year in _kasie_years else 0),key="kasie_report_year")
        with _kasie_filters[2]:
            _kasie_month_options=sorted(_kasie_dates.loc[_kasie_dates.dt.year.eq(int(_kasie_year))].dropna().dt.month.astype(int).unique().tolist()) or list(MONTHS)
            _kasie_default_month=pd.Timestamp.now().month if pd.Timestamp.now().month in _kasie_month_options else _kasie_month_options[-1]
            _kasie_month=st.selectbox("Bulan report",_kasie_month_options,index=_kasie_month_options.index(_kasie_default_month),format_func=lambda value:MONTHS[int(value)],key="kasie_report_month")

        _kasie_data=df[[_kasie_bg_col,_kasie_date_col,_kasie_sales_col,_kasie_target_col,_kasie_principal_col,_kasie_customer_col]].copy()
        _kasie_data.columns=["Bisnis Group","Tanggal","Actual","Target","Principal ID","Customer"]
        _kasie_data["Tanggal"]=_kasie_dates
        _kasie_data["Bisnis Group"]=_kasie_data["Bisnis Group"].astype("string").str.strip()
        _kasie_data["Actual"]=pd.to_numeric(_kasie_data["Actual"],errors="coerce").fillna(0)
        _kasie_data["Target"]=pd.to_numeric(_kasie_data["Target"],errors="coerce").fillna(0)
        if _kasie_groups:
            _kasie_data=_kasie_data[_kasie_data["Bisnis Group"].isin(_kasie_groups)]
        _kasie_period=_kasie_data[_kasie_data["Tanggal"].dt.year.eq(int(_kasie_year))&_kasie_data["Tanggal"].dt.month.eq(int(_kasie_month))].copy()
        _kasie_previous_start=pd.Timestamp(int(_kasie_year),int(_kasie_month),1)-pd.offsets.MonthBegin(1)
        _kasie_previous=_kasie_data[_kasie_data["Tanggal"].dt.year.eq(_kasie_previous_start.year)&_kasie_data["Tanggal"].dt.month.eq(_kasie_previous_start.month)].copy()
        _kasie_ytd=_kasie_data[_kasie_data["Tanggal"].dt.year.eq(int(_kasie_year))&_kasie_data["Tanggal"].dt.month.le(int(_kasie_month))].copy()
        _kasie_actual=float(_kasie_period["Actual"].sum()); _kasie_target=float(_kasie_period["Target"].sum())
        _kasie_achievement=_kasie_actual/_kasie_target*100 if _kasie_target else None
        _kasie_deviation=_kasie_actual-_kasie_target
        _kasie_ytd_actual=float(_kasie_ytd["Actual"].sum()); _kasie_ytd_target=float(_kasie_ytd["Target"].sum())
        _kasie_ytd_achievement=_kasie_ytd_actual/_kasie_ytd_target*100 if _kasie_ytd_target else None
        _kasie_previous_actual=float(_kasie_previous["Actual"].sum())
        _kasie_mom_change=_kasie_actual-_kasie_previous_actual
        _kasie_achievement_text=lambda value:f"{value:.1f}%" if value is not None and pd.notna(value) else "—"
        _kasie_money=lambda value:f"Rp {float(value):,.0f}"
        _kasie_business_label=" & ".join(_kasie_groups) if _kasie_groups else "Semua Bisnis Group"
        _kasie_month_name=MONTHS[int(_kasie_month)]
        _kasie_report_key=f"{int(_kasie_year)}_{int(_kasie_month):02d}_"+("_".join(key(v).replace(" ","_") for v in _kasie_groups) if _kasie_groups else "all")
        _kasie_metric_cols=st.columns(4)
        _kasie_metric_cols[0].metric("Actual Gross",_kasie_money(_kasie_actual))
        _kasie_metric_cols[1].metric("Target Gross",_kasie_money(_kasie_target))
        _kasie_metric_cols[2].metric("A/T",_kasie_achievement_text(_kasie_achievement))
        _kasie_metric_cols[3].metric("Deviasi",_kasie_money(_kasie_deviation))

        _kasie_business=(_kasie_period.groupby("Bisnis Group",dropna=False).agg(Actual=("Actual","sum"),Target=("Target","sum")).reset_index())
        _kasie_business["A/T"]=_kasie_business["Actual"].div(_kasie_business["Target"].replace(0,float("nan")))*100
        _kasie_business["Dev"]=_kasie_business["Actual"]-_kasie_business["Target"]
        _kasie_business=_kasie_business.sort_values("A/T",ascending=False,na_position="last")
        _kasie_principal=(_kasie_period.groupby("Principal ID",dropna=False).agg(Actual=("Actual","sum"),Target=("Target","sum")).reset_index())
        _kasie_principal["Principal ID"]=_kasie_principal["Principal ID"].astype("string").fillna("(Tidak tercantum)").str.strip()
        _kasie_principal["A/T"]=_kasie_principal["Actual"].div(_kasie_principal["Target"].replace(0,float("nan")))*100
        _kasie_principal["Dev"]=_kasie_principal["Actual"]-_kasie_principal["Target"]
        _kasie_pot=_kasie_principal[_kasie_principal["A/T"].ge(100)].sort_values("A/T",ascending=False)
        _kasie_below=_kasie_principal[_kasie_principal["A/T"].lt(100)].sort_values("Dev")
        _kasie_pot_count=int(len(_kasie_pot))

        _kasie_customer_current=(_kasie_period.groupby("Customer",dropna=False)["Actual"].sum().rename("Actual").reset_index())
        _kasie_customer_previous=(_kasie_previous.groupby("Customer",dropna=False)["Actual"].sum().rename("Previous").reset_index())
        _kasie_customers=_kasie_customer_current.merge(_kasie_customer_previous,on="Customer",how="outer").fillna({"Actual":0,"Previous":0})
        _kasie_customers["Change"]=_kasie_customers["Actual"]-_kasie_customers["Previous"]
        _kasie_customers["Customer"]=_kasie_customers["Customer"].astype("string").fillna("(Tidak tercantum)").str.strip()
        _kasie_customers=_kasie_customers.sort_values("Change",ascending=False)
        _kasie_customers["Actual Text"]=_kasie_customers["Actual"].map(_kasie_money)
        _kasie_customers["Change Text"]=_kasie_customers["Change"].map(lambda value:("+" if value>0 else "")+_kasie_money(value))
        _kasie_customer_table=_kasie_customers.head(10).copy()

        _kasie_trend=[]; _kasie_trend_actual=[]; _kasie_trend_target=[]
        for _m in range(1,int(_kasie_month)+1):
            _m_scope=_kasie_data[_kasie_data["Tanggal"].dt.year.eq(int(_kasie_year))&_kasie_data["Tanggal"].dt.month.eq(_m)]
            _kasie_trend.append(MONTHS[_m][:3].upper())
            _kasie_trend_actual.append(float(_m_scope["Actual"].sum()))
            _kasie_trend_target.append(float(_m_scope["Target"].sum()))
        _kasie_2025=_kasie_data[_kasie_data["Tanggal"].dt.year.eq(2025)]
        _kasie_avg_2025=float(_kasie_2025.groupby(_kasie_2025["Tanggal"].dt.month)["Actual"].sum().mean()) if not _kasie_2025.empty else 0.0
        _kasie_contribution=_kasie_principal.sort_values("Actual",ascending=False).head(8)

        _kasie_program_rows=[]
        for _program in PROGRAMS:
            _program_selected={"year":[],"month":[],"principal_name":[],"principal":[],"channel_group":[],"business_group":[],"product_code":[],"product":[],"metric":"Total Relasi Beli Transaksi"}
            _program_selected=apply_active_program_config(_program,_program_selected,config_month=_kasie_month_name)
            _program_selected["year"]=[str(_kasie_year)]; _program_selected["month"]=[_kasie_month_name]
            if _kasie_groups:
                _program_bg=_program_selected.get("business_group",[])
                _program_bg=list(_program_bg) if isinstance(_program_bg,(list,tuple,set)) else ([] if _program_bg in (None,"All","ALL","") else [_program_bg])
                _program_matches=[value for value in _program_bg if str(value) in {str(x) for x in _kasie_groups}]
                _program_selected["business_group"]=_program_matches or (_kasie_groups if not _program_bg else ["__NO_MATCH__"])
            _,_p_actual,_p_target,_p_at,_p_dev,_=calculate_program_kpi(_program,_program_selected)
            _kasie_program_rows.append({"Program":program_label(_program,month=_kasie_month_name),"Metric":_program_selected["metric"],"Target":float(_p_target or 0),"Actual":float(_p_actual or 0),"A/T":_p_at,"Dev":_p_dev})
        _kasie_program_df=pd.DataFrame(_kasie_program_rows)
        _kasie_program_df["Target Text"]=_kasie_program_df.apply(lambda row:_kasie_money(row["Target"]) if row["Metric"]=="Total Sales" else f"{float(row['Target']):,.0f}",axis=1)
        _kasie_program_df["Actual Text"]=_kasie_program_df.apply(lambda row:_kasie_money(row["Actual"]) if row["Metric"]=="Total Sales" else f"{float(row['Actual']):,.0f}",axis=1)
        _kasie_program_df["AT Text"]=_kasie_program_df["A/T"].map(_kasie_achievement_text)

        _kasie_business_display=_kasie_business.head(14).copy()
        _kasie_business_display["Actual Text"]=_kasie_business_display["Actual"].map(_kasie_money)
        _kasie_business_display["Target Text"]=_kasie_business_display["Target"].map(_kasie_money)
        _kasie_business_display["AT Text"]=_kasie_business_display["A/T"].map(_kasie_achievement_text)
        _kasie_business_display["Dev Text"]=_kasie_business_display["Dev"].map(_kasie_money)
        _kasie_principal_display=_kasie_principal.sort_values("A/T",ascending=False,na_position="last").head(14).copy()
        _kasie_principal_display["Actual Text"]=_kasie_principal_display["Actual"].map(_kasie_money)
        _kasie_principal_display["Target Text"]=_kasie_principal_display["Target"].map(_kasie_money)
        _kasie_principal_display["AT Text"]=_kasie_principal_display["A/T"].map(_kasie_achievement_text)
        _kasie_principal_display["Dev Text"]=_kasie_principal_display["Dev"].map(_kasie_money)

        _best_business=_kasie_business.iloc[0] if not _kasie_business.empty else None
        _weak_business=_kasie_business.sort_values("Dev").iloc[0] if not _kasie_business.empty else None
        _kasie_top_customer=_kasie_customers.iloc[0] if not _kasie_customers.empty else None
        _kasie_down_customer=_kasie_customers.sort_values("Change").iloc[0] if not _kasie_customers.empty else None
        _kasie_sales_insights=[
            f"Actual {_kasie_money(_kasie_actual)} dari target {_kasie_money(_kasie_target)} ({_kasie_achievement_text(_kasie_achievement)}); deviasi {_kasie_money(_kasie_deviation)}.",
            (f"Bisnis Group tertinggi: {_best_business['Bisnis Group']} ({_kasie_achievement_text(_best_business['A/T'])})." if _best_business is not None else "Belum ada rincian bisnis group."),
            (f"Gap terbesar per Bisnis Group: {_weak_business['Bisnis Group']} ({_kasie_money(_weak_business['Dev'])})." if _weak_business is not None else "Belum ada gap bisnis group."),
            (f"Kenaikan customer terbesar: {_kasie_top_customer['Customer']} ({_kasie_money(_kasie_top_customer['Change'])} vs bulan lalu)." if _kasie_top_customer is not None and _kasie_top_customer['Change']>0 else "Belum terlihat kenaikan customer dibanding bulan lalu."),
            (f"Penurunan customer terbesar: {_kasie_down_customer['Customer']} ({_kasie_money(_kasie_down_customer['Change'])} vs bulan lalu)." if _kasie_down_customer is not None and _kasie_down_customer['Change']<0 else f"Actual bulan lalu: {_kasie_money(_kasie_previous_actual)}."),
        ]
        _pot_insight_lines=[f"{_kasie_pot_count} Principal ID mencapai target. Teratas: "+(", ".join(f"{row['Principal ID']} ({row['A/T']:.0f}%)" for _,row in _kasie_pot.head(3).iterrows()) if not _kasie_pot.empty else "belum ada POT.")]
        _below_insight_lines=[f"{row['Principal ID']}: deviasi {_kasie_money(row['Dev'])}, A/T {_kasie_achievement_text(row['A/T'])}." for _,row in _kasie_below.head(2).iterrows()] or ["Semua Principal ID dengan target sudah mencapai target."]
        _trend_insight_lines=[]
        if _kasie_top_customer is not None and _kasie_top_customer["Change"]>0: _trend_insight_lines.append(f"Naik: {_kasie_top_customer['Customer']} bertambah {_kasie_money(_kasie_top_customer['Change'])} dibanding bulan lalu.")
        if _kasie_down_customer is not None and _kasie_down_customer["Change"]<0: _kasie_trend_insight_lines=f"Turun: {_kasie_down_customer['Customer']} berkurang {_kasie_money(abs(_kasie_down_customer['Change']))} dibanding bulan lalu."; _trend_insight_lines.append(_kasie_trend_insight_lines)
        _trend_insight_lines.append(f"Total sales {_kasie_money(_kasie_mom_change)} dibanding bulan lalu ({_kasie_money(_kasie_previous_actual)})." if not _kasie_previous.empty else "Data pembanding bulan lalu belum tersedia.")
        _top_group_lines=[f"Bisnis Group tertinggi: {_best_business['Bisnis Group']} · A/T {_kasie_achievement_text(_best_business['A/T'])}" if _best_business is not None else "Belum ada pencapaian per Bisnis Group.",f"Deviasi bisnis terbesar: {_weak_business['Bisnis Group']} · {_kasie_money(_weak_business['Dev'])}" if _weak_business is not None else "Belum ada deviasi bisnis."]
        _principal_score_lines=[f"POT: {_kasie_pot_count} Principal ID mencapai target.",f"Prioritas: {', '.join(_kasie_below.head(3)['Principal ID'].astype(str).tolist()) if not _kasie_below.empty else 'jaga momentum seluruh principal on target.'}"]
        _meeting_summary_lines=[f"Periode: {_kasie_business_label} · {_kasie_month_name} {int(_kasie_year)}.",f"Actual {_kasie_money(_kasie_actual)} dari target {_kasie_money(_kasie_target)} ({_kasie_achievement_text(_kasie_achievement)}).",f"POT tercapai: {_kasie_pot_count} Principal ID.",f"Perubahan sales vs bulan lalu: {_kasie_money(_kasie_mom_change)}.",_kasie_sales_insights[3],_kasie_sales_insights[4]]

        st.markdown("##### Program PUSH dan analisa")
        st.dataframe(_kasie_program_df[["Program","Target Text","Actual Text","AT Text"]].rename(columns={"Target Text":"Target","Actual Text":"Actual","AT Text":"A/T"}),width="stretch",hide_index=True)
        st.markdown("##### Analisa bulanan")
        for _insight in _kasie_sales_insights: st.markdown(f"- {_insight}")
        _kasie_template=ASSET_DIR/"monthly_report_kasie_template_2026.pptx"
        if not _kasie_template.exists():
            st.error("Template Monthly Report Cabang belum tersedia di folder assets.")
        elif not _PPTX_OK or generate_meeting_kasie_pptx is None:
            st.error("Dependency generator PowerPoint tidak tersedia.")
        elif st.button("✨ Generate Slide Meeting Kasie",type="primary",key=f"generate_kasie_{_kasie_report_key}"):
            _kasie_payload={
                "month_name":_kasie_month_name,"year":int(_kasie_year),"business_label":_kasie_business_label,
                "actual":_kasie_actual,"target":_kasie_target,"achievement":_kasie_achievement,"deviation":_kasie_deviation,
                "ytd_actual":_kasie_ytd_actual,"ytd_target":_kasie_ytd_target,"ytd_achievement":_kasie_ytd_achievement,"ytd_deviation":_kasie_ytd_actual-_kasie_ytd_target,
                "pot_count":_kasie_pot_count,"sales_insight_lines":_kasie_sales_insights,"pot_insight_lines":_pot_insight_lines,
                "below_insight_lines":_below_insight_lines,"trend_insight_lines":_trend_insight_lines,
                "trend_month_labels":_kasie_trend,"trend_actual":_kasie_trend_actual,"trend_target":_kasie_trend_target,"average_2025":_kasie_avg_2025,
                "principal_contribution_names":_kasie_contribution["Principal ID"].astype(str).tolist(),"principal_contribution_values":_kasie_contribution["Actual"].astype(float).tolist(),
                "customer_table":_kasie_customer_table,"principal_table":_kasie_principal_display,"attainment_chart_names":_kasie_principal_display["Principal ID"].astype(str).tolist(),
                "attainment_chart_values":_kasie_principal_display["A/T"].fillna(0).astype(float).tolist(),"top_group_lines":_top_group_lines,
                "principal_score_lines":_principal_score_lines,"meeting_summary_lines":_meeting_summary_lines,
                "business_table":_kasie_business_display,"program_table":_kasie_program_df,
            }
            try:
                st.session_state[f"kasie_ppt_{_kasie_report_key}"]=generate_meeting_kasie_pptx(_kasie_template,_kasie_payload)
                st.success("Slide Meeting Kasie berhasil dibuat sesuai periode dan filter yang dipilih.")
            except Exception as _kasie_error:
                st.error(f"PPT tidak berhasil dibuat: {_kasie_error}")
        _kasie_ppt_data=st.session_state.get(f"kasie_ppt_{_kasie_report_key}")
        if _kasie_ppt_data:
            st.download_button("⬇ Export file PPT (.pptx)",data=_kasie_ppt_data,file_name=f"Slide_Meeting_Kasie_{key(_kasie_business_label).replace(' ','_')}_{int(_kasie_year)}_{int(_kasie_month):02d}.pptx",mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",key=f"download_kasie_{_kasie_report_key}")
        st.stop()

    st.caption("Upload template teks laporan, sistem akan mengisi variabel secara otomatis dari data aktual dashboard.")

    # ── Upload template ───────────────────────────────────────────────────
    report_col1, report_col2 = st.columns([2,3], gap="medium")

    with report_col1:
        st.markdown("#### 📄 Upload Template")
        st.caption(
            "Buat file .txt dengan variabel dalam format `{{NAMA_VARIABEL}}`. "
            "Contoh: `{{ACTUAL_TOTAL}}`, `{{BULAN}}`, `{{ACHIEVEMENT}}`."
        )
        template_file = st.file_uploader(
            "Upload template (.txt)",
            type=["txt"],
            key="report_template_upload",
        )
        st.markdown("---")
        st.markdown("**Variabel yang tersedia:**")

        # Kumpulkan data aktual dari semua program aktif
        _rpt_prog_data = []
        _rpt_yr_c = next((c for c in df.columns if c.lower()=="year"),None)
        _rpt_mn_c = next((c for c in df.columns if c.lower()=="month"),None)
        _rpt_sales_c = col(df,"sales")
        _rpt_cust_c  = col(df,"customer")
        _rpt_sp_c    = col(df,"salesperson")
        _rpt_pid_c   = col(df,"principal")

        _now_ts  = pd.Timestamp.now()
        _rpt_yr  = str(_now_ts.year)
        _rpt_mn  = MONTHS.get(_now_ts.month, "")

        # Total semua program aktif (dari session_state filter)
        _rpt_total_act = 0; _rpt_total_tgt = 0
        for _rp in programs:
            _rp_sel = {k:st.session_state.get(f"filter_{k}_{key(_rp).replace(' ','_')}",
                       PROGRAM_FILTER_DEFAULTS.get(_rp,{}).get(k,"All"))
                       for k in ["year","month","principal_name","principal","channel_group","product_code","product"]}
            _rp_sel["metric"] = "Total Relasi Beli Transaksi"
            _rp_scoped, _rp_act, _rp_tgt, _rp_ach, _rp_dev, _ = calculate_program_kpi(_rp, _rp_sel)
            _rpt_total_act += (_rp_act or 0)
            _rpt_total_tgt += (_rp_tgt or 0)
            _rp_ach_fmt = f"{_rp_ach:.1f}%" if _rp_ach else "—"
            _rp_dev_fmt = f"{_rp_dev:,.0f}" if _rp_dev is not None else "—"
            _rpt_prog_data.append({
                "prog": _rp,
                "label": program_label(_rp),
                "actual": _rp_act or 0,
                "target": _rp_tgt or 0,
                "achievement": _rp_ach,
                "deviation": _rp_dev,
            })

        _rpt_total_ach = _rpt_total_act/_rpt_total_tgt*100 if _rpt_total_tgt else None

        # Build variabel dict
        _vars = {
            "BULAN":           _rpt_mn,
            "TAHUN":           _rpt_yr,
            "TANGGAL":         _now_ts.strftime("%d %B %Y"),
            "TOTAL_ACTUAL":    f"{_rpt_total_act:,.0f}",
            "TOTAL_TARGET":    f"{_rpt_total_tgt:,.0f}",
            "TOTAL_ACHIEVEMENT": f"{_rpt_total_ach:.1f}%" if _rpt_total_ach else "—",
            "TOTAL_DEVIATION": f"{_rpt_total_act-_rpt_total_tgt:,.0f}",
            "JUMLAH_PROGRAM":  str(len(programs)),
        }
        # Per-program variables
        for _idx, _pd in enumerate(_rpt_prog_data, 1):
            _pfx = f"PROGRAM{_idx}"
            _vars[f"{_pfx}_NAMA"]        = _pd["label"]
            _vars[f"{_pfx}_ACTUAL"]      = f"{_pd['actual']:,.0f}"
            _vars[f"{_pfx}_TARGET"]      = f"{_pd['target']:,.0f}"
            _vars[f"{_pfx}_ACHIEVEMENT"] = f"{_pd['achievement']:.1f}%" if _pd['achievement'] else "—"
            _vars[f"{_pfx}_DEVIATION"]   = f"{_pd['deviation']:,.0f}" if _pd['deviation'] is not None else "—"

        # Tampilkan daftar variabel tersedia
        _var_rows = [{"Variabel": f"{{{{{k}}}}}", "Nilai Saat Ini": v} for k,v in _vars.items()]
        _var_df = pd.DataFrame(_var_rows)
        st.dataframe(_var_df, width="stretch", hide_index=True,
                     height=38+len(_var_df)*32)

        # Download template contoh
        _sample_template = (
            "LAPORAN PROGRAM PUSH CHP AAM\n"
            "Periode: {{BULAN}} {{TAHUN}}\n"
            "Tanggal: {{TANGGAL}}\n\n"
            "RINGKASAN PENCAPAIAN\n"
            "Total Actual: {{TOTAL_ACTUAL}} relasi\n"
            "Total Target: {{TOTAL_TARGET}} relasi\n"
            "Achievement:  {{TOTAL_ACHIEVEMENT}}\n"
            "Deviasi:      {{TOTAL_DEVIATION}} relasi\n\n"
            "DETAIL PER PROGRAM\n"
        )
        for _i in range(1, 6):
            _sample_template += (
                f"Program {_i}: {{{{PROGRAM{_i}_NAMA}}}} — "
                f"Actual {{{{PROGRAM{_i}_ACTUAL}}}} / Target {{{{PROGRAM{_i}_TARGET}}}} "
                f"({{{{PROGRAM{_i}_ACHIEVEMENT}}}})\n"
            )
        st.download_button(
            "⬇ Download template contoh (.txt)",
            data=_sample_template.encode("utf-8"),
            file_name="template_report_chp.txt",
            mime="text/plain",
            key="download_sample_template",
        )

    with report_col2:
        st.markdown("#### 📝 Hasil Report")

        if template_file is not None:
            _tpl_text = template_file.read().decode("utf-8", errors="replace")

            # Substitusi semua variabel {{VAR}} → nilai
            _result_text = _tpl_text
            import re as _re
            for _vk, _vv in _vars.items():
                _result_text = _result_text.replace(f"{{{{{_vk}}}}}", _vv)

            # Cari variabel yang tidak disubstitusi
            _unfilled = _re.findall(r"\{\{([A-Z0-9_]+)\}\}", _result_text)
            if _unfilled:
                st.warning(f"Variabel tidak dikenal: {', '.join(set(_unfilled))} — dibiarkan apa adanya.")

            # Tampilkan hasil
            st.text_area(
                "Preview hasil report",
                value=_result_text,
                height=420,
                key="report_result_preview",
            )

            # Export hasil
            _dl1, _dl2 = st.columns(2)
            with _dl1:
                st.download_button(
                    "⬇ Download hasil (.txt)",
                    data=_result_text.encode("utf-8"),
                    file_name=f"report_{_rpt_mn}_{_rpt_yr}.txt",
                    mime="text/plain",
                    key="export_report_txt",
                )
            with _dl2:
                # Export juga sebagai CSV variabel
                st.download_button(
                    "⬇ Download variabel (.csv)",
                    data=_var_df.to_csv(index=False).encode("utf-8-sig"),
                    file_name=f"variabel_report_{_rpt_mn}_{_rpt_yr}.csv",
                    mime="text/csv",
                    key="export_report_vars",
                )
        else:
            st.info("Upload template .txt di sebelah kiri untuk melihat hasil report.")
            # Tampilkan preview variabel saja
            st.markdown("**Preview data yang akan digunakan:**")
            for _pd in _rpt_prog_data:
                ach_c = "#39C77B" if (_pd["achievement"] or 0)>=90 else ("#F5C542" if (_pd["achievement"] or 0)>=60 else "#C8102E")
                _ach_str = f"{_pd['achievement']:.1f}%" if _pd['achievement'] else "—"
                st.markdown(
                    f"<div style='background:#0C1D42;border-left:4px solid {ach_c};"
                    f"border-radius:8px;padding:8px 14px;margin-bottom:6px'>"
                    f"<b style='color:#FCF1CF'>{_pd['label']}</b> · "
                    f"<span style='color:#B9CADB'>Actual: {_pd['actual']:,.0f} · "
                    f"Target: {_pd['target']:,.0f} · A/T: {_ach_str}</span>"
                    f"</div>",
                    unsafe_allow_html=True,
                )

elif main_section=="Performance Sales":
    st.markdown("<div class='score-heading'>PERFORMANCE SALES · GROSS</div>",unsafe_allow_html=True)
    bg_col=col(df,"business_group"); date_col=col(df,"date"); principal_col=col(df,"principal"); sales_col=col(df,"sales"); target_value_col=col(df,"target_value")
    if not all([bg_col,date_col,principal_col,sales_col,target_value_col]):
        missing=[label for label,c in [("Bisnis Group",bg_col),("Full Date",date_col),("Principal ID",principal_col),("Sales Value Gross",sales_col),("Target Value",target_value_col)] if not c]
        st.error("Kolom database belum lengkap untuk Performance Sales: "+", ".join(missing))
    else:
        perf=df.copy()
        perf[date_col]=pd.to_datetime(perf[date_col],errors="coerce")
        perf=perf[perf[date_col].notna()].copy()
        perf["_perf_year"]=perf[date_col].dt.year
        perf["_perf_month"]=perf[date_col].dt.month
        perf["_perf_sales"]=pd.to_numeric(perf[sales_col],errors="coerce").fillna(0)
        perf["_perf_target"]=pd.to_numeric(perf[target_value_col],errors="coerce").fillna(0)
        bg_options=sorted(perf[bg_col].dropna().astype(str).unique().tolist())
        years=sorted(perf["_perf_year"].dropna().astype(int).unique().tolist())
        default_year=max(years) if years else 2026
        sp_col=col(df,"salesperson")
        sp_options=sorted(perf[sp_col].dropna().astype(str).unique().tolist()) if sp_col else []
        fc1,fc2,fc3,fc4=st.columns([2,1,1,2])
        with fc1:
            bg_saved=settings.get("performance_filters",{}).get("business_group",[])
            bg_default=[x for x in bg_saved if x in bg_options] if isinstance(bg_saved,list) else []
            selected_bg=st.multiselect("Bisnis Group",bg_options,default=bg_default,key="performance_business_group",placeholder="Semua Bisnis Group")
        with fc4:
            sp_saved=settings.get("performance_filters",{}).get("salesperson",[])
            sp_default=[x for x in sp_saved if x in sp_options] if isinstance(sp_saved,list) else []
            selected_sp=st.multiselect("Sales Person Title",sp_options,default=sp_default,key="performance_salesperson",placeholder="Semua Sales Person") if sp_options else []
        with fc2:
            selected_year=st.selectbox("Year",years,index=years.index(int(settings.get("performance_filters",{}).get("year",default_year))) if int(settings.get("performance_filters",{}).get("year",default_year)) in years else (years.index(default_year) if default_year in years else 0),key="performance_year") if years else default_year
        with fc3:
            months=sorted(perf.loc[perf["_perf_year"]==selected_year,"_perf_month"].dropna().astype(int).unique().tolist())
            default_month=int(settings.get("performance_filters",{}).get("month",max(months) if months else 1))
            selected_month=st.selectbox("Month",months,index=months.index(default_month) if default_month in months else max(0,len(months)-1),format_func=lambda m:MONTHS.get(int(m),str(m)),key="performance_month") if months else 1
        settings["performance_filters"]={"business_group":selected_bg,"year":int(selected_year),"month":int(selected_month),"salesperson":selected_sp}

        # ── Scope — semua filter ikut (BG, SP, Year) ─────────────────────────
        # Target Value di database diisi per baris transaksi, di-sum langsung.
        scoped = perf[(perf["_perf_year"] == int(selected_year))]
        if selected_bg:
            scoped = scoped[scoped[bg_col].astype(str).isin(selected_bg)]
        if selected_sp and sp_col:
            scoped = scoped[scoped[sp_col].astype(str).isin(selected_sp)]
        month_df = scoped[scoped["_perf_month"] == int(selected_month)]
        ytd_df   = scoped[scoped["_perf_month"] <= int(selected_month)]

        def perf_totals(data):
            actual = float(data["_perf_sales"].sum())
            target = float(data["_perf_target"].sum())
            ach    = actual / target * 100 if target else None
            return actual, target, ach, actual - target

        m_actual, m_target, m_ach, m_dev = perf_totals(month_df)
        y_actual, y_target, y_ach, y_dev = perf_totals(ytd_df)
        month_name = MONTHS.get(int(selected_month), str(selected_month))

        # ── Fungsi styling tabel Performance — tersedia untuk semua sub menu ──
        def soften_performance_table(data):
            data = data.copy()
            # Keep identifier/name columns as text: coercing IDs such as FPP or
            # salesperson titles to numbers turns every visible value into NaN.
            for c in ["Principal ID", "Sales Person Title"]:
                if c in data.columns:
                    text = data[c].astype("string").str.strip()
                    data[c] = text.replace({"": pd.NA, "None": pd.NA, "nan": pd.NA, "<NA>": pd.NA}).fillna("(Tidak tercantum)")
            # Pastikan semua kolom numerik benar-benar numeric dtype
            for c in data.columns:
                if c not in ["Principal ID","Sales Person Title","A/T","Tipe","Sumber","Parameter","Range","No"]:
                    try:
                        data[c] = pd.to_numeric(data[c], errors="coerce")
                    except Exception:
                        pass
            # Streamlit's grid can render numeric NaN as the theme's low-contrast
            # "None" placeholder, ignoring Styler.na_rep. Render missing A/T as
            # an explicit string so it inherits the readable grid text color.
            if "A/T" in data.columns:
                at_values = pd.to_numeric(data["A/T"], errors="coerce")
                data["A/T"] = at_values.map(lambda value: "—" if pd.isna(value) else f"{value:,.1f}%")
            n = len(data)
            row_colors = ["background-color: #F4F6F9" if i%2==0 else "background-color: #E9EEF5" for i in range(n)]
            def _stripe(col_series): return [row_colors[i] for i in range(len(col_series))]
            fmt_dict = {c:"{:,.0f}" for c in data.columns
                        if pd.api.types.is_numeric_dtype(data[c]) and c!="A/T"}
            styled = data.style\
                .set_table_styles([{"selector":"th","props":[("background","linear-gradient(180deg,#33509B 0%,#22396F 55%,#182A52 100%)"),("color","#FCF1CF"),("border-color","#0C1D42"),("font-weight","800"),("letter-spacing",".03em"),("text-shadow","0 1px 2px rgba(0,0,0,.45)"),("box-shadow","inset 0 1px 0 rgba(255,255,255,.22), inset 0 -2px 0 rgba(0,0,0,.30)")]},{"selector":"td","props":[("border-color","#DDE4EC")]},{"selector":"table","props":[("background-color","#F4F6F9"),("border-collapse","collapse")] }])\
                .set_properties(**{"background-color":"#F4F6F9","color":"#000000"})\
                .apply(_stripe,axis=0)\
                .set_properties(subset=[c for c in data.columns if c in ["Principal ID", "Sales Person Title"]], **{"color":"#000000", "font-weight":"600"})\
                .format(fmt_dict,na_rep="—")
            for column in data.columns:
                label = str(column).strip().casefold()
                if label.startswith("a/t") or label.startswith("achievement") or label.startswith("pencapaian"):
                    styled = styled.map(lambda value: "background-color:#D7F2E2;color:#146C3E;font-weight:800" if (_performance_value(value) is not None and _performance_value(value) >= 100) else "", subset=[column])
                elif any(token in label for token in ("dev", "deviation", "gap")):
                    def _dev_color(value):
                        number = _performance_value(value)
                        if number is None: return ""
                        if number < 0: return "background-color:#FCE1E4;color:#9E1B32;font-weight:800"
                        if number > 0: return "background-color:#D7F2E2;color:#146C3E;font-weight:700"
                        return ""
                    styled = styled.map(_dev_color, subset=[column])
            return styled

        def bold_performance_total(data):
            """Keep the aggregate visually distinct in its fixed bottom row."""
            return soften_performance_table(data).set_properties(
                **{"font-weight":"800","border-top":"2px solid #22396F"}
            )

        def render_sortable_total_table(detail, total, height):
            """Render a sortable table with its aggregate pinned in the same table."""
            columns=list(detail.columns)
            def _value(v):
                try:
                    if pd.isna(v): return ""
                except (TypeError,ValueError):
                    pass
                return str(v)
            def _cell(v,column,is_total=False):
                raw=_value(v)
                number=_performance_value(v)
                label=str(column).strip().casefold()
                if label.startswith("a/t"):
                    shown="—" if number is None else f"{number:,.1f}%"
                elif number is not None and not isinstance(v,str):
                    shown=f"{number:,.0f}"
                else:
                    shown=raw or "—"
                classes=[]
                if label.startswith("a/t") and number is not None and number>=100:
                    classes.append("positive")
                if any(token in label for token in ("dev","deviation","gap")) and number is not None:
                    classes.append("negative" if number<0 else "positive" if number>0 else "")
                cls=" ".join(x for x in classes if x)
                return f'<td class="{cls}" data-value="{html.escape(raw,quote=True)}">{html.escape(shown)}</td>'
            def _row(values,is_total=False):
                cells="".join(_cell(values.get(c),c,is_total) for c in columns)
                return f'<tr>{cells}</tr>'
            header="".join(f'<th data-col="{i}">{html.escape(str(c))}<span class="sort-mark"> ↕</span></th>' for i,c in enumerate(columns))
            rows="".join(_row(row) for row in detail.to_dict(orient="records"))
            total_row=total.iloc[0].to_dict() if isinstance(total,pd.DataFrame) else dict(total)
            markup=f'''<div class="table-wrap"><table id="perf-sort-table"><thead><tr>{header}</tr></thead>
<tbody>{rows}</tbody><tfoot>{_row(total_row,True)}</tfoot></table></div>
<style>
html,body{{margin:0;padding:0;background:transparent;font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif}}
.table-wrap{{width:100%;overflow:visible}}table{{width:100%;border-collapse:collapse;table-layout:auto;background:#F4F7FB;border:1px solid #C7D4E5}}
th{{background:linear-gradient(180deg,#335A9F 0%,#24477F 100%);color:#fff;font-weight:800;font-size:13px;line-height:1.1;border:1px solid #1D3B70;padding:7px 6px;text-align:left;white-space:nowrap;cursor:pointer}}
td{{font-size:13px;line-height:1.1;border:1px solid #D4DEEA;color:#1B2E46;padding:7px 6px;white-space:nowrap}}
tbody tr:nth-child(odd) td{{background:#F4F7FB}}tbody tr:nth-child(even) td{{background:#E8EEF6}}
tbody td:first-child,tfoot td:first-child{{font-weight:600}}
td.positive{{background:#D7F2E2!important;color:#146C3E;font-weight:800}}td.negative{{background:#FCE1E4!important;color:#9E1B32;font-weight:800}}
tfoot td{{background:#DCE6F3!important;color:#1B2E46;font-weight:900;border-top:3px solid #335A9F}}
.sort-mark{{opacity:.65;font-size:10px}}th.sorted .sort-mark{{opacity:1}}
</style>
<script>
(()=>{{const table=document.getElementById('perf-sort-table');if(!table)return;let last=-1,asc=true;
table.querySelectorAll('thead th').forEach((th,index)=>th.addEventListener('click',()=>{{
if(last===index)asc=!asc;else{{last=index;asc=true}}
table.querySelectorAll('thead th').forEach(h=>h.classList.remove('sorted'));th.classList.add('sorted');
const body=table.tBodies[0];const rows=Array.from(body.rows);rows.sort((a,b)=>{{
const av=a.cells[index].dataset.value||a.cells[index].textContent.trim();const bv=b.cells[index].dataset.value||b.cells[index].textContent.trim();
const an=Number(av.replace(/,/g,'')),bn=Number(bv.replace(/,/g,''));const numeric=av!==''&&bv!==''&&Number.isFinite(an)&&Number.isFinite(bn);
const result=numeric?an-bn:av.localeCompare(bv,undefined,{{numeric:true,sensitivity:'base'}});return (asc?1:-1)*result;
}});rows.forEach(row=>body.appendChild(row));
}}));}})();
</script>'''
            st.components.v1.html(markup,height=height,scrolling=False)

        perf_view=st.radio("Sub menu Performance Sales",["Monitoring Sales","Toserba","Customer di Bawah Rata-rata"],horizontal=True,key="performance_view")
        if perf_view=="Monitoring Sales":
            st.markdown(f"#### {month_name} {int(selected_year)}")
            top=st.columns(4,gap="small")
            monthly_metrics=[(f"Actual Gross · {month_name}",m_actual,"#22396F"),(f"Target · {month_name}",m_target,"#0C1D42"),("A/T",m_ach,"#22396F"),("Deviasi",m_dev,"#0C1D42")]
            for c,(label,value,color) in zip(top,monthly_metrics):
                with c:
                    display=f"{value:,.1f}%" if label=="A/T" and value is not None else ("—" if value is None else f"{value:,.0f}")
                    st.markdown(f"<div style='background:{color};border:1px solid #344a78;border-radius:12px;padding:12px 15px;min-height:92px;box-shadow:0 5px 0 #000737'><div style='font-size:.82rem;color:#FCF1CF;opacity:.85'>{label}</div><div style='font-size:1.55rem;font-weight:800;color:#FCF1CF;margin-top:8px'>{display}</div></div>",unsafe_allow_html=True)
            cum=st.columns(4,gap="small")
            cum_metrics=[("Actual Gross · YTD",y_actual),("Target · YTD",y_target),("A/T · YTD",y_ach),("Deviasi · YTD",y_dev)]
            for c,(label,value) in zip(cum,cum_metrics):
                with c:
                    display=f"{value:,.1f}%" if label.startswith("A/T") and value is not None else ("—" if value is None else f"{value:,.0f}")
                    st.markdown(f"<div style='background:#0C1D42;border:1px solid #344a78;border-radius:12px;padding:10px 14px;min-height:78px;box-shadow:0 4px 0 #000737'><div style='font-size:.78rem;color:#FCF1CF;opacity:.8'>{label}</div><div style='font-size:1.28rem;font-weight:800;color:#FCF1CF;margin-top:6px'>{display}</div></div>",unsafe_allow_html=True)
            m_principal=month_df.groupby(principal_col,dropna=False).agg(ACT=("_perf_sales","sum"),TGT=("_perf_target","sum")).reset_index()
            m_principal=m_principal.rename(columns={principal_col:"Principal ID"})
            y_principal=ytd_df.groupby(principal_col,dropna=False).agg(ACT=("_perf_sales","sum"),TGT=("_perf_target","sum")).reset_index().rename(columns={principal_col:"Principal ID"})
            for _principal_frame in (m_principal, y_principal):
                _principal_frame["Principal ID"] = _principal_frame["Principal ID"].astype("string").str.strip()
                _principal_frame["Principal ID"] = _principal_frame["Principal ID"].replace({"": pd.NA, "None": pd.NA, "nan": pd.NA, "<NA>": pd.NA}).fillna("(Tidak tercantum)")
            principal_table=m_principal.merge(y_principal,on="Principal ID",how="outer",suffixes=("_Bulanan","_YTD")).fillna(0)
            for suffix in ["_Bulanan","_YTD"]:
                principal_table[f"A/T{suffix}"]=principal_table[f"ACT{suffix}"].div(principal_table[f"TGT{suffix}"].replace(0,float("nan")))*100
                principal_table[f"DEV{suffix}"]=principal_table[f"ACT{suffix}"]-principal_table[f"TGT{suffix}"]
            principal_table=principal_table.rename(columns={"ACT_Bulanan":"ACT Bulanan","TGT_Bulanan":"TGT Bulanan","A/T_Bulanan":"A/T Bulanan","DEV_Bulanan":"DEV Bulanan","ACT_YTD":"ACT Kumulatif","TGT_YTD":"TGT Kumulatif","A/T_YTD":"A/T Kumulatif","DEV_YTD":"DEV Kumulatif"})
            monthly_over=int((principal_table["A/T Bulanan"]>100).sum())
            ytd_over=int((principal_table["A/T Kumulatif"]>100).sum())
            trend_actual=scoped.groupby("_perf_month")["_perf_sales"].sum().rename("Sales")
            trend_target=scoped.groupby("_perf_month")["_perf_target"].sum().rename("Target")
            trend=pd.concat([trend_actual,trend_target],axis=1).fillna(0).reset_index().sort_values("_perf_month")
            trend["Month"]=trend["_perf_month"].map(lambda m:MONTHS.get(int(m),str(m)))
            trend["A/T %"]=trend["Sales"].div(trend["Target"].replace(0,float("nan")))*100
            fot=int(((trend["Target"]>0)&(trend["Sales"]>=trend["Target"])).sum())
            chart_header, fot_box=st.columns([5,1])
            with chart_header: st.markdown("#### Trend Sales Bulanan · Actual vs Target")
            with fot_box:
                st.markdown(f"<div style='background:#858585;border-radius:6px;text-align:center;padding:4px 8px;color:white;font-weight:800'>FOT {int(selected_year)}<br><span style='font-size:1.45rem'>{fot}x</span></div>",unsafe_allow_html=True)
            at_max=float(trend["A/T %"].dropna().max()) if trend["A/T %"].notna().any() else 150
            at_axis_max=max(150,math.ceil(at_max/25)*25)
            bars=alt.Chart(trend).transform_fold(["Sales","Target"],as_=["Jenis","Nilai"]).mark_bar(size=25,opacity=.9).encode(
                x=alt.X("Month:N",sort=trend["Month"].tolist(),axis=alt.Axis(title="BULAN",labelAngle=-30,labelPadding=7)),
                xOffset=alt.XOffset("Jenis:N"),y=alt.Y("Nilai:Q",title="Sales Value Gross / Target",scale=alt.Scale(domainMin=0),axis=alt.Axis(format="~s",titlePadding=10)),
                color=alt.Color("Jenis:N",scale=alt.Scale(domain=["Sales","Target"],range=["#B9CADB","#83AED0"]),legend=alt.Legend(orient="top",direction="horizontal",title=None,offset=4)),
                tooltip=["Month:N","Jenis:N",alt.Tooltip("Nilai:Q",format=",.0f")])
            line=alt.Chart(trend).mark_line(color="#796FAE",strokeDash=[7,5],strokeWidth=3,point=alt.OverlayMarkDef(filled=True,size=58)).encode(
                x=alt.X("Month:N",sort=trend["Month"].tolist()),y=alt.Y("A/T %:Q",title="A/T (%)",scale=alt.Scale(domain=[0,at_axis_max]),axis=alt.Axis(orient="right",format=".0f",labelExpr="datum.value + '%'",grid=False,titlePadding=10)),
                tooltip=["Month:N",alt.Tooltip("A/T %:Q",title="A/T",format=".1f")])
            combo=alt.layer(bars,line).resolve_scale(y="independent").properties(height=350,padding={"left":10,"right":8,"top":8,"bottom":2})
            combo=combo.configure(background="#F4F6F9").configure_view(fill="#F4F6F9",stroke="#E1E7EF",strokeWidth=1).configure_axis(labelFontSize=11,titleFontSize=12,labelColor="#526176",titleColor="#3E4D62",domainColor="#CBD4DF",tickColor="#CBD4DF",grid=True,gridColor="#E0E6ED",gridDash=[3,4],titlePadding=8).configure_legend(labelFontSize=11,labelColor="#46566B",symbolSize=130,padding=6)
            st.altair_chart(combo,width="stretch")

            # ═══════════════════════════════════════════════════════════════
            # AI INSIGHT — Strategi Capai Target Bulanan
            # ═══════════════════════════════════════════════════════════════
            st.markdown("#### 🧠 AI Insight · Strategi Capai Target Bulanan")
            days_in_month=pd.Timestamp(int(selected_year),int(selected_month),1).days_in_month
            today=pd.Timestamp.now()
            if today.year==int(selected_year) and today.month==int(selected_month):
                days_passed=today.day
                days_left=days_in_month-days_passed
            else:
                days_passed=days_in_month
                days_left=0
            gap=m_target-m_actual if m_target and m_actual is not None else 0
            daily_run_rate=m_actual/days_passed if days_passed>0 else 0
            daily_needed=gap/days_left if days_left>0 and gap>0 else 0
            run_rate_gap=daily_needed-daily_run_rate if daily_needed>0 else 0
            pct_time_used=days_passed/days_in_month*100
            pct_target_done=m_ach if m_ach is not None else 0

            # Analisis trend — bandingkan 3 bulan terakhir
            recent_trend=trend.tail(3)
            trend_ats=recent_trend["A/T %"].dropna().tolist()
            if len(trend_ats)>=2:
                trend_dir="naik" if trend_ats[-1]>trend_ats[-2] else ("turun" if trend_ats[-1]<trend_ats[-2] else "stagnan")
            else:
                trend_dir="belum cukup data"

            # Analisis principal — siapa yang paling jauh dari target
            pid_below=principal_table[principal_table["DEV Bulanan"]<0].sort_values("DEV Bulanan")
            pid_above=principal_table[principal_table["ACT Bulanan"]>0].sort_values("ACT Bulanan",ascending=False)
            top_pid=pid_above.iloc[0]["Principal ID"] if not pid_above.empty else "—"
            top_pid_act=pid_above.iloc[0]["ACT Bulanan"] if not pid_above.empty else 0
            top_pid_at=pid_above.iloc[0]["A/T Bulanan"] if not pid_above.empty else 0
            worst_pids=pid_below.head(3)

            # Analisis Sales Person Title jika tersedia
            sp_insight_lines=[]
            if sp_col:
                sp_month=month_df.groupby(sp_col,dropna=False).agg(ACT=("_perf_sales","sum"),TGT=("_perf_target","sum")).reset_index()
                sp_month["DEV"]=sp_month["ACT"]-sp_month["TGT"]
                sp_month["AT"]=sp_month["ACT"]/sp_month["TGT"].replace(0,float("nan"))*100
                sp_below=sp_month[sp_month["DEV"]<0].sort_values("DEV")
                sp_above=sp_month[sp_month["AT"]>=100]
                if not sp_above.empty:
                    sp_insight_lines.append(f"**{len(sp_above)} Sales Person Title** sudah melampaui target bulan ini — jadikan benchmark dan distribusikan praktik terbaik mereka ke tim lain.")
                if not sp_below.empty:
                    top3_sp=sp_below.head(3)
                    sp_names=", ".join(f"**{r[sp_col]}** (gap {r['DEV']:,.0f})" for _,r in top3_sp.iterrows())
                    sp_insight_lines.append(f"Sales Person Title dengan gap terbesar: {sp_names} — prioritaskan coaching intensif dan redistribusi rute kunjungan.")

            # Analisis bulan terbaik historis
            best_month_row=trend[trend["Sales"]==trend["Sales"].max()].iloc[0] if not trend.empty else None
            best_month_name=best_month_row["Month"] if best_month_row is not None else "—"
            best_month_sales=best_month_row["Sales"] if best_month_row is not None else 0

            # Tentukan status & tone
            _ach_fmt = m_ach if m_ach is not None else 0.0
            if pct_target_done>=100:
                status_icon="✅"; status_label="TARGET TERCAPAI"
                headline=f"Target {month_name} sudah terlampaui ({_ach_fmt:,.1f}%). Fokus sekarang adalah mempertahankan momentum dan mengamankan buffer lebih besar untuk bulan berikutnya."
            elif pct_target_done>=90:
                status_icon="🟡"; status_label="HAMPIR TERCAPAI"
                headline=f"Pencapaian {_ach_fmt:,.1f}% — dalam jangkauan target. Kurang **{gap:,.0f}** dalam {days_left} hari tersisa ({daily_needed:,.0f}/hari). Akselerasi kecil di akhir bulan sudah cukup."
            elif pct_target_done>=60:
                status_icon="🟠"; status_label="PERLU AKSELERASI"
                headline=f"Pencapaian baru {_ach_fmt:,.1f}% padahal {pct_time_used:.0f}% waktu bulan sudah terpakai. Butuh akselerasi nyata — gap **{gap:,.0f}** harus ditutup dalam {days_left} hari."
            else:
                status_icon="🔴"; status_label="KRITIS — INTERVENSI SEGERA"
                headline=f"Pencapaian hanya {_ach_fmt:,.1f}% dengan {days_left} hari tersisa. Gap **{gap:,.0f}** sangat besar — perlu intervensi langsung dari AAM dan redistribusi sumber daya segera."

            # Render insight card
            insight_bg={"✅":"#0d3b1e","🟡":"#3b2e00","🟠":"#3b1a00","🔴":"#3b0000"}[status_icon]
            insight_border={"✅":"#39C77B","🟡":"#F5C542","🟠":"#FF7A00","🔴":"#C8102E"}[status_icon]
            st.markdown(f"""
<div style='background:{insight_bg};border-left:5px solid {insight_border};border-radius:10px;padding:18px 22px;margin-bottom:12px'>
<div style='color:{insight_border};font-weight:900;font-size:.85rem;letter-spacing:.08em'>{status_icon} {status_label}</div>
<div style='color:#F0F4FF;font-size:1.05rem;font-weight:600;margin-top:6px;line-height:1.5'>{headline}</div>
</div>""",unsafe_allow_html=True)

            # 5 blok strategi detail
            col_a,col_b=st.columns(2,gap="medium")
            with col_a:
                # Blok 1 — Angka kunci
                st.markdown(f"""**📊 Angka Kunci Bulan Ini**
- Actual: **{m_actual:,.0f}** · Target: **{m_target:,.0f}**
- Gap tersisa: **{max(gap,0):,.0f}** dalam **{days_left}** hari
- Run rate harian saat ini: **{daily_run_rate:,.0f}/hari**
- Run rate yang dibutuhkan: **{daily_needed:,.0f}/hari**{"" if daily_needed<=0 else f" (naik **{run_rate_gap:,.0f}** dari sekarang)"}
- Waktu terpakai: **{pct_time_used:.0f}%** — Pencapaian: **{pct_target_done:.1f}%**
- {"⚠️ Pencapaian **tertinggal** dari proporsi waktu — perlu akselerasi." if pct_target_done<pct_time_used-10 else "✅ Pencapaian **seimbang** dengan proporsi waktu berjalan."}""")

                # Blok 2 — Analisis trend
                st.markdown(f"""**📈 Analisis Trend**
- Trend A/T 3 bulan terakhir: **{trend_dir}**{f" ({', '.join(f'{v:.1f}%' for v in trend_ats)})" if trend_ats else ""}
- FOT {int(selected_year)}: **{fot} bulan** dari bulan berjalan
- Bulan terbaik tahun ini: **{best_month_name}** ({best_month_sales:,.0f})
- Principal ID sudah >100% bulan ini: **{monthly_over}**
- {"📌 Trend sedang melemah — review strategi segera." if trend_dir=="turun" else ("📌 Trend menguat — pertahankan momentum." if trend_dir=="naik" else "📌 Trend stagnan — dorong variasi taktik.")}""")

            with col_b:
                # Blok 3 — Fokus Principal
                pid_lines=[f"- 🥇 Principal terkuat: **{top_pid}** (ACT {top_pid_act:,.0f}, A/T {top_pid_at:.1f}%) — perbesar volume order dan frekuensi push."]
                if not worst_pids.empty:
                    for _,r in worst_pids.iterrows():
                        at_val=r["A/T Bulanan"]
                        pid_lines.append(f"- ⚠️ **{r['Principal ID']}**: ACT {r['ACT Bulanan']:,.0f} vs TGT {r['TGT Bulanan']:,.0f} (A/T {at_val:.1f}% · gap {abs(r['DEV Bulanan']):,.0f}) — aktivasi customer lama yang vakum dan intensifkan prospek baru.")
                st.markdown("**🏭 Fokus per Principal ID**\n"+"\n".join(pid_lines))

                # Blok 4 — Fokus Sales Person Title
                if sp_insight_lines:
                    st.markdown("**👤 Fokus per Sales Person Title**\n"+"\n".join(f"- {l}" for l in sp_insight_lines))
                else:
                    st.markdown("**👤 Fokus per Sales Person Title**\n- Data Sales Person Title belum tersedia di filter aktif.")

            # Blok 5 — Rekomendasi taktis (full width)
            rekomendasi=[]
            if days_left>0 and gap>0:
                if daily_needed>daily_run_rate*1.5:
                    rekomendasi.append(f"🔴 **Daily needed {daily_needed:,.0f}/hari** = **{daily_needed/max(daily_run_rate,1):.1f}× lipat** run rate sekarang — tidak bisa dicapai dengan cara biasa. Diperlukan: flash sale, bundling promosi, atau aktivasi batch customer dormant secara serentak.")
                elif daily_needed>daily_run_rate*1.1:
                    rekomendasi.append(f"🟠 **Perlu menaikkan run rate {(run_rate_gap/max(daily_run_rate,1)*100):.0f}%** — lakukan daily review dengan setiap Sales Person Title, pastikan rute Fixed Call dieksekusi penuh tanpa skip.")
                else:
                    rekomendasi.append(f"🟢 **Gap masih dalam jangkauan** — pertahankan disiplin eksekusi harian dan fokus pada upselling di customer yang sudah aktif transaksi.")
            if not pid_below.empty:
                rekomendasi.append(f"📋 **{len(pid_below)} Principal ID** masih di bawah target bulanan — lakukan weekly check-in khusus per principal, identifikasi apakah hambatan ada di sisi stok, harga, atau coverage.")
            if fot<int(selected_month)//2:
                rekomendasi.append(f"📉 **FOT rendah ({fot}× dari {int(selected_month)} bulan berjalan)** — konsistensi capai target bulanan lemah. Evaluasi apakah target terlalu tinggi atau eksekusi tidak terstruktur.")
            rekomendasi.append("📅 **Action harian**: Mulai setiap hari dengan review gap per Sales Person Title, tutup hari dengan konfirmasi order yang sudah masuk. Jangan tunggu akhir bulan untuk koreksi.")
            st.markdown("**⚡ Rekomendasi Taktis**")
            for r in rekomendasi:
                st.markdown(f"- {r}")
            st.divider()
            # ═══════════════════════════════════════════════════════════════

            st.markdown("#### Detail Performance Sales")
            st.caption("Pilih tampilan detail. Kedua tabel mengikuti filter dan dasar perhitungan KPI yang sama.")
            k1,k2=st.columns(2)
            k1.metric("Principal ID >100% · Bulanan",f"{monthly_over:,}")
            k2.metric("Principal ID >100% · Kumulatif",f"{ytd_over:,}")
            detail_view=st.radio("Tampilan tabel",["By Principal ID","By Sales Person Title"],horizontal=True,key="performance_detail_view")
            if detail_view=="By Principal ID":
                st.markdown("**Monitoring by Principal ID**")
                st.caption("Dua panel memakai urutan Principal ID yang sama: ACT, TGT, A/T, dan DEV.")
                principal_ordered=principal_table.sort_values("ACT Bulanan",ascending=False).copy()
                principal_total={"Principal ID":"TOTAL","ACT Bulanan":m_actual,"TGT Bulanan":m_target,"A/T Bulanan":m_ach,"DEV Bulanan":m_dev,"ACT Kumulatif":y_actual,"TGT Kumulatif":y_target,"A/T Kumulatif":y_ach,"DEV Kumulatif":y_dev}
                pm,py=st.columns(2,gap="small")
                with pm:
                    st.markdown(f"**Bulanan · {month_name} {int(selected_year)}**")
                    month_cols=["Principal ID","ACT Bulanan","TGT Bulanan","A/T Bulanan","DEV Bulanan"]
                    month_view=principal_ordered[month_cols].rename(columns={"ACT Bulanan":"ACT","TGT Bulanan":"TGT","A/T Bulanan":"A/T","DEV Bulanan":"DEV"})
                    month_total_view=pd.DataFrame([{key:principal_total[key] for key in month_cols}]).rename(columns={"ACT Bulanan":"ACT","TGT Bulanan":"TGT","A/T Bulanan":"A/T","DEV Bulanan":"DEV"})
                    render_sortable_total_table(month_view,month_total_view,48+(len(month_view)+1)*35)
                with py:
                    st.markdown(f"**Kumulatif · Jan–{month_name} {int(selected_year)}**")
                    ytd_cols=["Principal ID","ACT Kumulatif","TGT Kumulatif","A/T Kumulatif","DEV Kumulatif"]
                    ytd_view=principal_ordered[ytd_cols].rename(columns={"ACT Kumulatif":"ACT","TGT Kumulatif":"TGT","A/T Kumulatif":"A/T","DEV Kumulatif":"DEV"})
                    ytd_total_view=pd.DataFrame([{key:principal_total[key] for key in ytd_cols}]).rename(columns={"ACT Kumulatif":"ACT","TGT Kumulatif":"TGT","A/T Kumulatif":"A/T","DEV Kumulatif":"DEV"})
                    render_sortable_total_table(ytd_view,ytd_total_view,48+(len(ytd_view)+1)*35)
            else:
                salesperson_col=col(df,"salesperson")
                if not salesperson_col:
                    st.warning("Kolom Sales Person Title tidak ditemukan di database.")
                else:
                    st.markdown("**Actual Sales Value Gross & Target Value by Sales Person Title**")
                    st.caption("Target dijumlahkan dari kolom Target Value per baris, sama dengan kartu KPI dan grafik.")
                    title_month=month_df.groupby(salesperson_col,dropna=False).agg(ACT=("_perf_sales","sum"),TGT=("_perf_target","sum")).reset_index().rename(columns={salesperson_col:"Sales Person Title"})
                    title_ytd=ytd_df.groupby(salesperson_col,dropna=False).agg(ACT=("_perf_sales","sum"),TGT=("_perf_target","sum")).reset_index().rename(columns={salesperson_col:"Sales Person Title"})
                    title_table=title_month.merge(title_ytd,on="Sales Person Title",how="outer",suffixes=("_Bulanan","_Kumulatif"))
                    for value_col in ["ACT_Bulanan","TGT_Bulanan","ACT_Kumulatif","TGT_Kumulatif"]:
                        title_table[value_col]=pd.to_numeric(title_table[value_col],errors="coerce").fillna(0)
                    title_table["Sales Person Title"]=title_table["Sales Person Title"].replace("","(Tidak tercantum)").fillna("(Tidak tercantum)").astype(str)
                    for suffix in ["_Bulanan","_Kumulatif"]:
                        title_table[f"A/T{suffix}"]=title_table[f"ACT{suffix}"].div(title_table[f"TGT{suffix}"].replace(0,float("nan")))*100
                        title_table[f"DEV{suffix}"]=title_table[f"ACT{suffix}"]-title_table[f"TGT{suffix}"]
                    title_table=title_table.rename(columns={"ACT_Bulanan":"ACT Bulanan","TGT_Bulanan":"TGT Bulanan","A/T_Bulanan":"A/T Bulanan","DEV_Bulanan":"DEV Bulanan","ACT_Kumulatif":"ACT Kumulatif","TGT_Kumulatif":"TGT Kumulatif","A/T_Kumulatif":"A/T Kumulatif","DEV_Kumulatif":"DEV Kumulatif"})
                    total_row={"Sales Person Title":"TOTAL","ACT Bulanan":m_actual,"TGT Bulanan":m_target,"A/T Bulanan":m_ach,"DEV Bulanan":m_dev,"ACT Kumulatif":y_actual,"TGT Kumulatif":y_target,"A/T Kumulatif":y_ach,"DEV Kumulatif":y_dev}
                    title_table=title_table.sort_values("ACT Bulanan",ascending=False).reset_index(drop=True)
                    tm,ty=st.columns(2,gap="small")
                    with tm:
                        st.markdown(f"**Bulanan · {month_name} {int(selected_year)}**")
                        title_month_cols=["Sales Person Title","ACT Bulanan","TGT Bulanan","A/T Bulanan","DEV Bulanan"]
                        title_month_view=title_table[title_month_cols].rename(columns={"ACT Bulanan":"ACT","TGT Bulanan":"TGT","A/T Bulanan":"A/T","DEV Bulanan":"DEV"})
                        title_month_total=pd.DataFrame([{key:total_row[key] for key in title_month_cols}]).rename(columns={"ACT Bulanan":"ACT","TGT Bulanan":"TGT","A/T Bulanan":"A/T","DEV Bulanan":"DEV"})
                        render_sortable_total_table(title_month_view,title_month_total,48+(len(title_month_view)+1)*35)
                    with ty:
                        st.markdown(f"**Kumulatif · Jan–{month_name} {int(selected_year)}**")
                        title_ytd_cols=["Sales Person Title","ACT Kumulatif","TGT Kumulatif","A/T Kumulatif","DEV Kumulatif"]
                        title_ytd_view=title_table[title_ytd_cols].rename(columns={"ACT Kumulatif":"ACT","TGT Kumulatif":"TGT","A/T Kumulatif":"A/T","DEV Kumulatif":"DEV"})
                        title_ytd_total=pd.DataFrame([{key:total_row[key] for key in title_ytd_cols}]).rename(columns={"ACT Kumulatif":"ACT","TGT Kumulatif":"TGT","A/T Kumulatif":"A/T","DEV Kumulatif":"DEV"})
                        render_sortable_total_table(title_ytd_view,title_ytd_total,48+(len(title_ytd_view)+1)*35)
        elif perf_view=="Toserba":
            # ── TOSERBA: Principal ID CHW & ETW ─────────────────────────────
            toserba_view=st.radio("Sub Menu Toserba",["Monitoring Toserba","Target Tantangan"],horizontal=True,key="toserba_submenu")
            st.markdown(f"#### {'Target Tantangan' if toserba_view=='Target Tantangan' else 'Monitoring Toserba'} · {month_name} {int(selected_year)}")
            st.caption("Scope: Principal ID CHW dan ETW · filter Bisnis Group & Sales Person Title mengikuti pilihan di atas.")
            TOSERBA_PIDS = ["CHW","ETW"]

            tsb_base = perf[(perf["_perf_year"]==int(selected_year))]
            if selected_bg:
                tsb_base = tsb_base[tsb_base[bg_col].astype(str).isin(selected_bg)]
            if selected_sp and sp_col:
                tsb_base = tsb_base[tsb_base[sp_col].astype(str).isin(selected_sp)]
            tsb_base = tsb_base[tsb_base[principal_col].astype(str).str.upper().str.strip().isin(TOSERBA_PIDS)]

            tsb_month = tsb_base[tsb_base["_perf_month"]==int(selected_month)]
            tsb_ytd   = tsb_base[tsb_base["_perf_month"]<=int(selected_month)]

            if toserba_view=="Target Tantangan":
                st.caption("Target tantangan per principal dialokasikan otomatis ke Sales Person Title mengikuti proporsi target Toserba asli pada bulan dan filter yang dipilih. Jika target asli tidak tersedia, alokasi dibagi merata ke title pada principal tersebut.")
                _challenge_business_scope="_".join(sorted(selected_bg)) if selected_bg else "all_business"
                _challenge_salesperson_scope="_".join(sorted(selected_sp)) if selected_sp else "all_salespeople"
                _challenge_scope_key=f"{int(selected_year)}_{int(selected_month):02d}_{key(_challenge_business_scope).replace(' ','_')}_{key(_challenge_salesperson_scope).replace(' ','_')}"
                _challenge_store=settings.setdefault("toserba_challenge_targets",{}).setdefault(_challenge_scope_key,{})
                _original_principal_totals={}
                for _challenge_pid in TOSERBA_PIDS:
                    _challenge_pid_data=tsb_month[tsb_month[principal_col].astype(str).str.upper().str.strip().eq(_challenge_pid)]
                    _original_principal_totals[_challenge_pid]=float(pd.to_numeric(_challenge_pid_data["_perf_target"],errors="coerce").fillna(0).sum())
                _challenge_cols=st.columns(2,gap="medium")
                _challenge_inputs={}
                for _challenge_col,_challenge_pid in zip(_challenge_cols,TOSERBA_PIDS):
                    with _challenge_col:
                        _challenge_inputs[_challenge_pid]=st.number_input(
                            f"Target tantangan {_challenge_pid} · Rp",
                            min_value=0.0,
                            value=float(_challenge_store.get(_challenge_pid,_original_principal_totals[_challenge_pid])),
                            step=1_000_000.0,
                            format="%.0f",
                            key=f"toserba_challenge_target_{_challenge_scope_key}_{_challenge_pid}",
                            help=f"Target Toserba asli {_challenge_pid}: Rp {_original_principal_totals[_challenge_pid]:,.0f}. Nilai target tantangan akan dibagi ke Sales Person Title berdasarkan proporsi target asli.",
                        )
                if st.button("Simpan target tantangan",type="primary",key=f"save_toserba_challenge_{_challenge_scope_key}"):
                    _challenge_store.update({pid:float(value) for pid,value in _challenge_inputs.items()})
                    save_settings(settings)
                    st.success(f"Target tantangan {month_name} {int(selected_year)} tersimpan.")

                _challenge_detail_parts=[]
                for _challenge_pid in TOSERBA_PIDS:
                    _challenge_pid_data=tsb_month[tsb_month[principal_col].astype(str).str.upper().str.strip().eq(_challenge_pid)]
                    if not sp_col:
                        st.warning("Kolom Sales Person Title tidak ditemukan; target tantangan hanya dapat ditampilkan per principal.")
                        continue
                    _challenge_by_title=(_challenge_pid_data.groupby(sp_col,dropna=False)
                        .agg({"_perf_sales":"sum","_perf_target":"sum"})
                        .reset_index().rename(columns={sp_col:"Sales Person Title","_perf_sales":"ACT","_perf_target":"Target Toserba Asli"}))
                    _challenge_by_title["Target Toserba Asli"]=pd.to_numeric(_challenge_by_title["Target Toserba Asli"],errors="coerce").fillna(0)
                    _original_sum=float(_challenge_by_title["Target Toserba Asli"].sum())
                    _challenge_total=float(_challenge_inputs[_challenge_pid])
                    if _original_sum>0:
                        _challenge_by_title["Bobot Target (%)"]=_challenge_by_title["Target Toserba Asli"]/_original_sum*100
                        _challenge_by_title["Target Tantangan"]=_challenge_total*_challenge_by_title["Target Toserba Asli"]/_original_sum
                    elif len(_challenge_by_title):
                        _challenge_by_title["Bobot Target (%)"]=100/len(_challenge_by_title)
                        _challenge_by_title["Target Tantangan"]=_challenge_total/len(_challenge_by_title)
                    else:
                        _challenge_by_title["Bobot Target (%)"]=pd.Series(dtype=float)
                        _challenge_by_title["Target Tantangan"]=pd.Series(dtype=float)
                    _challenge_by_title["A/T Tantangan"]=_challenge_by_title["ACT"].div(_challenge_by_title["Target Tantangan"].replace(0,float("nan")))*100
                    _challenge_by_title["Dev Tantangan"]=_challenge_by_title["ACT"]-_challenge_by_title["Target Tantangan"]
                    _challenge_by_title.insert(0,"Principal ID",_challenge_pid)
                    _challenge_detail_parts.append(_challenge_by_title)

                if _challenge_detail_parts:
                    _challenge_detail=pd.concat(_challenge_detail_parts,ignore_index=True)
                    _challenge_detail=_challenge_detail.rename(columns={"A/T Tantangan":"A/T","Dev Tantangan":"Dev"})
                    _challenge_detail=_challenge_detail.sort_values(["Principal ID","Target Tantangan"],ascending=[True,False]).reset_index(drop=True)
                    _challenge_totals=[]
                    for _challenge_pid in TOSERBA_PIDS:
                        _pid_detail=_challenge_detail[_challenge_detail["Principal ID"].eq(_challenge_pid)]
                        if not _pid_detail.empty:
                            _act_sum=float(_pid_detail["ACT"].sum()); _challenge_sum=float(_pid_detail["Target Tantangan"].sum())
                            _challenge_totals.append({"Principal ID":_challenge_pid,"Sales Person Title":"TOTAL","Target Toserba Asli":float(_pid_detail["Target Toserba Asli"].sum()),"Bobot Target (%)":100.0,"ACT":_act_sum,"Target Tantangan":_challenge_sum,"A/T":_act_sum/_challenge_sum*100 if _challenge_sum else None,"Dev":_act_sum-_challenge_sum})
                    if _challenge_totals:
                        _challenge_detail=pd.concat([_challenge_detail,pd.DataFrame(_challenge_totals)],ignore_index=True)
                    st.dataframe(_challenge_detail,width="stretch",hide_index=True,column_config={
                        "Target Toserba Asli":st.column_config.NumberColumn(format="Rp %,.0f"),
                        "Bobot Target (%)":st.column_config.NumberColumn(format="%.1f%%"),
                        "ACT":st.column_config.NumberColumn(format="Rp %,.0f"),
                        "Target Tantangan":st.column_config.NumberColumn(format="Rp %,.0f"),
                        "A/T":st.column_config.NumberColumn(format="%.1f%%"),
                        "Dev":st.column_config.NumberColumn(format="Rp %,.0f"),
                    })
                    st.download_button(
                        f"⬇ Export Target Tantangan {month_name} {int(selected_year)} (.csv)",
                        data=_challenge_detail.to_csv(index=False).encode("utf-8-sig"),
                        file_name=f"target_tantangan_toserba_{month_name}_{int(selected_year)}.csv",
                        mime="text/csv",key=f"export_toserba_challenge_{_challenge_scope_key}",
                    )
                else:
                    st.info("Tidak ada data CHW/ETW untuk bulan dan filter yang dipilih.")
                st.stop()

            tsb_act_m = float(tsb_month["_perf_sales"].sum())
            tsb_tgt_m = float(tsb_month["_perf_target"].sum())
            tsb_ach_m = tsb_act_m/tsb_tgt_m*100 if tsb_tgt_m else None
            tsb_dev_m = tsb_act_m - tsb_tgt_m

            tsb_act_y = float(tsb_ytd["_perf_sales"].sum())
            tsb_tgt_y = float(tsb_ytd["_perf_target"].sum())
            tsb_ach_y = tsb_act_y/tsb_tgt_y*100 if tsb_tgt_y else None
            tsb_dev_y = tsb_act_y - tsb_tgt_y

            # ── KPI cards bulanan — 4 kolom ───────────────────────────────────
            ach_color_tsb = RED if (tsb_ach_m or 0)<60 else ("#F5C542" if (tsb_ach_m or 0)<90 else "#39C77B")
            _cs = "border-radius:10px;padding:9px 12px;min-height:78px;box-shadow:0 4px 0 #000737;border:1px solid #344a78"
            top_tsb = st.columns(4,gap="small")
            for c_tsb,(lbl,val,clr) in zip(top_tsb,[
                (f"Actual · {month_name}", tsb_act_m, BLUE),
                (f"Target · {month_name}", tsb_tgt_m, MIDNIGHT_BLUE),
                ("A/T",                    tsb_ach_m, ach_color_tsb),
                ("Deviasi",                tsb_dev_m, MIDNIGHT_BLUE),
            ]):
                with c_tsb:
                    disp_val = f"{val:,.1f}%" if lbl=="A/T" and val is not None else ("—" if val is None else f"Rp {val:,.0f}")
                    st.markdown(f"<div style='background:{clr};{_cs}'><div style='font-size:.72rem;color:#FCF1CF;opacity:.85'>{lbl}</div><div style='font-size:1.1rem;font-weight:800;color:#FCF1CF;margin-top:5px'>{disp_val}</div></div>",unsafe_allow_html=True)

            # YTD cards — 4 kolom
            cum_tsb = st.columns(4,gap="small")
            for c_tsb,(lbl,val) in zip(cum_tsb,[
                ("Actual YTD",tsb_act_y),("Target YTD",tsb_tgt_y),
                ("A/T YTD",tsb_ach_y),("Deviasi YTD",tsb_dev_y),
            ]):
                with c_tsb:
                    disp_val = f"{val:,.1f}%" if lbl.startswith("A/T") and val is not None else ("—" if val is None else f"Rp {val:,.0f}")
                    st.markdown(f"<div style='background:{MIDNIGHT_BLUE};{_cs}'><div style='font-size:.7rem;color:#FCF1CF;opacity:.8'>{lbl}</div><div style='font-size:1rem;font-weight:800;color:#FCF1CF;margin-top:4px'>{disp_val}</div></div>",unsafe_allow_html=True)

            # ── AI Insight Toserba dengan susunan yang sama seperti Performance Sales ──
            st.markdown("#### 🧠 AI Insight · Strategi Capai Target Bulanan")
            _tsb_days_in_month=pd.Timestamp(int(selected_year),int(selected_month),1).days_in_month
            _tsb_today=pd.Timestamp.now()
            _tsb_is_current=(_tsb_today.year==int(selected_year) and _tsb_today.month==int(selected_month))
            _tsb_days_passed=min(_tsb_today.day,_tsb_days_in_month) if _tsb_is_current else _tsb_days_in_month
            _tsb_days_left=max(_tsb_days_in_month-_tsb_days_passed,0) if _tsb_is_current else 0
            _tsb_gap=tsb_tgt_m-tsb_act_m
            _tsb_daily_actual=tsb_act_m/_tsb_days_passed if _tsb_days_passed else 0
            _tsb_daily_needed=max(_tsb_gap,0)/_tsb_days_left if _tsb_days_left and _tsb_gap>0 else 0
            _tsb_run_rate_gap=_tsb_daily_needed-_tsb_daily_actual if _tsb_daily_needed else 0
            _tsb_pct_time=_tsb_days_passed/_tsb_days_in_month*100 if _tsb_days_in_month else 0
            _tsb_pct_target=tsb_ach_m or 0
            if tsb_tgt_m<=0:
                _tsb_icon="🟡"; _tsb_status="TARGET BELUM ADA"
                _tsb_head="Target bulanan Toserba belum tersedia. Lengkapi target CHW dan ETW agar strategi pencapaian dapat dihitung."
            elif _tsb_pct_target>=100:
                _tsb_icon="✅"; _tsb_status="TARGET TERCAPAI"
                _tsb_head=f"Target Toserba {month_name} sudah terlampaui ({_tsb_pct_target:,.1f}%). Pertahankan momentum dan amankan penjualan sampai akhir bulan."
            elif _tsb_pct_target>=90:
                _tsb_icon="🟡"; _tsb_status="HAMPIR TERCAPAI"
                _tsb_head=f"Pencapaian {_tsb_pct_target:,.1f}% — target masih dalam jangkauan. Kurang Rp {max(_tsb_gap,0):,.0f} dalam {_tsb_days_left} hari ({_tsb_daily_needed:,.0f}/hari)."
            elif _tsb_pct_target>=60:
                _tsb_icon="🟠"; _tsb_status="PERLU AKSELERASI"
                _tsb_head=f"Pencapaian baru {_tsb_pct_target:,.1f}% sementara {_tsb_pct_time:.0f}% waktu bulan telah terpakai. Gap Rp {max(_tsb_gap,0):,.0f} perlu dikejar dalam {_tsb_days_left} hari."
            else:
                _tsb_icon="🔴"; _tsb_status="KRITIS — INTERVENSI SEGERA"
                _tsb_head=f"Pencapaian hanya {_tsb_pct_target:,.1f}% dengan {_tsb_days_left} hari tersisa. Gap Rp {max(_tsb_gap,0):,.0f} sangat besar — perlu intervensi langsung dan redistribusi fokus penjualan."
            _tsb_insight_bg={"✅":"#0d3b1e","🟡":"#3b2e00","🟠":"#3b1a00","🔴":"#3b0000"}[_tsb_icon]
            _tsb_insight_border={"✅":"#39C77B","🟡":"#F5C542","🟠":"#FF7A00","🔴":"#C8102E"}[_tsb_icon]
            st.markdown(f"""
<div style='background:{_tsb_insight_bg};border-left:5px solid {_tsb_insight_border};border-radius:10px;padding:18px 22px;margin-bottom:12px'>
<div style='color:{_tsb_insight_border};font-weight:900;font-size:.85rem;letter-spacing:.08em'>{_tsb_icon} {_tsb_status}</div>
<div style='color:#F0F4FF;font-size:1.05rem;font-weight:600;margin-top:6px;line-height:1.5'>{_tsb_head}</div>
</div>""",unsafe_allow_html=True)

            # Ringkasan trend Toserba historis sesuai filter aktif.
            _tsb_trend=(tsb_base.groupby("_perf_month").agg(ACT=("_perf_sales","sum"),TGT=("_perf_target","sum"))
                        .reset_index().sort_values("_perf_month"))
            _tsb_trend["AT"]=_tsb_trend["ACT"].div(_tsb_trend["TGT"].replace(0,float("nan")))*100
            _tsb_recent=_tsb_trend[_tsb_trend["_perf_month"]<=int(selected_month)].tail(3)
            _tsb_trend_ats=_tsb_recent["AT"].dropna().tolist()
            _tsb_trend_dir=("naik" if _tsb_trend_ats[-1]>_tsb_trend_ats[-2] else "turun" if _tsb_trend_ats[-1]<_tsb_trend_ats[-2] else "stagnan") if len(_tsb_trend_ats)>=2 else "belum cukup data"
            _tsb_fot=int(((_tsb_trend["_perf_month"]<=int(selected_month))&(_tsb_trend["TGT"]>0)&(_tsb_trend["ACT"]>=_tsb_trend["TGT"])).sum())
            _tsb_trend_ytd=_tsb_trend[_tsb_trend["_perf_month"]<=int(selected_month)]
            _tsb_best=_tsb_trend_ytd.loc[_tsb_trend_ytd["ACT"].idxmax()] if not _tsb_trend_ytd.empty else None
            _tsb_best_sales=float(_tsb_best["ACT"]) if _tsb_best is not None else 0
            _tsb_pid_rows=[]
            for _pid in TOSERBA_PIDS:
                _pid_data=tsb_month[tsb_month[principal_col].astype(str).str.upper().str.strip().eq(_pid)]
                _pa=float(_pid_data["_perf_sales"].sum()); _pt=float(_pid_data["_perf_target"].sum())
                _tsb_pid_rows.append({"Principal ID":_pid,"ACT":_pa,"TGT":_pt,"AT":_pa/_pt*100 if _pt else 0,"DEV":_pa-_pt})
            _tsb_pid_frame=pd.DataFrame(_tsb_pid_rows)
            _tsb_pid_below=_tsb_pid_frame[_tsb_pid_frame["DEV"]<0].sort_values("DEV")
            _tsb_pid_best=_tsb_pid_frame.sort_values("AT",ascending=False).iloc[0] if not _tsb_pid_frame.empty else None
            _tsb_title_lines=[]
            if sp_col:
                _tsb_title_frame=(tsb_month.groupby(sp_col,dropna=False).agg(ACT=("_perf_sales","sum"),TGT=("_perf_target","sum"))
                                  .reset_index().rename(columns={sp_col:"Sales Person Title"}))
                _tsb_title_frame["DEV"]=_tsb_title_frame["ACT"]-_tsb_title_frame["TGT"]
                _tsb_title_frame["AT"]=_tsb_title_frame["ACT"].div(_tsb_title_frame["TGT"].replace(0,float("nan")))*100
                _tsb_title_below=_tsb_title_frame[_tsb_title_frame["DEV"]<0].sort_values("DEV").head(3)
                _tsb_title_above=_tsb_title_frame[_tsb_title_frame["AT"]>=100]
                if not _tsb_title_above.empty:
                    _tsb_title_lines.append(f"**{len(_tsb_title_above)} Sales Person Title** sudah mencapai target — pertahankan praktik baik dan jaga order berulang.")
                if not _tsb_title_below.empty:
                    _tsb_title_names=", ".join(f"**{r['Sales Person Title']}** (gap Rp {abs(r['DEV']):,.0f})" for _,r in _tsb_title_below.iterrows())
                    _tsb_title_lines.append(f"Gap terbesar ada pada {_tsb_title_names}; prioritaskan coaching, coverage, dan follow-up order.")

            _tsb_col_a,_tsb_col_b=st.columns(2,gap="medium")
            with _tsb_col_a:
                st.markdown(f"""**📊 Angka Kunci Bulan Ini**
- Actual: **Rp {tsb_act_m:,.0f}** · Target: **Rp {tsb_tgt_m:,.0f}**
- Gap tersisa: **Rp {max(_tsb_gap,0):,.0f}** dalam **{_tsb_days_left}** hari
- Run rate harian saat ini: **Rp {_tsb_daily_actual:,.0f}/hari**
- Run rate yang dibutuhkan: **Rp {_tsb_daily_needed:,.0f}/hari**{'' if _tsb_daily_needed<=0 else f' (naik **Rp {_tsb_run_rate_gap:,.0f}** dari sekarang)'}
- Waktu terpakai: **{_tsb_pct_time:.0f}%** — Pencapaian: **{_tsb_pct_target:.1f}%**
- {'⚠️ Pencapaian **tertinggal** dari proporsi waktu — perlu akselerasi.' if _tsb_pct_target<_tsb_pct_time-10 else '✅ Pencapaian **seimbang** dengan proporsi waktu berjalan.'}""")
                st.markdown(f"""**📈 Analisis Trend**
- Trend A/T 3 bulan terakhir: **{_tsb_trend_dir}**{f" ({', '.join(f'{v:.1f}%' for v in _tsb_trend_ats)})" if _tsb_trend_ats else ''}
- FOT {int(selected_year)}: **{_tsb_fot} bulan** dari bulan berjalan
- Bulan terbaik tahun ini: **{MONTHS.get(int(_tsb_best['_perf_month']), '—') if _tsb_best is not None else '—'}** ({_tsb_best_sales:,.0f})
- Principal Toserba sudah >100% bulan ini: **{int((_tsb_pid_frame['AT']>=100).sum())}**
- {'📌 Trend sedang melemah — review strategi segera.' if _tsb_trend_dir=='turun' else ('📌 Trend menguat — pertahankan momentum.' if _tsb_trend_dir=='naik' else '📌 Trend stagnan — dorong variasi taktik.')}""")
            with _tsb_col_b:
                st.markdown("**🏭 Fokus per Principal ID**")
                if _tsb_pid_best is not None:
                    st.markdown(f"- 🥇 Principal terkuat: **{_tsb_pid_best['Principal ID']}** (ACT Rp {_tsb_pid_best['ACT']:,.0f}, A/T {_tsb_pid_best['AT']:.1f}%) — jaga ketersediaan dan tingkatkan volume order.")
                if not _tsb_pid_below.empty:
                    for _,_pr in _tsb_pid_below.iterrows():
                        st.markdown(f"- ⚠️ **{_pr['Principal ID']}**: ACT Rp {_pr['ACT']:,.0f} vs TGT Rp {_pr['TGT']:,.0f} (A/T {_pr['AT']:.1f}% · gap Rp {abs(_pr['DEV']):,.0f}) — aktivasi customer dormant dan intensifkan prospek baru.")
                st.markdown("**👤 Fokus per Sales Person Title**")
                if _tsb_title_lines:
                    for _line in _tsb_title_lines: st.markdown(f"- {_line}")
                else:
                    st.markdown("- Data Sales Person Title belum tersedia atau seluruh title telah mencapai target.")

            _tsb_rekom=[]
            if _tsb_days_left>0 and _tsb_gap>0:
                if _tsb_daily_needed>_tsb_daily_actual*1.5:
                    _tsb_rekom.append(f"🔴 **Run rate dibutuhkan Rp {_tsb_daily_needed:,.0f}/hari** jauh di atas laju saat ini — lakukan review harian, percepat pemenuhan PO, serta aktivasi customer potensial CHW/ETW.")
                elif _tsb_daily_needed>_tsb_daily_actual*1.1:
                    _tsb_rekom.append(f"🟠 **Naikkan run rate sekitar {(_tsb_run_rate_gap/max(_tsb_daily_actual,1)*100):.0f}%** — pastikan follow-up Sales Person Title dan jadwal kunjungan berjalan penuh.")
                else:
                    _tsb_rekom.append("🟢 **Gap masih dalam jangkauan** — pertahankan eksekusi harian dan fokus pada upselling serta repeat order.")
            if not _tsb_pid_below.empty:
                _tsb_rekom.append(f"📋 **{len(_tsb_pid_below)} Principal Toserba** masih di bawah target — lakukan check-in khusus dan cek hambatan stok, harga, atau coverage.")
            _tsb_rekom.append("📅 **Action harian**: review gap per Principal dan Sales Person Title di awal hari, lalu konfirmasi order yang masuk sebelum hari berakhir.")
            st.markdown("**⚡ Rekomendasi Taktis**")
            for _line in _tsb_rekom: st.markdown(f"- {_line}")


            st.markdown("---")

            # ── Breakdown per Principal ID ────────────────────────────────────
            st.markdown("##### Breakdown per Principal ID · Bulanan")
            pid_tsb_cols = st.columns(len(TOSERBA_PIDS),gap="small")
            for pcol_tsb, pid_val in zip(pid_tsb_cols, TOSERBA_PIDS):
                with pcol_tsb:
                    pid_data = tsb_month[tsb_month[principal_col].astype(str).str.upper().str.strip()==pid_val]
                    pid_act  = float(pid_data["_perf_sales"].sum())
                    pid_tgt  = float(pid_data["_perf_target"].sum())
                    pid_ach  = pid_act/pid_tgt*100 if pid_tgt else None
                    ach_c    = RED if (pid_ach or 0)<60 else ("#F5C542" if (pid_ach or 0)<90 else "#39C77B")
                    st.markdown(
                        f"<div style='background:{MIDNIGHT_BLUE};border-top:4px solid {ach_c};border-radius:10px;padding:12px 14px'>"
                        f"<div style='color:#FCF1CF;font-weight:900;font-size:1rem'>{pid_val}</div>"
                        f"<div style='color:#B9CADB;font-size:.78rem;margin-top:3px'>Actual: <b style='color:#FCF1CF'>Rp {pid_act:,.0f}</b></div>"
                        f"<div style='color:#B9CADB;font-size:.78rem'>Target: <b style='color:#FCF1CF'>Rp {pid_tgt:,.0f}</b></div>"
                        f"<div style='color:{ach_c};font-weight:800;font-size:1.05rem;margin-top:3px'>{f'{pid_ach:.1f}%' if pid_ach else '—'}</div>"
                        f"<div style='color:#B9CADB;font-size:.75rem'>Dev: Rp {pid_act-pid_tgt:,.0f}</div>"
                        f"</div>",unsafe_allow_html=True)

            st.markdown("---")

            # ── Monitoring by Sales Person Title ─────────────────────────────
            if not sp_col:
                st.warning("Kolom Sales Person Title tidak ditemukan di database.")
            else:
                def _make_sp_tsb_table(data):
                    sp_t = (data.groupby(sp_col,dropna=False)
                            .agg(ACT=("_perf_sales","sum"),TGT=("_perf_target","sum"))
                            .reset_index().rename(columns={sp_col:"Sales Person Title"}))
                    sp_t["A/T"] = sp_t["ACT"].div(sp_t["TGT"].replace(0,float("nan")))*100
                    sp_t["DEV"] = sp_t["ACT"] - sp_t["TGT"]
                    sp_t = sp_t.sort_values("ACT",ascending=False).reset_index(drop=True)
                    total_r = pd.DataFrame([{"Sales Person Title":"TOTAL",
                        "ACT":sp_t["ACT"].sum(),"TGT":sp_t["TGT"].sum(),
                        "A/T":sp_t["ACT"].sum()/sp_t["TGT"].sum()*100 if sp_t["TGT"].sum() else None,
                        "DEV":sp_t["DEV"].sum()}])
                    return pd.concat([sp_t,total_r],ignore_index=True)

                st.markdown(f"##### By Sales Person Title · {month_name} {int(selected_year)}")
                sp_tsb_m = _make_sp_tsb_table(tsb_month)
                st.dataframe(soften_performance_table(sp_tsb_m),width="stretch",hide_index=True,height=38+len(sp_tsb_m)*35)

                st.markdown(f"##### By Sales Person Title · YTD Jan–{month_name} {int(selected_year)}")
                sp_tsb_y = _make_sp_tsb_table(tsb_ytd)
                st.dataframe(soften_performance_table(sp_tsb_y),width="stretch",hide_index=True,height=38+len(sp_tsb_y)*35)

                exp_tsb = sp_tsb_m.copy()
                exp_tsb.insert(0,"Periode",f"{month_name} {int(selected_year)}")
                st.download_button(
                    f"⬇ Export Toserba {month_name} {int(selected_year)} (.csv)",
                    data=exp_tsb.to_csv(index=False).encode("utf-8-sig"),
                    file_name=f"toserba_{month_name}_{int(selected_year)}.csv",
                    mime="text/csv",key="export_toserba",
                )

        else:
            st.markdown(f"#### Customer di Bawah Rata-rata {int(selected_year)} · bulan berjalan {month_name}")
            st.caption("Rata-rata dihitung dari semua bulan sebelum bulan terpilih pada tahun terpilih (bulan berjalan tidak ikut dihitung). Pencapaian membandingkan actual bulan berjalan dengan rata-rata tersebut, growth membandingkan bulan berjalan dengan bulan sebelumnya, dan deviasi = actual bulan berjalan dikurangi rata-rata.")
            cust_c=col(df,"customer")
            if not cust_c:
                st.warning("Kolom Customer Name tidak ditemukan di database.")
            elif int(selected_month)<=1:
                st.info("Belum ada bulan sebelum bulan terpilih pada tahun ini untuk menghitung rata-rata.")
            else:
                n_prior=int(selected_month)-1
                prev_name=MONTHS.get(int(selected_month)-1,str(int(selected_month)-1))
                prior=scoped[scoped["_perf_month"]<int(selected_month)]
                cur=scoped[scoped["_perf_month"]==int(selected_month)]
                prev=scoped[scoped["_perf_month"]==int(selected_month)-1]

                # Normalize customer key agar tidak duplikat akibat spasi/kapitalisasi
                def norm_cust(s): return s.astype(str).str.strip().str.casefold().str.replace(r"\s+"," ",regex=True)

                # Bangun lookup nama asli (ambil modus agar konsisten)
                all_cust=scoped[[cust_c]].copy()
                all_cust["_ck"]=norm_cust(all_cust[cust_c])
                canonical_name=(all_cust.groupby("_ck")[cust_c]
                                .agg(lambda s: s.mode().iat[0] if len(s.mode()) else s.iat[0])
                                .rename("Customer Name"))

                def sales_by_ck(frame):
                    f=frame[[cust_c,"_perf_sales"]].copy()
                    f["_ck"]=norm_cust(f[cust_c])
                    return f.groupby("_ck")["_perf_sales"].sum()

                prior_s=sales_by_ck(prior)
                cur_s  =sales_by_ck(cur)
                prev_s =sales_by_ck(prev)

                tbl=pd.DataFrame({"Average":prior_s/n_prior})
                tbl["cur"] =cur_s .reindex(tbl.index).fillna(0)
                tbl["prev"]=prev_s.reindex(tbl.index).fillna(0)

                if principal_col:
                    tmp=scoped[[cust_c,principal_col]].copy()
                    tmp["_ck"]=norm_cust(tmp[cust_c])
                    tbl["Principal ID"]=(tmp.groupby("_ck")[principal_col]
                        .agg(lambda s: s.mode().iat[0] if len(s.mode()) else (s.dropna().astype(str).iat[0] if len(s.dropna()) else ""))
                        .reindex(tbl.index).fillna(""))
                if sp_col:
                    tmp=scoped[[cust_c,sp_col]].copy()
                    tmp["_ck"]=norm_cust(tmp[cust_c])
                    tbl["Sales Person Title"]=(tmp.groupby("_ck")[sp_col]
                        .agg(lambda s: s.mode().iat[0] if len(s.mode()) else (s.dropna().astype(str).iat[0] if len(s.dropna()) else ""))
                        .reindex(tbl.index).fillna(""))

                tbl=tbl[tbl["Average"]>0]
                below=tbl[tbl["cur"]<tbl["Average"]].copy()
                below["Pencapaian vs Avg %"]=below["cur"]/below["Average"]*100
                below[f"Growth vs {prev_name} %"]=below["cur"]/below["prev"].replace(0,float("nan"))*100-100
                below["Deviasi"]=below["cur"]-below["Average"]
                below=below.sort_values("Deviasi")

                # Gabung nama asli
                below=below.join(canonical_name,how="left")
                below["Customer Name"]=below["Customer Name"].fillna(below.index.to_series().astype(str))

                # ── FILTER Principal ID & Sales Person Title ───────────────────
                below_filter_cols=st.columns(2,gap="small")
                with below_filter_cols[0]:
                    if principal_col and "Principal ID" in below.columns:
                        pid_options=sorted(below["Principal ID"].dropna().astype(str).unique().tolist())
                        sel_pid=st.multiselect("Filter Principal ID",pid_options,key="below_avg_pid",placeholder="Semua Principal ID")
                    else:
                        sel_pid=[]
                with below_filter_cols[1]:
                    if sp_col and "Sales Person Title" in below.columns:
                        sp_options=sorted(below["Sales Person Title"].dropna().astype(str).unique().tolist())
                        sel_sp=st.multiselect("Filter Sales Person Title",sp_options,key="below_avg_sp",placeholder="Semua Sales Person Title")
                    else:
                        sel_sp=[]
                filtered_below=below.copy()
                if sel_pid:
                    filtered_below=filtered_below[filtered_below["Principal ID"].astype(str).isin(sel_pid)]
                if sel_sp:
                    filtered_below=filtered_below[filtered_below["Sales Person Title"].astype(str).isin(sel_sp)]

                c1,c2=st.columns(2,gap="small")
                c1.metric("Customer di bawah rata-rata",f"{len(filtered_below):,}")
                c2.metric("Total deviasi bulan berjalan",f"{filtered_below['Deviasi'].sum():,.0f}")

                final_cols=["Customer Name",f"Average {int(selected_year)}",f"Actual {month_name}",
                            "Pencapaian vs Avg %",f"Growth vs {prev_name} %","Deviasi"]
                disp2=filtered_below.reset_index(drop=True).rename(columns={
                    "Average":f"Average {int(selected_year)}",
                    "cur":f"Actual {month_name}",
                })
                if principal_col: final_cols.append("Principal ID")
                if sp_col: final_cols.append("Sales Person Title")
                disp2=disp2[final_cols]

                col_cfg={
                    "Customer Name":st.column_config.TextColumn("Customer Name"),
                    f"Average {int(selected_year)}":st.column_config.NumberColumn(f"Avg {int(selected_year)}",format="%,.0f"),
                    f"Actual {month_name}":st.column_config.NumberColumn(f"Actual {month_name}",format="%,.0f"),
                    "Pencapaian vs Avg %":st.column_config.NumberColumn("Pencapaian vs Avg",format="%.1f%%"),
                    f"Growth vs {prev_name} %":st.column_config.NumberColumn(f"Growth vs {prev_name}",format="%.1f%%"),
                    "Deviasi":st.column_config.NumberColumn("Deviasi",format="%,.0f"),
                }
                if principal_col: col_cfg["Principal ID"]=st.column_config.TextColumn("Principal ID")
                if sp_col: col_cfg["Sales Person Title"]=st.column_config.TextColumn("Sales Person Title")

                st.download_button(
                    "Unduh CSV daftar di bawah rata-rata",
                    data=disp2.to_csv(index=False).encode("utf-8-sig"),
                    file_name=f"customer_di_bawah_rata_rata_{int(selected_year)}_{month_name}.csv",
                    mime="text/csv",
                    key="below_avg_csv",
                )
                tbl_h=38+len(disp2)*35
                st.dataframe(disp2,width="stretch",hide_index=True,height=tbl_h,column_config=col_cfg)
                if filtered_below.empty:
                    st.success("Tidak ada customer dengan sales bulan berjalan di bawah rata-rata pada filter ini.")
                elif principal_col and "Principal ID" in filtered_below.columns:
                    st.markdown("#### Principal yang masih perlu dikejar")
                    st.caption("Diurutkan dari total deviasi terdalam; customer terkait berada di bawah rata-rata sales 2026 mereka.")
                    adv=filtered_below.groupby("Principal ID").agg(customer=("Deviasi","size"),deviasi=("Deviasi","sum")).sort_values("deviasi").head(5)
                    for pid,row in adv.iterrows():
                        st.markdown(f"- **{pid}**: {int(row['customer']):,} customer di bawah rata-rata, total deviasi {row['deviasi']:,.0f} — prioritaskan kunjungan dan amankan order ulang dari customer terbesar di principal ini.")
    # ── Export PPT Performance Sales ────────────────────────────────────────
    if _PPTX_OK:
        st.divider()
        st.markdown("#### 📊 Export Presentasi")
        if st.button("⬇ Download PPT — Performance Sales",key="ppt_perf",type="primary"):
            pid_c_perf=col(df,"principal"); sp_c_perf=col(df,"salesperson")
            pid_df_ppt=None; sp_df_ppt=None
            if pid_c_perf and "month_df" in dir():
                try:
                    mp=month_df.groupby(principal_col,dropna=False).agg(ACT=("_perf_sales","sum"),TGT=("_perf_target","sum")).reset_index().rename(columns={principal_col:"Principal ID"})
                    mp["A/T"]=mp["ACT"].div(mp["TGT"].replace(0,float("nan")))*100
                    mp["DEV"]=mp["ACT"]-mp["TGT"]
                    mp["ACT"]=mp["ACT"].map("{:,.0f}".format); mp["TGT"]=mp["TGT"].map("{:,.0f}".format)
                    mp["A/T"]=mp["A/T"].map(lambda v:f"{v:.1f}%" if pd.notna(v) else "—")
                    mp["DEV"]=mp["DEV"].map("{:,.0f}".format)
                    pid_df_ppt=mp.sort_values("ACT",ascending=False)
                except Exception: pass
            if sp_c_perf and "month_df" in dir():
                try:
                    ms=month_df.groupby(sp_c_perf,dropna=False).agg(ACT=("_perf_sales","sum"),TGT=("_perf_target","sum")).reset_index().rename(columns={sp_c_perf:"Sales Person Title"})
                    ms["A/T"]=ms["ACT"].div(ms["TGT"].replace(0,float("nan")))*100
                    ms["DEV"]=ms["ACT"]-ms["TGT"]
                    ms["ACT"]=ms["ACT"].map("{:,.0f}".format); ms["TGT"]=ms["TGT"].map("{:,.0f}".format)
                    ms["A/T"]=ms["A/T"].map(lambda v:f"{v:.1f}%" if pd.notna(v) else "—")
                    ms["DEV"]=ms["DEV"].map("{:,.0f}".format)
                    sp_df_ppt=ms.sort_values("ACT",ascending=False)
                except Exception: pass
            perf_ppt_data={
                "subtitle":f"{month_name} {int(selected_year)} · CHP AAM Cabang {ACTIVE_BRANCH_NAME}",
                "month_name":month_name,"year":int(selected_year),
                "m_actual":m_actual or 0,"m_target":m_target or 0,
                "m_ach":m_ach or 0,"m_dev":m_dev or 0,
                "y_actual":y_actual or 0,"y_target":y_target or 0,
                "y_ach":y_ach or 0,"y_dev":y_dev or 0,
                "fot":fot if "fot" in dir() else 0,
                "headline":       headline if "headline" in dir() else "",
                "status_label":   status_label if "status_label" in dir() else "—",
                "rekomendasi":    rekomendasi if "rekomendasi" in dir() else [],
                "sp_insight_lines": sp_insight_lines if "sp_insight_lines" in dir() else [],
                "trend_dir":      trend_dir if "trend_dir" in dir() else "—",
                "best_month_name":best_month_name if "best_month_name" in dir() else "—",
                "monthly_over":   monthly_over if "monthly_over" in dir() else 0,
                "daily_run_rate": daily_run_rate if "daily_run_rate" in dir() else 0,
                "daily_needed":   daily_needed if "daily_needed" in dir() else 0,
                "days_left":      days_left if "days_left" in dir() else 0,
                "gap":            gap if "gap" in dir() else 0,
                "principal_df":pid_df_ppt,"salesperson_df":sp_df_ppt,
            }
            ppt_bytes=build_ppt_performance(perf_ppt_data)
            if ppt_bytes:
                st.download_button("📥 Klik untuk simpan file PPT",data=ppt_bytes,
                    file_name=f"performance_sales_{month_name}_{int(selected_year)}.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    key="ppt_perf_dl")
    save_settings(settings)
elif page=="Summary Program":
    st.markdown(f"<div class='score-heading'>◇ SUMMARY {len(programs)} PROGRAM PUSH</div>",unsafe_allow_html=True)
    st.caption("Target mingguan dihitung dari persentase target bulanan; actual dan AI Insight mengikuti Week acuan yang dipilih pada Konfigurasi Program.")
    filter_values={}
    summary_insight_data=[]
    cards=st.columns(max(1,len(programs)),gap="small")
    for card,program in zip(cards,programs):
        with card:
            selected={"year":[],"month":[],"principal_name":[],"principal":[],
                      "channel_group":[],"business_group":[],"product_code":[],"product":[],
                      "metric":"Total Relasi Beli Transaksi"}
            selected=apply_active_program_config(program,selected)
            filter_values[program]=selected
            scoped,actual,tgt,achievement,deviation,_=calculate_program_kpi(program,selected)
            title_detail=title_performance(scoped,program,selected)
            weekly_by_title,weekly_actual,weekly_count,weekly_label=weekly_program_performance(
                scoped,selected["metric"],selected,calendar_path)
            saved_weekly_target=_cfg_save.get(program,{}).get("weekly_target") if isinstance(_cfg_save,dict) else None
            saved_weekly_percent=_cfg_save.get(program,{}).get("weekly_target_percent") if isinstance(_cfg_save,dict) else None
            if saved_weekly_percent is not None:
                weekly_target_percent=float(saved_weekly_percent or 0)
                weekly_target=float(tgt or 0)*weekly_target_percent/100.0
            elif saved_weekly_target is None:
                selected_year=(selected.get("year") or [None])[0]
                selected_month=(selected.get("month") or [None])[0]
                target_weeks=calendar_weeks_in_period(calendar_path,selected_year,selected_month) if selected_year and selected_month else 0
                weekly_target=round(float(tgt or 0)/target_weeks) if target_weeks else 0.0
                weekly_target_percent=(weekly_target/float(tgt)*100 if tgt and float(tgt)>0 else 0.0)
            else:
                weekly_target=float(saved_weekly_target or 0)
                weekly_target_percent=(weekly_target/float(tgt)*100 if tgt and float(tgt)>0 else 0.0)
            weekly_achievement=(float(weekly_actual)/weekly_target*100
                                if weekly_actual is not None and weekly_target>0 else None)
            weekly_deviation=(float(weekly_actual)-float(weekly_target)
                              if weekly_actual is not None and weekly_target is not None else None)
            # Warna progress mingguan menunjukkan apakah target mingguannya sudah
            # tercapai, terpisah dari warna pencapaian target bulanan.
            weekly_color="#39C77B" if weekly_achievement is not None and weekly_achievement>=100 else RED
            actual_text=fmt_summary(actual,selected["metric"])
            if selected["metric"]!="Total Sales" and actual is not None: actual_text+= " relasi"
            ach_text=f"{achievement:,.1f}%" if achievement is not None else "—"
            if achievement is None or achievement<60:
                ach_color=RED
            elif achievement<90:
                ach_color="#F5C542"
            else:
                ach_color="#39C77B"
            if selected["metric"]=="Total Sales":
                deviation_text=fmt_summary(deviation,selected["metric"])
            else:
                deviation_text=f"{deviation:,.0f} relasi" if deviation is not None else "—"
            if weekly_deviation is None:
                weekly_deviation_text="—"
            elif selected["metric"]=="Total Sales":
                weekly_deviation_text=fmt_summary(weekly_deviation,selected["metric"])
            else:
                weekly_deviation_text=f"{weekly_deviation:+,.0f} relasi"
            insight_parts=summary_weekly_insight(program,selected["metric"],weekly_label,
                                                 weekly_actual,weekly_target,weekly_by_title,title_detail)
            monthly_insight_parts=summary_program_insight(scoped,program,selected,actual,tgt,achievement,deviation,title_detail)
            summary_insight_data.append({
                "program":program,"label":program_label(program),"selected":selected,
                "actual":actual,"target":tgt,"achievement":achievement,"deviation":deviation,
                "weekly_actual":weekly_actual,"weekly_target":weekly_target,
                "weekly_achievement":weekly_achievement,"weekly_label":weekly_label,
                "title_detail":title_detail,"weekly_parts":insight_parts,
                "monthly_parts":monthly_insight_parts,
            })
            if weekly_actual is None:
                insight_level, insight_status = "medium", "ACTUAL MINGGUAN TIDAK TERSEDIA"
            elif weekly_target<=0:
                insight_level, insight_status = "medium", "TARGET MINGGUAN BELUM ADA"
            elif weekly_achievement < 60:
                insight_level, insight_status = "low", "RENDAH"
            elif weekly_achievement < 90:
                insight_level, insight_status = "medium", "SEDANG"
            else:
                insight_level, insight_status = "good", "BAGUS"
            def safe(value): return html.escape(str(value))
            insight_html="<br>".join(f"<span class='negative-contribution'>{safe(text)}</span>" if is_negative else (f"<span class='on-track-contribution'>{safe(text)}</span>" if text.startswith("On track:") else safe(text)) for text,is_negative in insight_parts)
            active_scope=[]
            for _kind,_label in [("year","Year"),("month","Bulan"),("principal","Principal ID"),
                                 ("product_code","Product Code"),("product","Produk"),
                                 ("channel_group","Channel Group"),("business_group","Bisnis Group")]:
                _values=selected.get(_kind,[])
                if _values:
                    active_scope.append(f"{_label}: {', '.join(map(str,_values)) if isinstance(_values,(list,tuple)) else _values}")
            active_scope.append(_metric_labels.get(selected["metric"],selected["metric"]))
            scope_note=" · ".join(active_scope)
            card_html=f"""<div class='scorecard'>
<div class='score-program'>{safe(PROGRAM_NUMBERS.get(program,'PROGRAM'))}</div>
<div class='score-name'>{safe(program_label(program))}</div>
<div class='score-filter'>{safe(scope_note)}</div>
<div class='score-row'>Target Bulanan: <b>{safe(fmt_summary(tgt,selected['metric']))}</b></div>
<div class='score-row'>Actual Bulanan: <b>{safe(actual_text)}</b></div>
<div class='score-row'>Deviasi Bulanan: <b>{safe(deviation_text)}</b></div>
<div class='score-ach'>Pencapaian Bulanan: <strong style='color:{ach_color}'>{safe(ach_text)}</strong></div>
<div class='score-row'>Target Mingguan: <b>{safe(fmt_summary(weekly_target,selected['metric']))}</b></div>
<div class='score-row'>Actual {safe(weekly_label or 'Minggu berjalan')}: <b>{safe(fmt_summary(weekly_actual,selected['metric']))}</b></div>
<div class='score-row'>Deviasi Mingguan: <b style='color:{weekly_color}'>{safe(weekly_deviation_text)}</b></div>
<div class='score-row'>A/T Mingguan: <b>{safe(f'{weekly_achievement:,.1f}%' if weekly_achievement is not None else '—')}</b></div>
<div class='weekly-plot'><div class='weekly-plot-label'>Progres target mingguan</div><div class='weekly-plot-track'><div class='weekly-plot-fill' style='width:{min(100.0,max(0.0,float(weekly_achievement or 0))):.1f}%;background:{weekly_color}'></div></div><div class='weekly-plot-value' style='color:{weekly_color}'>{safe(f'{weekly_achievement:,.1f}%' if weekly_achievement is not None else '—')}</div></div>
<div class='score-insight insight-{insight_level}'><b>AI Insight:</b> <span class='insight-status insight-status-{insight_level}'>{safe(insight_status)}</span><br>{insight_html}</div>
</div>"""
            st.markdown(card_html,unsafe_allow_html=True)
    if not programs: st.info("Pilih minimal satu program dari panel kiri.")
    elif summary_insight_data:
        st.markdown("#### 🧠 AI Insight · Strategi On Target Program Push")
        _targeted=[item for item in summary_insight_data if item["target"] is not None and float(item["target"] or 0)>0]
        _worst=min(_targeted,key=lambda item:(item["achievement"] if item["achievement"] is not None else -1)) if _targeted else None
        _safe_insight=lambda value:html.escape(str(value))
        if _worst is None:
            _banner_icon="🟡"; _banner_label="TARGET PROGRAM BELUM LENGKAP"; _banner_color="#F5C542"; _banner_bg="#3b2e00"
            _banner_text="Lengkapi target bulanan pada Konfigurasi Program agar prioritas gap dan strategi closing bisa dihitung."
        else:
            _worst_achievement=float(_worst["achievement"] or 0)
            if _worst_achievement<60:
                _banner_icon="🔴"; _banner_label="KRITIS — INTERVENSI SEGERA"; _banner_color="#C8102E"; _banner_bg="#3b0000"
            elif _worst_achievement<90:
                _banner_icon="🟠"; _banner_label="PERLU AKSELERASI"; _banner_color="#FF9B35"; _banner_bg="#3b1a00"
            else:
                _banner_icon="🟢"; _banner_label="ON TRACK — JAGA MOMENTUM"; _banner_color="#39C77B"; _banner_bg="#0d3b1e"
            _metric=_worst["selected"]["metric"]
            _gap=max(0.0,float(_worst["target"] or 0)-float(_worst["actual"] or 0))
            _period_month=(_worst["selected"].get("month") or [""])[0]
            _period_year=(_worst["selected"].get("year") or [""])[0]
            _month_no=next((num for num,name in MONTHS.items() if key(name)==key(_period_month)),None)
            _days_remaining=0
            if _month_no and str(_period_year).isdigit():
                _period_start=pd.Timestamp(int(_period_year),int(_month_no),1)
                _period_end=_period_start+pd.offsets.MonthEnd(0)
                _now=pd.Timestamp.now().normalize()
                if _period_start<=_now<=_period_end:_days_remaining=int((_period_end-_now).days)
            _banner_text=(f"{_worst['label']} paling perlu perhatian: pencapaian {_worst_achievement:,.1f}% dengan gap "
                          f"{fmt_summary(_gap,_metric)}. Fokuskan follow-up pada title dengan gap terbesar dan pantau actual mingguan.")
            if _days_remaining>0:
                _banner_text+=f" Sisa {_days_remaining} hari; kebutuhan rata-rata {_safe_insight(fmt_summary(_gap/_days_remaining,_metric))}/hari."
        st.markdown(f"""
<div style='background:{_banner_bg};border-left:6px solid {_banner_color};border-radius:12px;padding:17px 22px;margin:10px 0 18px'>
<div style='color:{_banner_color};font-weight:900;font-size:.9rem;letter-spacing:.07em'>{_banner_icon} {_banner_label}</div>
<div style='color:#F0F4FF;font-size:1.02rem;font-weight:650;margin-top:7px;line-height:1.5'>{_safe_insight(_banner_text)}</div>
</div>""",unsafe_allow_html=True)
        _insight_left,_insight_right=st.columns(2,gap="large")
        with _insight_left:
            st.markdown("**📊 Angka Kunci Program**")
            for _item in summary_insight_data:
                _metric=_item["selected"]["metric"]
                _att="—" if _item["achievement"] is None else f"{_item['achievement']:,.1f}%"
                _monthly_gap=max(0.0,float(_item["target"] or 0)-float(_item["actual"] or 0)) if _item["target"] is not None else 0
                _weekly_att="—" if _item["weekly_achievement"] is None else f"{_item['weekly_achievement']:,.1f}%"
                _weekly_gap=max(0.0,float(_item["weekly_target"] or 0)-float(_item["weekly_actual"] or 0)) if _item["weekly_actual"] is not None else 0
                _line=(f"**{_safe_insight(_item['label'])}** — Bulanan {_safe_insight(fmt_summary(_item['actual'],_metric))} / "
                       f"{_safe_insight(fmt_summary(_item['target'],_metric))} ({_att}); gap {_safe_insight(fmt_summary(_monthly_gap,_metric))}. "
                       f"{_safe_insight(_item['weekly_label'] or 'Mingguan')}: {_safe_insight(fmt_summary(_item['weekly_actual'],_metric))} / "
                       f"{_safe_insight(fmt_summary(_item['weekly_target'],_metric))} ({_weekly_att}); gap {_safe_insight(fmt_summary(_weekly_gap,_metric))}.")
                st.markdown(f"- {_line}")
            st.markdown("**📈 Prioritas Minggu Ini**")
            for _item in summary_insight_data:
                _urgent=[text for text,is_negative in _item["weekly_parts"] if is_negative]
                if not _urgent:
                    _urgent=[text for text,is_negative in _item["weekly_parts"][:1]]
                for _text in _urgent[:2]:
                    st.markdown(f"- **{_safe_insight(_item['label'])}:** {_safe_insight(_text)}")
        with _insight_right:
            st.markdown("**👤 Fokus per Sales Person Title**")
            _focus_count=0
            for _item in summary_insight_data:
                _detail=_item["title_detail"]
                if _detail is None or _detail.empty:continue
                _targets=pd.to_numeric(_detail["Target"],errors="coerce").fillna(0)
                _actuals=pd.to_numeric(_detail["Actual"],errors="coerce").fillna(0)
                _below=_detail.loc[(_targets>0)&(_actuals<_targets)].copy()
                if _below.empty:continue
                _below["_gap"]=_targets.loc[_below.index]-_actuals.loc[_below.index]
                _below=_below.sort_values("_gap",ascending=False).head(2)
                _focus_count+=len(_below)
                _focus_text=", ".join(f"{_safe_insight(row['Sales Person Title'])} (gap {_safe_insight(fmt_summary(row['_gap'],_item['selected']['metric']))})" for _,row in _below.iterrows())
                st.markdown(f"- **{_safe_insight(_item['label'])}:** {_focus_text} — follow-up relasi yang belum transaksi dan minta komitmen order/closing.")
            if not _focus_count:
                st.markdown("- Semua Sales Person Title bertarget sudah on track; jaga ritme order dan pastikan transaksi tercatat.")
            st.markdown("**⚡ Rekomendasi Taktis**")
            _on_track=[]
            for _item in summary_insight_data:
                _metric=_item["selected"]["metric"]
                _monthly_parts=_item["monthly_parts"]
                _negative=[text for text,is_negative in _monthly_parts if is_negative]
                if _item["achievement"] is not None and _item["achievement"]>=90:
                    _on_track.append(_item["label"])
                if _negative:
                    st.markdown(f"- **{_safe_insight(_item['label'])}:** {_safe_insight(_negative[-1])}")
                elif _item["achievement"] is not None and _item["achievement"]>=100:
                    st.markdown(f"- **{_safe_insight(_item['label'])}:** Target tercapai; amankan repeat order dan pastikan seluruh PO terealisasi.")
            if _on_track:
                st.markdown(f"- **Sudah on track:** {_safe_insight(', '.join(_on_track))}. Pertahankan aktivitas dan jangan lepaskan peluang closing.")
    settings.setdefault("filters",{}).update(filter_values)
    settings["programs"]=programs
    settings["dashboard_page"]=page
    # ── Export PPT Summary Program ───────────────────────────────────────────
    if _PPTX_OK and programs:
        st.divider()
        st.markdown("#### 📊 Export Presentasi")
        if st.button("⬇ Download PPT — Summary Program",key="ppt_summary",type="primary"):
            _prog_data_list=[]
            for _prog in programs:
                _psel=filter_values[_prog]
                _pscoped,_pact,_ptgt,_pach,_pdev,_=calculate_program_kpi(_prog,_psel)
                _ptd=title_performance(_pscoped,_prog,_psel)
                _metric=_psel["metric"]
                _pact_txt=fmt(_pact,_metric)
                if _metric!="Total Sales" and _pact is not None: _pact_txt+=f" relasi"
                _ptgt_txt=fmt(_ptgt,_metric)
                _pdev_txt=fmt(_pdev,_metric) if _metric=="Total Sales" else (f"{_pdev:,.0f} relasi" if _pdev is not None else "—")
                _status="BAGUS" if (_pach or 0)>=90 else ("SEDANG" if (_pach or 0)>=60 else "RENDAH")
                _prog_data_list.append({
                    "label":program_label(_prog),
                    "actual_text":_pact_txt,"target_text":_ptgt_txt,
                    "achievement":_pach,"deviation_text":_pdev_txt,
                    "status":_status,"title_df":_ptd,
                    "insight_parts":[(t,n) for t,n in summary_program_insight(_pscoped,_prog,_psel,_pact,_ptgt,_pach,_pdev,_ptd)],
                    "insight_status":("BAGUS" if (_pach or 0)>=90 else ("SEDANG" if (_pach or 0)>=60 else "RENDAH")),
                    "insight_level": ("good" if (_pach or 0)>=90 else ("medium" if (_pach or 0)>=60 else "low")),
                })
            _month_str=filter_values[programs[0]]["month"] if programs else ""
            _year_str =filter_values[programs[0]]["year"] if programs else ""
            _period=f"{_month_str} {_year_str}".strip() or "CHP AAM"
            _ppt=build_ppt_summary_program(_prog_data_list,f"{_period} · CHP AAM Cabang {ACTIVE_BRANCH_NAME}")
            if _ppt:
                st.download_button("📥 Klik untuk simpan file PPT",data=_ppt,
                    file_name=f"summary_program_{_period.replace(' ','_')}.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    key="ppt_summary_dl")
    save_settings(settings)

elif page=="Program by Sales Person Title":
    st.subheader("2. Program by Sales Person Title")
    if programs:
        program=st.selectbox("Pilih program",programs,key="title_program",format_func=program_label)
        # Program configuration for the active month is the single source of
        # truth for this view. Keep only the program picker here so its scope
        # cannot drift from the settings used by Program Push Summary.
        selected=apply_active_program_config(program,{
            "year":[],"month":[],"principal_name":[],"principal":[],
            "channel_group":[],"business_group":[],"product_code":[],"product":[],
            "metric":"Total Relasi Beli Transaksi",
        })
        scoped,actual,tgt,achievement,deviation,_=calculate_program_kpi(program,selected)
        by_title=title_performance(scoped,program,selected)
        if by_title.empty:
            st.info("Kolom Sales Person Title belum tersedia pada database.")
        else:
            weekly_by_title,weekly_actual_program,weekly_count,weekly_label=weekly_program_performance(scoped,selected["metric"],selected,calendar_path)
            _active_idx=int(st.session_state.get("push_month_tab_idx",pd.Timestamp.now().month-1))
            _active_idx=min(max(_active_idx,0),11)
            _active_config_month=list(MONTHS.values())[_active_idx]
            _weekly_config=settings.get("program_config_by_month",{}).get(_active_config_month,{}).get(program,{})
            total_target=float(tgt or 0)
            weekly_target_value=_weekly_config.get("weekly_target")
            if weekly_target_value is None:
                _selected_year=(selected.get("year") or [None])[0]
                _selected_month=(selected.get("month") or [None])[0]
                _week_total=calendar_weeks_in_period(calendar_path,_selected_year,_selected_month) if _selected_year and _selected_month else 0
                weekly_target_value=round(total_target/_week_total) if _week_total else 0.0
            weekly_target_value=float(weekly_target_value or 0)
            weekly_target_percent=_weekly_config.get("weekly_target_percent")
            if weekly_target_percent is not None:
                weekly_ratio=max(0.0,float(weekly_target_percent or 0))/100.0
            else:
                # Legacy configs only kept the program weekly amount. Convert it
                # to a percentage using the configured target total, then apply
                # that percentage separately to each Sales Person Title target.
                _configured_target_total=float(_weekly_config.get("target_total") or 0)
                _weekly_denominator=_configured_target_total if _configured_target_total>0 else total_target
                weekly_ratio=weekly_target_value/_weekly_denominator if _weekly_denominator>0 else 0.0
            by_title["TGT WEEKLY"]=pd.to_numeric(by_title["Target"],errors="coerce").fillna(0)*weekly_ratio
            by_title["DEV WEEKLY"]=by_title.apply(
                lambda row:math.ceil(weekly_by_title.get(key(row.get("Sales Person Title")),0)-row["TGT WEEKLY"]) if pd.notna(row["TGT WEEKLY"]) else None,
                axis=1,
            )
            date_col=col(scoped,"date")
            available_dates=pd.to_datetime(scoped[date_col],errors="coerce").dropna() if date_col else []
            daily_table=append_daily_additions(scoped,by_title,selected["metric"]) if len(available_dates) else by_title.copy()
            title_minimized=False
            st.markdown(f"#### PROGRAM PUSH · {program_label(program)}")
            title_minimized=st.toggle("Minimize keterangan & pengaturan kolom",value=True,key="title_minimize_v2_"+key(program).replace(" ","_"),help="Sembunyikan keterangan penambahan harian / target weekly, panel Show/Hide Columns, dan tombol CSV agar tabel tampil lebih ringkas.")
            if not title_minimized:
                if len(available_dates):
                    st.caption("Penambahan harian = kumulatif hari H dikurangi H−1.")
                else:
                    st.info("Kolom tanggal tidak ditemukan; penambahan harian belum dapat dihitung.")
                if weekly_count:
                    st.caption(f"TGT WEEKLY tiap Sales Person Title = target bulanan per title × {weekly_ratio*100:.1f}% dari Konfigurasi Program; DEV WEEKLY = actual {weekly_label} − target weekly per title.")
                else:
                    st.warning("Jumlah minggu kalender pada periode filter tidak tersedia; TGT WEEKLY mengikuti nilai yang disimpan pada Konfigurasi Program.")
            daily_table=daily_table.rename(columns={"Sales Person Title":"Salesperson_Title","Actual":"ACT","Target":"TGT","A/T (%)":"A/T","Deviation":"DEV"})
            has_actual=pd.to_numeric(daily_table["ACT"],errors="coerce").fillna(0).ne(0)
            has_target=pd.to_numeric(daily_table["TGT"],errors="coerce").fillna(0).ne(0)
            daily_table=daily_table[has_actual|has_target].reset_index(drop=True)
            total_actual=0 if actual is None else actual
            total_target=0 if tgt is None else tgt
            total_weekly_target=float(pd.to_numeric(by_title["TGT WEEKLY"],errors="coerce").fillna(0).sum())
            total_weekly_actual=weekly_actual_program
            total_row={column:0 for column in daily_table.columns}
            total_row.update({
                "Salesperson_Title":"TOTAL PROGRAM",
                "ACT":total_actual,
                "TGT":total_target,
                "A/T":total_actual/total_target*100 if total_target else None,
                "DEV":total_actual-total_target,
                "TGT WEEKLY":total_weekly_target,
                "DEV WEEKLY":math.ceil(total_weekly_actual-total_weekly_target) if total_weekly_actual is not None and total_weekly_target is not None else None,
            })
            for daily_col in daily_table.columns[7:]:
                total_row[daily_col]=pd.to_numeric(daily_table[daily_col],errors="coerce").fillna(0).sum()
            daily_table=pd.concat([daily_table,pd.DataFrame([total_row])],ignore_index=True)
            all_columns=list(daily_table.columns)
            visible_store=settings.setdefault("title_visible_columns",{})
            known_store=settings.setdefault("title_known_columns",{})
            previously_known=known_store.get(program,[])
            saved_visible=visible_store.get(program)
            visible_key="title_visible_columns_"+key(program).replace(" ","_")
            visible_state_key="title_visible_columns_state_"+key(program).replace(" ","_")
            # Keep the selected columns outside the widget state. Streamlit
            # clears widget keys whenever the multiselect is hidden by the
            # minimize toggle or the page changes.
            if visible_state_key in st.session_state:
                saved_visible=st.session_state[visible_state_key]
            elif saved_visible is None and visible_key in st.session_state:
                saved_visible=st.session_state[visible_key]
            if saved_visible is None:
                visible_columns=list(all_columns)
            else:
                visible_columns=[column for column in all_columns if column in saved_visible or column not in previously_known]
            if not visible_columns:
                visible_columns=list(all_columns)
            st.session_state[visible_state_key]=visible_columns
            if not title_minimized:
                if visible_key not in st.session_state:
                    st.session_state[visible_key]=list(visible_columns)
                with st.expander("Atur Show / Hide Columns"):
                    st.multiselect("Kolom yang ditampilkan",options=all_columns,key=visible_key,help="Pilihan kolom disimpan per program dan akan digunakan kembali saat dashboard dibuka.")
                selected_columns=[column for column in all_columns if column in st.session_state.get(visible_key,visible_columns)]
                if selected_columns:
                    visible_columns=selected_columns
                    st.session_state[visible_state_key]=list(visible_columns)
            total_mask=daily_table["Salesperson_Title"].astype(str).eq("TOTAL PROGRAM")
            body_rows=daily_table[~total_mask].copy()
            total_rows=daily_table[total_mask]
            if "ACT" in body_rows:
                body_rows=body_rows.sort_values("ACT",ascending=False,na_position="last")
            display_table=pd.concat([body_rows,total_rows],ignore_index=True).loc[:,[column for column in all_columns if column in visible_columns]]
            if not title_minimized:
                st.download_button("Unduh CSV tabel ini",data=display_table.to_csv(index=False).encode("utf-8-sig"),file_name=f"program_{key(program)}_by_sales_person_title.csv",mime="text/csv",key="title_csv_"+key(program).replace(" ","_"))
            _push_column_config={}
            for _column in display_table.columns:
                if _column=="Salesperson_Title":
                    _push_column_config[_column]=st.column_config.TextColumn("Sales Person Title")
                elif pd.api.types.is_numeric_dtype(display_table[_column]):
                    _format="%.1f%%" if _column=="A/T" else "%,.1f" if _column=="TGT WEEKLY" else "%,.0f"
                    _push_column_config[_column]=st.column_config.NumberColumn(_column,format=_format)
            st.dataframe(display_table,width="stretch",hide_index=True,
                         height=38+len(display_table)*35,
                         column_config=_push_column_config)
            visible_store[program]=visible_columns
            known_store[program]=list(dict.fromkeys(previously_known+all_columns))
            st.bar_chart(by_title.set_index("Sales Person Title")[["Actual","Target"]])
        settings.setdefault("filters",{})[program]=selected
        settings["dashboard_page"]=page
        settings["programs"]=programs
        # ── Export PPT Program by Sales Person Title ─────────────────────────
        if _PPTX_OK:
            st.divider()
            st.markdown("#### 📊 Export Presentasi")
            if st.button("⬇ Download PPT — Program by Sales Person Title",key="ppt_bytitle",type="primary"):
                _prog_insight=[latest_daily_progress(scoped,selected["metric"])] if scoped is not None else []
                _prog_rek=[f"Actual: {fmt(actual,selected['metric'])} | Target: {fmt(tgt,selected['metric'])} | A/T: {f'{achievement:.1f}%' if achievement else '—'} | Deviasi: {fmt(deviation,selected['metric'])}"]
                _ppt=build_ppt_program_by_title(
                    program_label(program), by_title, selected,
                    actual, tgt, achievement, deviation, selected["metric"],
                    insight_lines=_prog_insight,
                    rekomendasi_lines=_prog_rek,
                    weekly_label=weekly_label if "weekly_label" in dir() else "")
                if _ppt:
                    st.download_button("📥 Klik untuk simpan file PPT",data=_ppt,
                        file_name=f"program_{key(program).replace(' ','_')}_by_title.pptx",
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                        key="ppt_bytitle_dl")
        save_settings(settings)
    else: st.info("Pilih program dari panel kiri untuk melihat capaian per Sales Person Title.")

elif page=="Sasaran Program":
    st.subheader("3. Sasaran Program")
    st.caption("Rencana kunjungan disusun dari Sales Person Title → relasi → status transaksi pada setiap program. Outlet yang belum transaksi diprioritaskan untuk kunjungan.")

    visit_filter_cols=st.columns(3)
    with visit_filter_cols[0]:
        visit_years=filter_select(df,"year","Year kunjungan","visit_year_filter")
    with visit_filter_cols[1]:
        visit_months=filter_select(df,"month","Month kunjungan","visit_month_filter")
    title_col=col(df,"salesperson")
    customer_col=col(df,"customer")
    sales_col=col(df,"sales")
    if not title_col or not customer_col:
        st.warning("Database perlu memiliki Sales Person Title dan Customer Name agar sasaran kunjungan dapat dirinci per relasi.")
        st.stop()
    with visit_filter_cols[2]:
        visit_titles=filter_select(df,"salesperson","Sales Person Title","visit_salesperson_filter")
    if not visit_titles:
        st.info("Pilih Sales Person Title terlebih dahulu untuk menampilkan sasaran dan rencana kunjungannya.")
        st.stop()
    st.caption("Atur basis potensi per program. Tahunan mencari pembeli tahun sebelumnya yang belum membeli tahun sasaran; bulanan mencari pembeli pada bulan tepat sebelumnya yang belum membeli di bulan sasaran.")
    st.caption("Kriteria populasi khusus: RPB LUMA mengambil relasi yang pernah bertransaksi pada Principal ID MCS; RPB NOZA mengambil 50% relasi unik yang belum transaksi dari Channel Group Pharmacy, diprioritaskan berdasarkan nilai transaksi historis.")
    cadence_settings=settings.setdefault("visit_cadence",{})
    cadence_defaults={prog:("Bulanan" if PROGRAM_FILTER_DEFAULTS.get(prog,{}).get("month","All") not in (None,"All","") else "Tahunan") for prog in PROGRAMS}
    cadence_cols=st.columns(max(1,len(programs)))
    for box,prog in zip(cadence_cols,programs):
        default_cadence=cadence_settings.get(prog,cadence_defaults[prog])
        with box:
            cadence_settings[prog]=st.selectbox(f"Basis potensi · {program_label(prog)}",["Bulanan","Tahunan"],index=0 if default_cadence=="Bulanan" else 1,key=f"visit_cadence_{key(prog).replace(' ','_')}")
    settings["visit_cadence"]=cadence_settings
    if not title_col or not customer_col:
        st.warning("Database perlu memiliki Sales Person Title dan Customer Name agar sasaran kunjungan dapat dirinci per relasi.")
    elif not programs:
        st.info("Pilih minimal satu program Push di panel kiri untuk melihat status transaksi outlet per program.")
    else:
        in_scope_by_program={}
        bought_customers_by_program={}
        current_pairs_by_program={}
        current_customers_by_title={}
        program_status_columns=[]
        def selected_values(value):
            if value in (None,"All","ALL",""): return []
            return list(value) if isinstance(value,(list,tuple,set)) else [value]
        def relation_key(title,relasi):
            return (" ".join(str(title).strip().casefold().split()),
                    " ".join(str(relasi).strip().casefold().split()))
        def narrowed_period(configured,visit_selection):
            configured_values=selected_values(configured)
            visit_values=selected_values(visit_selection)
            if configured_values and visit_values:
                common=[value for value in configured_values if str(value) in {str(item) for item in visit_values}]
                return common if common else ["__NO_MATCH__"]
            return configured_values or visit_values
        all_visit_pairs=set()
        all_years=filter_options(df,"year")
        all_months=filter_options(df,"month")
        # The special LUMA/NOZA populations use all historical transactions.
        # Build this positive-sales view once per page run, not once per program.
        _pool_cols=list(dict.fromkeys(c for c in [col(df,"sales"),col(df,"salesperson"),col(df,"customer"),col(df,"principal"),col(df,"channel_group")] if c))
        _pool_cache=st.session_state.setdefault("_sasaran_positive_pool_cache",{})
        if _pool_cache.get("frame_id")!=id(df):
            _pool_cache.clear()
            _pool_cache["frame_id"]=id(df)
        _pool_cache_key=tuple(_pool_cols)
        if _pool_cache_key not in _pool_cache:
            _pool_cache[_pool_cache_key]=positive_sales_rows(df[_pool_cols]) if _pool_cols else None
        pool_positive=_pool_cache[_pool_cache_key]
        if pool_positive is not None and not pool_positive.empty:
            _pool_title_col=col(pool_positive,"salesperson")
            if _pool_title_col:
                pool_positive=pool_positive[pool_positive[_pool_title_col].astype(str).str.strip().isin(visit_titles)]
        def month_number(value):
            try: return int(float(str(value).strip()))
            except (ValueError,TypeError):
                lowered=str(value).strip().casefold()
                return next((number for number,name in MONTHS.items() if name.casefold()==lowered),None)
        def month_name(number): return MONTHS.get(int(number),str(number))
        for prog in programs:
            saved_filter={**PROGRAM_FILTER_DEFAULTS.get(prog,{}),**settings.get("filters",{}).get(prog,{})}
            prog_filter={kind:saved_filter.get(kind,"All") for kind in ["year","month","principal_name","principal","channel_group","product_code","product"]}
            status_col=f"{PROGRAM_NUMBERS.get(prog,'PROGRAM')} · {program_label(prog)}"
            program_status_columns.append(status_col)
            year_values=narrowed_period(prog_filter.get("year"),visit_years)
            year_conflict=year_values==["__NO_MATCH__"]
            if year_conflict: year_values=["__NO_MATCH__"]
            prog_filter["year"]=year_values
            years=[] if year_conflict else (year_values or ([max(all_years,key=lambda value:int(float(value)))] if all_years else []))
            cadence=cadence_settings.get(prog,cadence_defaults[prog])
            month_values=narrowed_period(prog_filter.get("month"),visit_months)
            month_conflict=month_values==["__NO_MATCH__"]
            if month_conflict: month_values=["__NO_MATCH__"]
            prog_filter["month"]=month_values
            if cadence=="Bulanan" and not month_values and not month_conflict:
                observed_months=[m for m in all_months if month_number(m) is not None]
                month_values=[max(observed_months,key=month_number)] if observed_months else []
            historical_scopes=[]
            for year_value in years:
                try: year_number=int(float(str(year_value)))
                except (ValueError,TypeError): continue
                if cadence=="Tahunan":
                    history_filter={**prog_filter,"year":[str(year_number-1)],"month":[]}
                    historical_scopes.append(filtered_data(history_filter))
                else:
                    valid_months=sorted({month_number(value) for value in month_values if month_number(value) is not None})
                    for month_value in valid_months:
                        previous_month=12 if month_value==1 else month_value-1
                        previous_year=str(year_number-1) if month_value==1 else year_value
                        history_filter={**prog_filter,"year":[previous_year],"month":[month_name(previous_month)]}
                        historical_scopes.append(filtered_data(history_filter))
            current_scope=filtered_data(prog_filter)
            history_scope=pd.concat(historical_scopes,ignore_index=True) if historical_scopes else df.iloc[0:0].copy()
            p_title=col(current_scope,"salesperson"); p_customer=col(current_scope,"customer")
            history_title=col(history_scope,"salesperson"); history_customer=col(history_scope,"customer")
            if p_title:
                current_scope=current_scope[current_scope[p_title].astype(str).str.strip().isin(visit_titles)]
            if history_title:
                history_scope=history_scope[history_scope[history_title].astype(str).str.strip().isin(visit_titles)]
            current_positive=positive_sales_rows(current_scope)
            history_positive=positive_sales_rows(history_scope)
            current_pairs=set(); history_pairs=set()
            current_customers=set()
            current_customers_by_title[prog]={}
            if current_positive is not None and not current_positive.empty and p_customer:
                current_customers=set(current_positive[p_customer].dropna().astype(str).str.strip())
                if p_title:
                    positive_with_titles=current_positive.dropna(subset=[p_title,p_customer])
                    current_pairs=set(zip(positive_with_titles[p_title].astype(str).str.strip(),positive_with_titles[p_customer].astype(str).str.strip()))
                    current_customers_by_title[prog]={title:set(group[p_customer].dropna().astype(str).str.strip()) for title,group in positive_with_titles.groupby(p_title)}
            if history_positive is not None and not history_positive.empty and history_title and history_customer:
                history_pairs=set(zip(history_positive[history_title].astype(str).str.strip(),history_positive[history_customer].astype(str).str.strip()))
            # Status per baris sasaran harus mengikuti title + relasi seperti
            # perhitungan Program by Sales Person Title, bukan status relasi
            # global yang bisa menandai title lain sebagai sudah transaksi.
            current_pair_keys={relation_key(title,customer) for title,customer in current_pairs}
            candidates={pair for pair in history_pairs if relation_key(*pair) not in current_pair_keys}
            # Program khusus memakai populasi sasaran yang ditetapkan bisnis:
            # RPB LUMA dari outlet yang pernah bertransaksi dengan Principal
            # ID MCS; RPB NOZA hanya outlet Channel Group Pharmacy (ketentuan BPOM).
            # Populasi diambil lintas periode agar relasi tetap muncul walau
            # belum memiliki riwayat produk/program pada periode berjalan.
            if pool_positive is not None and not pool_positive.empty:
                pool_title=col(pool_positive,"salesperson")
                pool_customer=col(pool_positive,"customer")
                if pool_title and pool_customer:
                    if prog=="PB LUMA":
                        principal_col=col(pool_positive,"principal")
                        if principal_col:
                            pool=pool_positive[pool_positive[principal_col].astype("string").str.strip().str.casefold().eq("mcs")]
                            candidates = {
                                (str(title).strip(),str(customer).strip())
                                for title,customer in zip(pool[pool_title],pool[pool_customer])
                                if pd.notna(title) and pd.notna(customer)
                                and relation_key(title,customer) not in current_pair_keys
                            }
                    elif prog=="PB NOZA":
                        channel_col=col(pool_positive,"channel_group")
                        if channel_col:
                            pool=pool_positive[pool_positive[channel_col].astype("string").str.strip().str.casefold().eq("pharmacy")]
                            current_relation_keys={relation_key("",customer)[1] for customer in current_customers}
                            candidates = {
                                (str(title).strip(),str(customer).strip())
                                for title,customer in zip(pool[pool_title],pool[pool_customer])
                                if pd.notna(title) and pd.notna(customer)
                                and relation_key("",customer)[1] not in current_relation_keys
                            }
                            # Ambil separuh dari jumlah relasi unik yang belum
                            # transaksi, bukan separuh nilai target program.
                            eligible_relations={relation_key("",customer)[1] for _,customer in candidates}
                            relation_limit=(len(eligible_relations)+1)//2
                            pharmacy_sales_col=col(pool,"sales")
                            if pharmacy_sales_col and pool_customer:
                                pool=pool.copy()
                                pool[pharmacy_sales_col]=pd.to_numeric(pool[pharmacy_sales_col],errors="coerce").fillna(0)
                                pool["_relation_key"]=pool[pool_customer].astype("string").str.strip().str.casefold().str.replace(r"\s+"," ",regex=True)
                                ranked_customers=(pool[pool["_relation_key"].isin(eligible_relations)]
                                    .groupby("_relation_key",dropna=True)[pharmacy_sales_col].sum()
                                    .sort_values(ascending=False,kind="stable"))
                                selected_relations=set(ranked_customers.head(relation_limit).index.astype(str))
                                candidates={pair for pair in candidates if relation_key("",pair[1])[1] in selected_relations}
            bought_customers_by_program[prog]=current_customers
            current_pairs_by_program[prog]=current_pair_keys
            in_scope_by_program[prog]={relation_key(*pair) for pair in current_pairs|candidates}
            for pairs in (current_pairs,candidates):
                all_visit_pairs.update(pairs)
        visit_table=(pd.DataFrame.from_records(
                         sorted(all_visit_pairs,key=lambda pair:(str(pair[0]),str(pair[1]))),
                         columns=["Sales Person Title","Relasi"])
                     if all_visit_pairs else pd.DataFrame(columns=["Sales Person Title","Relasi"]))
        if visit_titles:
            visit_table=visit_table[visit_table["Sales Person Title"].astype(str).isin(visit_titles)].copy()
        if visit_table.empty:
            st.info("Tidak ada relasi yang termasuk dalam filter program tersimpan untuk periode dan Sales Person Title yang dipilih.")
        else:
            for prog,status_col in zip(programs,program_status_columns):
                in_scope=in_scope_by_program[prog]
                visit_table[status_col]=[
                    ("✓ Sudah" if relation_key(title,relasi) in current_pairs_by_program[prog] else "○ Belum")
                    if relation_key(title,relasi) in in_scope else "— N/A"
                    for title,relasi in zip(visit_table["Sales Person Title"],visit_table["Relasi"])
                ]
            _pending_matrix=visit_table[program_status_columns].eq("○ Belum")
            visit_table["Program belum transaksi"] = _pending_matrix.sum(axis=1)
            _program_labels=[program_label(prog) for prog in programs]
            _pending_values=_pending_matrix.to_numpy(dtype=bool)
            visit_table["Fokus kunjungan"] = [
                ", ".join(label for label,is_pending in zip(_program_labels,row) if is_pending)
                or "Tidak ada potensi tertinggal dalam cakupan program"
                for row in _pending_values
            ]
            outlets_pending=visit_table.loc[visit_table["Program belum transaksi"].gt(0),"Relasi"].nunique()
            relations_in_view=visit_table["Relasi"].nunique()
            visit_table=visit_table.sort_values(["Program belum transaksi","Sales Person Title","Relasi"],ascending=[False,True,True]).drop(columns="Program belum transaksi").reset_index(drop=True)
            st.caption(f"{relations_in_view:,} relasi unik dalam cakupan · {outlets_pending:,} relasi memiliki minimal satu program yang belum transaksi. Actual status memakai scope tanggal dan filter yang sama dengan Summary Program; relasi di luar cakupan ditandai N/A.")
            st.markdown("##### Potensi transaksi per program")
            potential_rows=[]
            for prog,status_col in zip(programs,program_status_columns):
                if visit_titles:
                    done=len(set().union(*(current_customers_by_title.get(prog,{}).get(title,set()) for title in visit_titles)))
                else:
                    done=len(bought_customers_by_program.get(prog,set()))
                pending=visit_table.loc[visit_table[status_col].eq("○ Belum"),"Relasi"].nunique()
                potential_rows.append({"Program":program_label(prog),"Sudah transaksi":done,"Belum transaksi":pending})
            potential=pd.DataFrame(potential_rows)
            st.bar_chart(potential,x="Program",y=["Sudah transaksi","Belum transaksi"],color=["#39C77B","#F5C542"],horizontal=True,stack=True,height=min(420,max(220,48*len(potential)+60)))
            st.markdown("##### Ringkasan potensi belum transaksi per program")
            potential_summary=[]
            pending_rows_by_program={status_col:visit_table.loc[visit_table[status_col].eq("○ Belum"),["Sales Person Title","Relasi"]]
                                     for status_col in program_status_columns}
            for prog,status_col in zip(programs,program_status_columns):
                pending_rows=pending_rows_by_program[status_col]
                by_title=(pending_rows.groupby("Sales Person Title",dropna=False)
                          .agg(**{"Relasi potensial":("Relasi","nunique")})
                          .reset_index().sort_values(["Relasi potensial","Sales Person Title"],ascending=[False,True]))
                top_title="—"
                if not by_title.empty:
                    top=by_title.iloc[0]
                    top_title=f"{top['Sales Person Title']} ({int(top['Relasi potensial'])})"
                potential_summary.append({
                    "Program":program_label(prog),
                    "Relasi belum transaksi":int(pending_rows["Relasi"].nunique()),
                    "Sales Person Title terdampak":int(pending_rows["Sales Person Title"].nunique()),
                    "Prioritas utama":top_title,
                })
            potential_summary_df=pd.DataFrame(potential_summary)
            st.dataframe(potential_summary_df,width="stretch",hide_index=True)
            st.download_button(
                "⬇ Export Ringkasan Potensi (.csv)",
                data=potential_summary_df.to_csv(index=False).encode("utf-8-sig"),
                file_name=f"ringkasan_potensi_program.csv",
                mime="text/csv",
                key="export_potential_summary",
            )
            st.markdown("##### Detail belum transaksi per Sales Person Title")
            all_detail_rows=[]
            for prog,status_col in zip(programs,program_status_columns):
                pending_rows=pending_rows_by_program[status_col]
                by_title=(pending_rows.groupby("Sales Person Title",dropna=False)
                          .agg(**{"Relasi potensial":("Relasi","nunique")})
                          .reset_index().sort_values(["Relasi potensial","Sales Person Title"],ascending=[False,True]))
                by_title.insert(0,"Program",program_label(prog))
                all_detail_rows.append(by_title)
            if all_detail_rows:
                detail_df=pd.concat(all_detail_rows,ignore_index=True)
                st.dataframe(detail_df,width="stretch",hide_index=True)
                st.download_button(
                    "⬇ Export Detail per Salesperson (.csv)",
                    data=detail_df.to_csv(index=False).encode("utf-8-sig"),
                    file_name=f"detail_belum_transaksi_per_salesperson.csv",
                    mime="text/csv",
                    key="export_detail_salesperson",
                )
            # ── Export PPT Sasaran Program ────────────────────────────────────
            if _PPTX_OK:
                if st.button("⬇ Download PPT — Sasaran Program",key="ppt_sasaran",type="primary"):
                    _ps_df=potential_summary_df if "potential_summary_df" in dir() else None
                    _det_df=detail_df if "detail_df" in dir() else None
                    # Bangun insight lines dari data potensi
                    _ai_ins=[]; _ai_rek=[]
                    if _ps_df is not None and not _ps_df.empty:
                        _ai_ins.append(f"Total relasi belum transaksi: {_ps_df['Relasi belum transaksi'].sum():,} relasi")
                        for _,r in _ps_df.iterrows():
                            _ai_ins.append(f"  • {r['Program']}: {r['Relasi belum transaksi']} relasi belum transaksi, prioritas: {r.get('Prioritas utama','—')}")
                        _ai_rek.append(f"Fokuskan kunjungan ke relasi dengan status '○ Belum' — konversi ke transaksi sebelum akhir bulan.")
                        _ai_rek.append(f"Distribusikan beban kunjungan merata ke semua Sales Person Title yang masih memiliki relasi potensial.")
                    _ppt=build_ppt_sasaran_program(programs,program_label,_ps_df,_det_df,visit_table,
                                                    ai_insight_lines=_ai_ins,rekomendasi_lines=_ai_rek)
                    if _ppt:
                        st.download_button("📥 Klik untuk simpan file PPT",data=_ppt,
                            file_name="sasaran_program.pptx",
                            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                            key="ppt_sasaran_dl")
            program_cell_colors={
                "RB CHD":"#F4D7DC","NOO PDB":"#DCE6F7","PB NOZA":"#D8F0E8","PB HYC":"#E9DDF4","PB LUMA":"#FBE8D3"
            }
            visit_styler=style_unified_table(visit_table)
            for prog,status_col in zip(programs,program_status_columns):
                visit_styler=visit_styler.set_properties(subset=[status_col],**{"background-color":program_cell_colors.get(prog,"#E8EAF0"),"color":"#0C1D42","font-weight":"bold","text-align":"center"})
            visit_styler=visit_styler.set_properties(subset=["Fokus kunjungan"],**{"color":"#C8102E","font-weight":"bold"})
            # Native Streamlit grid avoids building a huge HTML string and DOM
            # for every relation when the transaction workbook has many rows.
            st.dataframe(visit_styler,width="stretch",hide_index=True)
elif page=="Sasaran Kunjungan Harian per Sales Person Title":
    st.subheader("4. Sasaran Kunjungan Harian per Sales Person Title")
    st.caption("Pilih Sales Person Title dan tanggal; status transaksi mengikuti filter konfigurasi masing-masing program push.")
    daily_filter_cols=st.columns(3)
    with daily_filter_cols[0]:
        visit_years=filter_select(df,"year","Year transaksi","daily_visit_year_filter")
    with daily_filter_cols[1]:
        visit_months=filter_select(df,"month","Month transaksi","daily_visit_month_filter")
    with daily_filter_cols[2]:
        visit_titles=filter_select(df,"salesperson","Sales Person Title","daily_visit_title_filter")
    if not visit_titles:
        visit_titles=filter_options(df,"salesperson")
        st.caption("Belum memilih Sales Person Title: semua salesperson ditampilkan. Gunakan filter di atas untuk mempersempit daftar.")
    def _daily_selected_values(value):
        if value in (None,"All","ALL",""): return []
        return list(value) if isinstance(value,(list,tuple,set)) else [value]
    def _daily_narrow(configured,selected):
        configured_values=_daily_selected_values(configured)
        selected_values=_daily_selected_values(selected)
        if configured_values and selected_values:
            common=[v for v in configured_values if str(v) in {str(x) for x in selected_values}]
            return common if common else ["__NO_MATCH__"]
        return configured_values or selected_values
    _daily_title_col=col(df,"salesperson")
    _daily_customer_col=col(df,"customer")
    if not _daily_title_col or not _daily_customer_col:
        st.warning("Database perlu memiliki Sales Person Title dan Customer Name untuk mencocokkan status transaksi.")
        st.stop()
    current_pairs_by_program={}
    for _prog in programs:
        _saved={**PROGRAM_FILTER_DEFAULTS.get(_prog,{}),**settings.get("filters",{}).get(_prog,{})}
        _filter={kind:_saved.get(kind,"All") for kind in ["year","month","principal_name","principal","channel_group","product_code","product"]}
        _filter["year"]=_daily_narrow(_filter.get("year"),visit_years)
        _filter["month"]=_daily_narrow(_filter.get("month"),visit_months)
        _scope=filtered_data(_filter)
        if _daily_title_col in _scope.columns:
            _scope=_scope[_scope[_daily_title_col].astype(str).str.strip().isin(visit_titles)]
        _positive=positive_sales_rows(_scope)
        if _positive is not None and not _positive.empty:
            _valid=_positive.dropna(subset=[_daily_title_col,_daily_customer_col])
            current_pairs_by_program[_prog]={
                (" ".join(str(t).strip().casefold().split())," ".join(str(c).strip().casefold().split()))
                for t,c in zip(_valid[_daily_title_col],_valid[_daily_customer_col])
            }
        else:
            current_pairs_by_program[_prog]=set()
    def relation_key(title,relasi):
        return (" ".join(str(title).strip().casefold().split())," ".join(str(relasi).strip().casefold().split()))
    calendar_data=read_file(calendar_path) if calendar_path.exists() else None
    fixed_call_data=read_file(fixed_call_path) if fixed_call_path.exists() else None
    def workbook_col(frame,name):
        return next((column for column in frame.columns if key(column)==key(name)),None) if frame is not None else None
    if calendar_data is None or fixed_call_data is None:
        st.warning("Overview Calendar dan Overview Fixed Call harus tersedia. Unggah keduanya melalui panel kiri untuk menampilkan sasaran harian.")
    else:
        calendar_date_col=workbook_col(calendar_data,"DATE")
        parsed_dates=pd.to_datetime(calendar_data[calendar_date_col],errors="coerce") if calendar_date_col else pd.Series(dtype="datetime64[ns]")
        running_date=pd.Timestamp.now().normalize()
        date_options=sorted(pd.Timestamp(value) for value in parsed_dates.dropna().dt.normalize().unique()) if len(parsed_dates) else []
        if not date_options:
            st.warning("Kolom DATE pada Overview Calendar tidak memiliki tanggal yang dapat dibaca.")
        else:
            default_date=running_date if running_date in date_options else date_options[-1]
            calendar_day_col=workbook_col(calendar_data,"DAY")
            calendar_week_col=workbook_col(calendar_data,"WEEK")
            calendar_year_col=workbook_col(calendar_data,"YEARID")
            calendar_month_col=workbook_col(calendar_data,"MONTHID")
            fc_year=workbook_col(fixed_call_data,"YEARID")
            fc_week=workbook_col(fixed_call_data,"WEEK")
            fc_day=workbook_col(fixed_call_data,"DAYS")
            fc_quarter=workbook_col(fixed_call_data,"PERIODE")
            fc_branch=workbook_col(fixed_call_data,"BRACH_CODE") or workbook_col(fixed_call_data,"BRANCH_CODE")
            fc_title=workbook_col(fixed_call_data,"POSITION_DESC")
            fc_customer=workbook_col(fixed_call_data,"CUSTOMER_NAME")
            fc_frequency=workbook_col(fixed_call_data,"FREQUENCY")
            if all([calendar_day_col,calendar_week_col,calendar_year_col,calendar_month_col,fc_year,fc_week,fc_day]):
                scheduled_keys=set()
                _schedule_columns=[fc_year,fc_week,fc_day]
                if fc_quarter: _schedule_columns.append(fc_quarter)
                if fc_branch: _schedule_columns.append(fc_branch)
                _schedule_rows=fixed_call_data[_schedule_columns].dropna(subset=[fc_year,fc_week,fc_day]).drop_duplicates()
                for _,_schedule_row in _schedule_rows.iterrows():
                    try: _schedule_year=int(float(_schedule_row[fc_year])); _schedule_week=int(float(_schedule_row[fc_week]))
                    except (TypeError,ValueError): continue
                    _schedule_day=" ".join(str(_schedule_row[fc_day]).strip().upper().split())
                    _schedule_quarter=str(_schedule_row[fc_quarter]).strip().upper() if fc_quarter else ""
                    _schedule_branch=str(_schedule_row[fc_branch]).strip().upper() if fc_branch else ACTIVE_BRANCH_CODE
                    if _schedule_branch==ACTIVE_BRANCH_CODE:
                        scheduled_keys.add((_schedule_year,_schedule_week,_schedule_day,_schedule_quarter))
                _scheduled_dates=[]
                for _date_value in date_options:
                    _calendar_rows=calendar_data.loc[parsed_dates.dt.normalize()==_date_value]
                    if _calendar_rows.empty: continue
                    _calendar_date_row=_calendar_rows.iloc[0]
                    try:
                        _calendar_year=int(float(_calendar_date_row[calendar_year_col]))
                        _calendar_week=int(float(_calendar_date_row[calendar_week_col]))
                        _calendar_month=int(float(_calendar_date_row[calendar_month_col]))
                    except (TypeError,ValueError): continue
                    _calendar_day=" ".join(str(_calendar_date_row[calendar_day_col]).strip().upper().split())
                    _calendar_quarter=f"Q{(_calendar_month-1)//3+1}" if fc_quarter else ""
                    if (_calendar_year,_calendar_week,_calendar_day,_calendar_quarter) in scheduled_keys:
                        _scheduled_dates.append(_date_value)
                if _scheduled_dates and default_date not in _scheduled_dates:
                    _upcoming_dates=[_date_value for _date_value in _scheduled_dates if _date_value>=running_date]
                    default_date=_upcoming_dates[0] if _upcoming_dates else _scheduled_dates[-1]
                    _saved_daily_date=st.session_state.get("daily_visit_date")
                    if _saved_daily_date is not None and pd.Timestamp(_saved_daily_date).normalize() not in _scheduled_dates:
                        st.session_state["daily_visit_date"]=default_date
            filter_cols=st.columns(2)
            with filter_cols[1]:
                selected_date=st.selectbox("Tanggal",date_options,index=date_options.index(default_date),format_func=lambda value:pd.Timestamp(value).strftime("%d %b %Y"),key="daily_visit_date")
            calendar_row=calendar_data.loc[parsed_dates.dt.normalize()==pd.Timestamp(selected_date)].iloc[0]
            day_options=sorted(calendar_data[calendar_day_col].dropna().astype(str).str.strip().str.upper().unique()) if calendar_day_col else []
            week_values=pd.to_numeric(calendar_data[calendar_week_col],errors="coerce").dropna().astype(int).unique().tolist() if calendar_week_col else []
            week_options=sorted(week_values)
            calendar_day=str(calendar_row[calendar_day_col]).strip().upper() if calendar_day_col else ""
            calendar_week=pd.to_numeric(pd.Series([calendar_row[calendar_week_col]]),errors="coerce").iloc[0] if calendar_week_col else None
            with filter_cols[0]:
                title_options=sorted(fixed_call_data[fc_title].dropna().astype(str).str.strip().unique()) if fc_title else []
                if visit_titles:
                    title_options=[value for value in title_options if value in visit_titles]
                selected_daily_titles=st.multiselect("Salesperson",title_options,key="daily_visit_salespersons")
            with filter_cols[0]:
                day_value=st.selectbox("Hari",day_options,index=day_options.index(calendar_day) if calendar_day in day_options else 0,key="daily_visit_day") if day_options else calendar_day
            with filter_cols[1]:
                week_value=st.selectbox("Week",week_options,index=week_options.index(int(calendar_week)) if pd.notna(calendar_week) and int(calendar_week) in week_options else 0,key="daily_visit_week") if week_options else calendar_week
            year_value=pd.to_numeric(pd.Series([calendar_row[calendar_year_col]]),errors="coerce").iloc[0] if calendar_year_col else running_date.year
            month_value=pd.to_numeric(pd.Series([calendar_row[calendar_month_col]]),errors="coerce").iloc[0] if calendar_month_col else running_date.month
            quarter=f"Q{(int(month_value)-1)//3+1}" if pd.notna(month_value) else ""
            if not all([fc_year,fc_week,fc_day,fc_title,fc_customer]):
                st.warning("Kolom YEARID, WEEK, DAYS, POSITION_DESC, dan CUSTOMER_NAME wajib tersedia pada Overview Fixed Call.")
            else:
                daily=fixed_call_data.copy()
                daily=daily[pd.to_numeric(daily[fc_year],errors="coerce").eq(int(year_value))]
                daily=daily[pd.to_numeric(daily[fc_week],errors="coerce").eq(float(week_value))]
                daily=daily[daily[fc_day].astype(str).str.strip().str.upper().eq(day_value)]
                if fc_quarter:
                    daily=daily[daily[fc_quarter].astype(str).str.strip().str.upper().eq(quarter)]
                if fc_branch:
                    daily=daily[daily[fc_branch].astype(str).str.strip().str.upper().eq(ACTIVE_BRANCH_CODE)]
                if selected_daily_titles:
                    daily=daily[daily[fc_title].astype(str).str.strip().isin(selected_daily_titles)]
                if visit_titles:
                    daily=daily[daily[fc_title].astype(str).str.strip().isin(visit_titles)]
                daily=daily.dropna(subset=[fc_title,fc_customer]).drop_duplicates(subset=[fc_title,fc_customer]).copy()
                if daily.empty:
                    st.info(f"Tidak ada rute Fixed Call untuk {day_value.title()}, week {week_value}, {pd.Timestamp(selected_date):%d %b %Y}.")
                else:
                    history_title_col=col(df,"salesperson")
                    history_customer_col=col(df,"customer")
                    _daily_pairs={relation_key(row[fc_title],row[fc_customer]) for _,row in daily.iterrows()}
                    _wanted_titles={pair[0] for pair in _daily_pairs}
                    _wanted_customers={pair[1] for pair in _daily_pairs}
                    _source_title_keys=df[history_title_col].astype("string").str.strip().str.casefold().str.replace(r"\s+"," ",regex=True)
                    _source_customer_keys=df[history_customer_col].astype("string").str.strip().str.casefold().str.replace(r"\s+"," ",regex=True)
                    _pair_mask=_source_title_keys.isin(_wanted_titles)&_source_customer_keys.isin(_wanted_customers)
                    _active_config_month=MONTHS.get(int(month_value),str(month_value)) if pd.notna(month_value) else list(MONTHS.values())[pd.Timestamp(selected_date).month-1]
                    historical_pairs_by_program={}
                    _transaction_date_col=col(df,"date")
                    _period_start=pd.Timestamp(int(year_value),int(month_value),1) if pd.notna(year_value) and pd.notna(month_value) else pd.Timestamp(selected_date).to_period("M").start_time
                    for _prog in programs:
                        _program_config=apply_active_program_config(_prog,{},config_month=_active_config_month)
                        _current_scope=filtered_data(_program_config)
                        _current_scope=_current_scope.loc[_pair_mask.reindex(_current_scope.index,fill_value=False)]
                        _current_positive=positive_sales_rows(_current_scope)
                        if _current_positive is not None and not _current_positive.empty:
                            _current_valid=_current_positive.dropna(subset=[history_title_col,history_customer_col])
                            current_pairs_by_program[_prog]={relation_key(title,customer) for title,customer in zip(_current_valid[history_title_col],_current_valid[history_customer_col])}
                        else:
                            current_pairs_by_program[_prog]=set()
                        _history_config=dict(_program_config)
                        _history_config["year"]=[]; _history_config["month"]=[]
                        _history_scope=filtered_data(_history_config)
                        _history_scope=_history_scope.loc[_pair_mask.reindex(_history_scope.index,fill_value=False)]
                        if _transaction_date_col and _transaction_date_col in _history_scope.columns:
                            _history_dates=pd.to_datetime(_history_scope[_transaction_date_col],errors="coerce")
                            _history_scope=_history_scope.loc[_history_dates<_period_start]
                        _history_positive=positive_sales_rows(_history_scope)
                        if _history_positive is not None and not _history_positive.empty:
                            _history_valid=_history_positive.dropna(subset=[history_title_col,history_customer_col])
                            historical_pairs_by_program[_prog]={relation_key(title,customer) for title,customer in zip(_history_valid[history_title_col],_history_valid[history_customer_col])}
                        else:
                            historical_pairs_by_program[_prog]=set()
                    bought_keys_by_program=current_pairs_by_program
                    daily_rows=[]
                    for _,fixed_row in daily.iterrows():
                        title=str(fixed_row[fc_title]).strip()
                        relation=str(fixed_row[fc_customer]).strip()
                        pair_key=relation_key(title,relation)
                        row={"Sales Person Title":title,"Relasi":relation}
                        # Kelompokkan seluruh kolom riwayat berdampingan setelah
                        # Relasi; status transaksi periode berjalan berada di ujung.
                        for prog in programs:
                            _compact_program=program_label(prog).replace("RB ","").replace("SALES ","")
                            row[f"H · {_compact_program}"]="✓" if pair_key in historical_pairs_by_program.get(prog,set()) else ""
                        for prog in programs:
                            _compact_program=program_label(prog).replace("RB ","").replace("SALES ","")
                            row[f"ACT · {_compact_program}"]="✓" if pair_key in bought_keys_by_program.get(prog,set()) else ""
                        daily_rows.append(row)
                    bought_keys_by_program=current_pairs_by_program
                    daily_targets=pd.DataFrame(daily_rows)
                    # Pulihkan reason tersimpan untuk tanggal, salesperson, dan
                    # relasi yang sama agar catatan bisa dilanjutkan kapan saja.
                    existing_visit_log=(pd.read_csv(VISIT_LOG_PATH,dtype=str).fillna("")
                                        if VISIT_LOG_PATH.exists() else pd.DataFrame())
                    visit_date_key=pd.Timestamp(selected_date).strftime("%Y-%m-%d")
                    saved_reason_by_pair={}
                    if not existing_visit_log.empty and {"Tanggal","Sales Person Title","Relasi","Reason"}.issubset(existing_visit_log.columns):
                        saved_for_date=existing_visit_log[existing_visit_log["Tanggal"].astype(str).eq(visit_date_key)]
                        saved_reason_by_pair={
                            relation_key(row["Sales Person Title"],row["Relasi"]):str(row["Reason"])
                            for _,row in saved_for_date.iterrows()
                        }
                    daily_targets["Reason"]=[saved_reason_by_pair.get(relation_key(title,relasi),"") for title,relasi in zip(daily_targets["Sales Person Title"],daily_targets["Relasi"])]
                    st.caption(f"{pd.Timestamp(selected_date):%d %B %Y} · {day_value.title()} · Week {int(week_value)} · {len(daily_targets):,} relasi rute. H = history sebelum bulan berjalan; ACT = transaksi bulan berjalan.")
                    # Bangun column_config: kolom program diberi warna via help text,
                    # kolom Reason bisa diedit bebas
                    prog_labels=[program_label(p) for p in programs]
                    _history_labels=[f"H · {program_label(prog).replace('RB ','').replace('SALES ','')}" for prog in programs]
                    _actual_labels=[f"ACT · {program_label(prog).replace('RB ','').replace('SALES ','')}" for prog in programs]
                    col_cfg={
                        "Sales Person Title":st.column_config.TextColumn("Sales Person Title",disabled=True,width="medium"),
                        "Relasi":st.column_config.TextColumn("Relasi",disabled=True,width="medium"),
                        **{_history_labels[i]:st.column_config.TextColumn(_history_labels[i],disabled=True,width="small",help=f"✓ = ada transaksi positif sebelumnya untuk {program_label(prog)}.") for i,prog in enumerate(programs)},
                        **{_actual_labels[i]:st.column_config.TextColumn(_actual_labels[i],disabled=True,width="small",help=f"✓ = ada transaksi pada bulan berjalan untuk {program_label(prog)}.") for i,prog in enumerate(programs)},
                        "Reason":st.column_config.TextColumn("Reason",help="Isi catatan / reason hasil kunjungan one-on-one dengan salesman",width="medium"),
                    }
                    daily_status_styler=daily_targets.style
                    if _actual_labels:
                        daily_status_styler=daily_status_styler.set_properties(
                            subset=_actual_labels,
                            **{
                                "border-bottom": "2px solid #17324D",
                                "border-right": "1px solid #17324D",
                                "text-align": "center",
                            },
                        )
                        daily_status_styler=daily_status_styler.set_properties(
                            subset=[_actual_labels[0]],
                            **{"border-left": "3px solid #17324D"},
                        )
                    if _history_labels:
                        daily_status_styler=daily_status_styler.map(
                            lambda value:"background-color: #D8F0E3; color: #14532D; font-weight: bold;" if str(value).strip()=="✓" else "",
                            subset=_history_labels,
                        )
                    if _actual_labels:
                        daily_status_styler=daily_status_styler.map(
                            lambda value: (
                                "background-color: #D8F0E3; color: #14532D; font-weight: bold;"
                                if str(value).strip()=="✓"
                                else "background-color: #FCE2E2; color: #9F1239; font-weight: bold;"
                            ),
                            subset=_actual_labels,
                        )
                    edited_daily=st.data_editor(
                        daily_status_styler,
                        width="stretch",
                        # Tampilkan seluruh rute tanpa scroll internal di dalam tabel.
                        height=max(120, 38 + len(daily_targets) * 35),
                        hide_index=True,
                        column_config=col_cfg,
                        key=f"daily_editor_{selected_date}_{day_value}_{week_value}",
                        num_rows="fixed",
                    )
                    # Simpan otomatis setelah setiap edit Reason. Baris kosong
                    # menghapus catatan lama untuk relasi itu; catatan lain aman.
                    save_rows=edited_daily.copy()
                    save_rows.insert(0,"Tanggal",visit_date_key)
                    save_rows.insert(1,"Hari",day_value.title())
                    save_rows.insert(2,"Week",int(week_value))
                    existing=existing_visit_log.copy()
                    if existing.empty:
                        existing=pd.DataFrame(columns=save_rows.columns)
                    all_visit_columns=list(dict.fromkeys([*existing.columns,*save_rows.columns]))
                    existing=existing.reindex(columns=all_visit_columns,fill_value="").fillna("").astype(str)
                    save_rows=save_rows.reindex(columns=all_visit_columns,fill_value="").fillna("").astype(str)
                    visit_keys={tuple(row) for row in save_rows[["Tanggal","Sales Person Title","Relasi"]].itertuples(index=False,name=None)}
                    existing_keys=existing[["Tanggal","Sales Person Title","Relasi"]].apply(tuple,axis=1)
                    existing=existing.loc[~existing_keys.isin(visit_keys)]
                    reasons_to_save=save_rows[save_rows["Reason"].astype(str).str.strip().ne("")]
                    combined=pd.concat([existing,reasons_to_save],ignore_index=True).reindex(columns=all_visit_columns).fillna("").astype(str)
                    if not combined.equals(existing_visit_log.reindex(columns=all_visit_columns,fill_value="").fillna("").astype(str)):
                        _visit_csv=combined.to_csv(index=False).encode("utf-8-sig")
                        VISIT_LOG_PATH.write_bytes(_visit_csv)
                        if st.secrets.get("gcp_service_account"):
                            try: _drive_upload_bytes(VISIT_LOG_PATH.name,_visit_csv,"text/csv","visit_logs",ACTIVE_BRANCH_CODE)
                            except Exception as exc: st.error(f"Log kunjungan tersimpan sementara, tetapi gagal disinkronkan ke Google Drive: {exc}")
                    # Export — sertakan tanggal & info sesi di nama file
                    export_label=f"{pd.Timestamp(selected_date):%Y%m%d}_{day_value}_W{int(week_value)}"
                    st.download_button(
                        label="⬇ Export Sasaran Harian + Reason (.csv)",
                        data=edited_daily.to_csv(index=False).encode("utf-8-sig"),
                        file_name=f"sasaran_harian_{export_label}.csv",
                        mime="text/csv",
                        key=f"export_daily_{export_label}",
                    )

                    st.caption("Reason tersimpan otomatis setiap kali diedit dan akan muncul kembali saat tanggal serta relasi yang sama dibuka.")

                    # ── RIWAYAT KUNJUNGAN (DATABASE MONITORING) ────────────────────
                    st.markdown("---")
                    st.markdown("### 📋 Riwayat Kunjungan — Database Monitoring")
                    st.caption("Rekap Reason yang tersimpan otomatis. Gunakan filter untuk melihat salesperson atau tanggal tertentu.")
                    if not VISIT_LOG_PATH.exists():
                        st.info("Belum ada Reason yang tersimpan. Isi kolom Reason pada tabel sasaran harian; perubahan akan tersimpan otomatis.")
                    else:
                        visit_log=pd.read_csv(VISIT_LOG_PATH,dtype=str).fillna("")
                        if visit_log.empty:
                            st.info("Database kunjungan masih kosong.")
                        else:
                            # Filter
                            log_filter_cols=st.columns(3)
                            all_log_dates=sorted(visit_log["Tanggal"].unique().tolist(),reverse=True) if "Tanggal" in visit_log.columns else []
                            all_log_titles=sorted(visit_log["Sales Person Title"].unique().tolist()) if "Sales Person Title" in visit_log.columns else []
                            with log_filter_cols[0]:
                                log_filter_title=st.multiselect("Filter Salesperson",all_log_titles,key="log_filter_title")
                            with log_filter_cols[1]:
                                log_filter_dates=st.multiselect("Filter Tanggal",all_log_dates,key="log_filter_dates")
                            with log_filter_cols[2]:
                                log_filter_reason=st.checkbox("Tampilkan hanya yang sudah ada Reason",key="log_filter_reason")
                            filtered_log=visit_log.copy()
                            if log_filter_title:
                                filtered_log=filtered_log[filtered_log["Sales Person Title"].isin(log_filter_title)]
                            if log_filter_dates:
                                filtered_log=filtered_log[filtered_log["Tanggal"].isin(log_filter_dates)]
                            if log_filter_reason and "Reason" in filtered_log.columns:
                                filtered_log=filtered_log[filtered_log["Reason"].str.strip().ne("")]
                            # Urutkan terbaru di atas
                            if "Tanggal" in filtered_log.columns:
                                filtered_log=filtered_log.sort_values("Tanggal",ascending=False).reset_index(drop=True)
                            # Ringkasan metrik
                            mc1,mc2,mc3=st.columns(3)
                            mc1.metric("Total entri tersimpan",f"{len(visit_log):,}")
                            mc2.metric("Entri ditampilkan",f"{len(filtered_log):,}")
                            has_reason=filtered_log["Reason"].str.strip().ne("").sum() if "Reason" in filtered_log.columns else 0
                            mc3.metric("Sudah ada Reason",f"{has_reason:,}")
                            # Tabel riwayat — bisa sort semua kolom
                            st.dataframe(filtered_log,width="stretch",hide_index=True)
                            # Export
                            dl1,dl2=st.columns(2)
                            with dl1:
                                st.download_button(
                                    label="⬇ Export Riwayat Kunjungan (.csv)",
                                    data=filtered_log.to_csv(index=False).encode("utf-8-sig"),
                                    file_name="riwayat_kunjungan_all.csv",
                                    mime="text/csv",
                                    key="export_visit_log",
                                )
                            with dl2:
                                # Tombol hapus semua log
                                if st.button("🗑 Hapus Semua Data Riwayat",key="delete_visit_log",type="secondary"):
                                    VISIT_LOG_PATH.unlink(missing_ok=True)
                                    if st.secrets.get("gcp_service_account"):
                                        try: _drive_delete_file(VISIT_LOG_PATH.name,"visit_logs",ACTIVE_BRANCH_CODE)
                                        except Exception as exc: st.error(f"Log lokal terhapus, tetapi gagal dihapus dari Google Drive: {exc}")
                                    st.warning("Database riwayat kunjungan telah dihapus.")
                                    st.rerun()


st.divider()
st.caption("Pilihan program, filter, halaman dashboard, dan sasaran disimpan di Google Drive per cabang. Insight dihitung lokal dari actual, target, gap, dan detail Sales Person Title.")
st.caption("RB menghitung Customer Name unik dengan Sales Value Gross > 0 pada bulan terpilih; tanpa kolom pelanggan, aplikasi menghitung baris transaksi positif. Total Sales menjumlahkan Sales Value Gross dari seluruh transaksi yang cocok dengan konfigurasi terpilih. Target default scorecard mengikuti template dan dapat diedit di submenu Sasaran Program.")

GEMINI_MODELS=["gemini-3.8-flash"]
GEMINI_SYSTEM=("Anda adalah analis penjualan pada dashboard CHP AAM Cabang Bandung. "
"Jawab hanya berbasis data konteks yang diberikan, dalam bahasa Indonesia yang ringkas dan konkret. "
"Sebutkan angka penting (actual, target, A/T, deviasi) seperlunya, tunjukkan area terlemah, "
"dan berikan rekomendasi aksi yang praktis per Sales Person Title, Principal, atau program bila relevan. "
"Jika data konteks tidak cukup untuk menjawab, katakan bagian dashboard mana yang perlu dibuka pengguna.")

def fmt_ai(value):
    return "—" if value is None or (isinstance(value,float) and pd.isna(value)) else f"{value:,.0f}"

def build_ai_context():
    lines=[f"Tampilan aktif: {main_section}"+(f" · {page}" if page else "")]
    if main_section=="Performance Sales":
        pf=settings.get("performance_filters",{})
        year=pf.get("year"); month=pf.get("month"); bgs=pf.get("business_group",[])
        date_c=col(df,"date"); sales_c=col(df,"sales"); target_c=col(df,"target_value"); bg_c=col(df,"business_group"); sp_c=col(df,"salesperson")
        dt=pd.to_datetime(df[date_c],errors="coerce")
        mask=(dt.dt.year==year).fillna(False)
        if bgs and bg_c: mask=mask & df[bg_c].astype(str).isin(bgs)
        tmp=pd.DataFrame({"_m":dt.dt.month,"_s":pd.to_numeric(df[sales_c],errors="coerce").fillna(0),
                           "_t":pd.to_numeric(df[target_c],errors="coerce").fillna(0),
                           "_sp":df[sp_c].astype(str) if sp_c else ""})[mask.to_numpy()]
        m=tmp[tmp["_m"]==month]; y=tmp[tmp["_m"]<=month]
        lines.append(f"Performance Sales · {MONTHS.get(int(month),month)} {int(year)}: Actual {fmt_ai(m['_s'].sum())}, Target {fmt_ai(m['_t'].sum())}, A/T {(m['_s'].sum()/m['_t'].sum()*100) if m['_t'].sum() else 0:.1f}%")
        lines.append(f"YTD s.d. {MONTHS.get(int(month),month)}: Actual {fmt_ai(y['_s'].sum())}, Target {fmt_ai(y['_t'].sum())}, A/T {(y['_s'].sum()/y['_t'].sum()*100) if y['_t'].sum() else 0:.1f}%")
        if sp_c:
            g=m.groupby("_sp",dropna=False)[["_s","_t"]].sum()
            g=g[g["_s"]+g["_t"]>0].sort_values("_s",ascending=False).head(10)
            lines.append("Actual & target bulanan per Sales Person Title (10 terbesar):")
            lines+=[f"- {title}: Actual {fmt_ai(row['_s'])}, Target {fmt_ai(row['_t'])}, A/T {(row['_s']/row['_t']*100) if row['_t'] else 0:.1f}%" for title,row in g.iterrows()]
            gap=m.groupby("_sp",dropna=False)[["_s","_t"]].sum()
            gap=gap[gap["_t"]>0].assign(_dev=lambda d:d["_s"]-d["_t"]).sort_values("_dev").head(5)
            lines.append("Gap bulanan terbesar (Actual - Target):")
            lines+=[f"- {title}: deviasi {fmt_ai(row['_dev'])}" for title,row in gap.iterrows()]
    else:
        for program in programs:
            selected=settings.get("filters",{}).get(program,{})
            if not selected: continue
            scoped,actual,tgt,achievement,deviation,_=calculate_program_kpi(program,selected)
            lines.append(f"{PROGRAM_NUMBERS.get(program,'')} {program_label(program)} · metric {selected.get('metric','')}: Actual {fmt_ai(actual)}, Target {fmt_ai(tgt)}, A/T {('—' if achievement is None else f'{achievement:,.1f}%')}, Deviation {fmt_ai(deviation)}")
            td=title_performance(scoped,program,selected)
            if not td.empty:
                below=td[pd.to_numeric(td["Target"],errors="coerce").fillna(0)>0]
                below=below[pd.to_numeric(below["Actual"],errors="coerce").fillna(0)<pd.to_numeric(below["Target"],errors="coerce")]
                below=below.sort_values("Deviation").head(5)
                if not below.empty:
                    lines.append("  Sales Person Title di bawah target terdalam:")
                    for _,row in below.iterrows():
                        at=row["A/T (%)"]
                        lines.append(f"  - {row['Sales Person Title']}: Actual {fmt_ai(row['Actual'])} vs Target {fmt_ai(row['Target'])} (A/T {('—' if at is None or pd.isna(at) else f'{at:,.1f}%')})")
    return "\n".join(lines)[:8000]

def ask_gemini(api_key,model,messages,context):
    url=f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent"
    body={"systemInstruction":{"parts":[{"text":GEMINI_SYSTEM+"\n\nData konteks dashboard terkini:\n"+context}]},
          "contents":[{"role":m["role"],"parts":[{"text":m["text"]}]} for m in messages]}
    resp=requests.post(url,headers={"x-goog-api-key":api_key},json=body,timeout=120)
    if resp.status_code!=200:
        raise RuntimeError(f"HTTP {resp.status_code}: {resp.text[:300]}")
    parts=resp.json()["candidates"][0]["content"].get("parts",[])
    return "".join(p.get("text","") for p in parts) or "(Gemini tidak mengembalikan teks.)"

with st.sidebar:
    st.divider()
    st.subheader("Chat AI · Gemini")
    env_key=os.environ.get("GEMINI_API_KEY","") or os.environ.get("GOOGLE_API_KEY","")
    api_key=st.text_input("Gemini API key",type="password",value=st.session_state.get("gemini_api_key","") or env_key,
                          help="Key hanya disimpan di sesi browser ini dan tidak ditulis ke file pengaturan.",
                          key="gemini_api_key_field")
    st.session_state["gemini_api_key"]=api_key
    if st.session_state.get("gemini_model") not in GEMINI_MODELS:
        st.session_state["gemini_model"]=GEMINI_MODELS[0]
    gemini_model=st.selectbox("Model Gemini",GEMINI_MODELS,key="gemini_model")
    if st.button("Bersihkan percakapan"):
        st.session_state["gemini_messages"]=[]
        st.rerun()
    for msg in st.session_state.setdefault("gemini_messages",[]):
        with st.chat_message("user" if msg["role"]=="user" else "assistant"):
            st.markdown(msg["text"])
    chat_prompt=st.chat_input("Tanya seputar performa, target, gap, atau strategi...")
    if chat_prompt:
        if not api_key:
            st.warning("Isi Gemini API key terlebih dahulu untuk memulai chat.")
        else:
            history=st.session_state["gemini_messages"]
            history.append({"role":"user","text":chat_prompt})
            with st.chat_message("user"):
                st.markdown(chat_prompt)
            with st.chat_message("assistant"):
                with st.spinner("Gemini menganalisis data dashboard..."):
                    try:
                        reply=ask_gemini(api_key,gemini_model,history,build_ai_context())
                    except Exception as exc:
                        reply=f"Gagal memanggil Gemini: {exc}"
            history.append({"role":"model","text":reply})
            with st.chat_message("assistant"):
                st.markdown(reply)
