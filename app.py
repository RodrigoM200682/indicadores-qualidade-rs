# app.py — INDICADORES QUALIDADE RS (WEB)
# - Login com senha única: QualidadeRS
# - Dashboard moderno (Plotly) com 4 gráficos coloridos por INTENSIDADE (2x2)
# - Exportar: Resumo Excel (DASHBOARD 2x2 + DADOS + RECORTE)
# - Exportar: PDF do Dashboard (1 página, 2x2, colorido)
#
# Requisitos (requirements.txt):
# streamlit
# pandas
# openpyxl
# plotly
# kaleido
# reportlab

import io
import pandas as pd
import streamlit as st
import plotly.express as px

from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.formatting.rule import CellIsRule
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.chart import BarChart
from openpyxl.chart.reference import Reference

from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.utils import ImageReader

# =========================================================
# APP
# =========================================================
APP_NAME = "INDICADORES QUALIDADE RS"
DEFAULT_SHEET = "Sheet1"
APP_PASSWORD = "QualidadeRS"

# Colunas esperadas
COL_CODIGO = "Código"
COL_TITULO = "Título"
COL_STATUS = "Status"
COL_DATA = "Data de emissão"
COL_MOTIVO = "Motivo Reclamação"
COL_TURNO = "Turno/Horário"
COL_RESP_OCORRENCIA = "Responsável"
COL_RESP_ANALISE = "Responsável da análise de causa"
COL_SITUACAO = "Situação"  # ATRASADA / NO PRAZO

FILTROS_COLS = [
    "Status",
    "Cliente",
    "Motivo Reclamação",
    "Responsável",
    "Responsável da análise de causa",
    "Turno/Horário",
    "Embalagem",
    COL_SITUACAO,
]

DATE_FMT_BR = "%d/%m/%Y"
MESES_ABREV = {
    1: "Jan", 2: "Fev", 3: "Mar", 4: "Abr", 5: "Mai", 6: "Jun",
    7: "Jul", 8: "Ago", 9: "Set", 10: "Out", 11: "Nov", 12: "Dez"
}
INV_MESES_ABREV = {v: k for k, v in MESES_ABREV.items()}

# Escolha da escala contínua (pode trocar por: "Viridis", "Plasma", "Magma", "Turbo", "Blues", "Reds" etc.)
CONTINUOUS_SCALE = "Turbo"


# =========================================================
# Login (senha fixa)
# =========================================================
def require_login():
    if "auth_ok" not in st.session_state:
        st.session_state.auth_ok = False

    if st.session_state.auth_ok:
        return

    st.title(f"🔒 {APP_NAME}")
    st.write("Acesso interno. Informe a senha para continuar.")
    p = st.text_input("Senha", type="password")

    if st.button("Entrar"):
        if p == APP_PASSWORD:
            st.session_state.auth_ok = True
            st.rerun()
        else:
            st.error("Senha inválida.")

    st.stop()


# =========================================================
# Utilitários
# =========================================================
def br_date_str(dt) -> str:
    if pd.isna(dt):
        return ""
    try:
        return pd.to_datetime(dt).strftime(DATE_FMT_BR)
    except Exception:
        return ""


def normalizar_situacao(x: str) -> str:
    s = str(x).strip().upper()
    if s in ("ATRASADA", "ATRASADO"):
        return "ATRASADA"
    if s in ("NO PRAZO", "NOPRAZO", "NO_PRAZO"):
        return "NO PRAZO"
    if s in ("", "NAN", "NONE"):
        return ""
    return s


def ano_padrao_para_relatorios(df: pd.DataFrame, col_data: str) -> int:
    anos = sorted(df[col_data].dt.year.dropna().unique().tolist())
    return int(anos[-1]) if anos else int(pd.Timestamp.today().year)


def _titulo_filtro(ano_sel: str, mes_sel: str, resp_occ_sel: str) -> str:
    return f"Ano {ano_sel} | Mês {mes_sel} | Resp ocorrência {resp_occ_sel}"


# =========================================================
# Excel helpers
# =========================================================
def _excel_theme():
    return {
        "title_fill": PatternFill("solid", fgColor="1F4E79"),
        "hdr_fill": PatternFill("solid", fgColor="2F5597"),
        "hdr_font": Font(bold=True, color="FFFFFF"),
        "thin": Side(style="thin", color="A6A6A6"),
        "kpi_fill": PatternFill("solid", fgColor="FFF2CC"),
    }


def _xl_col(n: int) -> str:
    s = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def _set_col_widths(ws, widths: dict):
    for col_letter, w in widths.items():
        ws.column_dimensions[col_letter].width = w


def _apply_border(ws, cell_range: str):
    th = _excel_theme()["thin"]
    border = Border(left=th, right=th, top=th, bottom=th)
    for row in ws[cell_range]:
        for c in row:
            c.border = border


def _merge_title(ws, cell_range: str, text: str):
    theme = _excel_theme()
    ws.merge_cells(cell_range)
    c = ws[cell_range.split(":")[0]]
    c.value = text
    c.fill = theme["title_fill"]
    c.font = Font(bold=True, color="FFFFFF", size=14)
    c.alignment = Alignment(horizontal="center", vertical="center")


def _add_table(ws, start_row, start_col, df: pd.DataFrame, table_name: str, style="TableStyleMedium9"):
    for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True), start=start_row):
        for c_idx, v in enumerate(row, start=start_col):
            ws.cell(row=r_idx, column=c_idx, value=v)

    end_row = start_row + len(df)
    end_col = start_col + df.shape[1] - 1
    ref = f"{_xl_col(start_col)}{start_row}:{_xl_col(end_col)}{end_row}"

    tab = Table(displayName=table_name, ref=ref)
    tab.tableStyleInfo = TableStyleInfo(
        name=style, showFirstColumn=False, showLastColumn=False,
        showRowStripes=True, showColumnStripes=False
    )
    ws.add_table(tab)

    theme = _excel_theme()
    for c in range(start_col, end_col + 1):
        cell = ws.cell(row=start_row, column=c)
        cell.fill = theme["hdr_fill"]
        cell.font = theme["hdr_font"]
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    ws.auto_filter.ref = ref
    _apply_border(ws, ref)
    return (start_row, start_col, end_row, end_col, ref)


def _chart_set_integer_y(chart: BarChart):
    try:
        chart.y_axis.majorUnit = 1
    except Exception:
        pass


def _chart_set_x_45(chart: BarChart):
    try:
        chart.x_axis.textRotation = 45
    except Exception:
        pass


def _add_bar_chart_from_sheet(
    data_ws, target_ws,
    title, cat_col, val_col, start_row, end_row, anchor_cell,
    y_title="Ocorrências", rotate_x_45=False,
    height=7.2, width=12.5
):
    chart = BarChart()
    chart.type = "col"
    chart.style = 10
    chart.title = title
    chart.y_axis.title = y_title
    chart.legend = None
    chart.height = float(height)
    chart.width = float(width)

    data = Reference(data_ws, min_col=val_col, min_row=start_row, max_row=end_row)
    cats = Reference(data_ws, min_col=cat_col, min_row=start_row + 1, max_row=end_row)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)

    _chart_set_integer_y(chart)
    if rotate_x_45:
        _chart_set_x_45(chart)

    target_ws.add_chart(chart, anchor_cell)


# =========================================================
# PDF do Dashboard (1 página) — Plotly -> PNG via kaleido
# =========================================================
def build_dashboard_pdf_bytes(app_name: str, filtro_txt: str, kpis: dict, figs_plotly: list) -> bytes:
    img_bytes_list = []
    for fig in figs_plotly:
        b = fig.to_image(format="png", scale=2)
        img_bytes_list.append(io.BytesIO(b))

    out = io.BytesIO()
    page_size = landscape(A4)
    c = canvas.Canvas(out, pagesize=page_size)
    W, H = page_size

    margin = 24
    header_h = 70
    footer_h = 18
    content_top = H - margin - header_h
    content_bottom = margin + footer_h
    content_h = content_top - content_bottom
    content_w = W - 2 * margin

    c.setFont("Helvetica-Bold", 16)
    c.drawString(margin, H - margin - 22, f"{app_name} — Dashboard")

    c.setFont("Helvetica", 9)
    c.drawString(margin, H - margin - 40, f"Filtro: {filtro_txt[:180]}")

    c.setFont("Helvetica-Bold", 10)
    kpi_line = (
        f"Total: {kpis.get('total', 0)}   |   "
        f"Em atraso: {kpis.get('atras', 0)}   |   "
        f"Período: {kpis.get('periodo', '-')}   |   "
        f"Ano ref: {kpis.get('ano_ref', '-')}"
    )
    c.drawString(margin, H - margin - 58, kpi_line)

    gap = 12
    cell_w = (content_w - gap) / 2
    cell_h = (content_h - gap) / 2

    positions = [(0, 0), (1, 0), (0, 1), (1, 1)]
    for i, (col, row) in enumerate(positions):
        if i >= len(img_bytes_list):
            break
        x = margin + col * (cell_w + gap)
        y = content_bottom + (1 - row) * (cell_h + gap)
        img = ImageReader(img_bytes_list[i])
        c.drawImage(img, x, y, width=cell_w, height=cell_h, preserveAspectRatio=True, anchor="c")

    c.setFont("Helvetica", 8)
    c.drawRightString(W - margin, margin / 2, f"Gerado em {pd.Timestamp.now().strftime('%d/%m/%Y %H:%M')}")

    c.showPage()
    c.save()
    out.seek(0)
    return out.read()


# =========================================================
# Carregamento e filtros
# =========================================================
@st.cache_data(show_spinner=False)
def carregar_df(upload_bytes: bytes, sheet_name: str) -> pd.DataFrame:
    df = pd.read_excel(io.BytesIO(upload_bytes), sheet_name=sheet_name)

    obrig = [COL_CODIGO, COL_TITULO, COL_STATUS, COL_DATA, COL_MOTIVO]
    for c in obrig:
        if c not in df.columns:
            raise ValueError(f"Não encontrei a coluna obrigatória '{c}' na planilha: {c}")

    df = df.copy()
    df[COL_DATA] = pd.to_datetime(df[COL_DATA], errors="coerce", dayfirst=True)
    df = df.dropna(subset=[COL_DATA])

    for c in df.columns:
        if df[c].dtype == "object":
            df[c] = df[c].astype(str).str.strip().replace("nan", "")

    if COL_SITUACAO in df.columns:
        df[COL_SITUACAO] = df[COL_SITUACAO].apply(normalizar_situacao)

    return df


def aplicar_filtros(df: pd.DataFrame, ano_sel, mes_sel, resp_occ_sel, multi_filters: dict) -> pd.DataFrame:
    dff = df.copy()

    if ano_sel != "(Todos)":
        dff = dff[dff[COL_DATA].dt.year == int(ano_sel)]

    if mes_sel != "(Todos)":
        mes_num = INV_MESES_ABREV.get(mes_sel)
        if mes_num:
            dff = dff[dff[COL_DATA].dt.month == int(mes_num)]

    if resp_occ_sel != "(Todos)" and COL_RESP_OCORRENCIA in dff.columns:
        dff = dff[dff[COL_RESP_OCORRENCIA].astype(str) == resp_occ_sel]

    for col, selecionados in multi_filters.items():
        if col not in dff.columns:
            continue
        if not selecionados:
            return dff.iloc[0:0]
        dff = dff[dff[col].astype(str).isin(selecionados)]

    return dff


# =========================================================
# Cálculos de gráficos (base para dashboard/PDF)
# =========================================================
def calc_datasets(df_base: pd.DataFrame, df_filtrado: pd.DataFrame, ano_ref: int):
    # 1) mês (recorte)
    dff_mes = df_filtrado.copy()
    dff_mes["MesNum"] = dff_mes[COL_DATA].dt.month.astype(int)
    g_mes = dff_mes.groupby("MesNum").size().reindex(range(1, 13), fill_value=0)
    df_mes = pd.DataFrame({"Mês": [MESES_ABREV[m] for m in range(1, 13)], "Ocorrências": g_mes.values.astype(int)})

    # 2) resp análise (recorte)
    resp_rec = (
        df_filtrado[COL_RESP_ANALISE].fillna("").replace("", "SEM RESPONSÁVEL")
        if COL_RESP_ANALISE in df_filtrado.columns else pd.Series(["SEM RESPONSÁVEL"] * len(df_filtrado))
    )
    df_resp = resp_rec.value_counts().reset_index()
    df_resp.columns = ["Responsável (análise)", "Ocorrências"]
    if df_resp.empty:
        df_resp = pd.DataFrame({"Responsável (análise)": ["SEM DADOS"], "Ocorrências": [0]})

    # 3) atrasadas resp análise (ano todo)
    df_ano = df_base[df_base[COL_DATA].dt.year == ano_ref].copy()
    resp_ano = (
        df_ano[COL_RESP_ANALISE].fillna("").replace("", "SEM RESPONSÁVEL")
        if COL_RESP_ANALISE in df_ano.columns else pd.Series(["SEM RESPONSÁVEL"] * len(df_ano))
    )
    sit_ano = (
        df_ano[COL_SITUACAO].fillna("").apply(normalizar_situacao)
        if COL_SITUACAO in df_ano.columns else pd.Series([""] * len(df_ano))
    )

    df_atras = (
        pd.DataFrame({"Responsável (análise)": resp_ano, "Situação": sit_ano})
        .query("Situação == 'ATRASADA'")["Responsável (análise)"]
        .value_counts()
        .reset_index()
    )
    df_atras.columns = ["Responsável (análise)", "Atrasadas (ano todo)"]
    if df_atras.empty:
        df_atras = pd.DataFrame({"Responsável (análise)": ["SEM DADOS"], "Atrasadas (ano todo)": [0]})

    # 4) motivo (top 12)
    top_mot = (
        df_filtrado[COL_MOTIVO].fillna("").replace("", "SEM MOTIVO").value_counts().head(12)
        if COL_MOTIVO in df_filtrado.columns else pd.Series(dtype=int)
    )
    df_mot = top_mot.reset_index()
    df_mot.columns = ["Motivo", "Ocorrências"]
    if df_mot.empty:
        df_mot = pd.DataFrame({"Motivo": ["SEM DADOS"], "Ocorrências": [0]})

    return df_mes, df_resp, df_atras, df_mot


# =========================================================
# Gráficos por INTENSIDADE (escala contínua)
# =========================================================
def build_intensity_figs(df_mes, df_resp, df_atras, df_mot, ano_ref: int):
    # Nota: color = medida numérica => escala contínua por intensidade
    fig_mes = px.bar(
        df_mes, x="Mês", y="Ocorrências", color="Ocorrências",
        color_continuous_scale=CONTINUOUS_SCALE,
        title="Ocorrências por mês (filtro)"
    )
    fig_mes.update_layout(yaxis=dict(dtick=1), coloraxis_showscale=False)

    fig_resp = px.bar(
        df_resp, x="Responsável (análise)", y="Ocorrências", color="Ocorrências",
        color_continuous_scale=CONTINUOUS_SCALE,
        title="Ocorrências por responsável (análise) — filtro"
    )
    fig_resp.update_layout(xaxis_tickangle=-45, yaxis=dict(dtick=1), coloraxis_showscale=False)

    fig_atras = px.bar(
        df_atras, x="Responsável (análise)", y="Atrasadas (ano todo)", color="Atrasadas (ano todo)",
        color_continuous_scale=CONTINUOUS_SCALE,
        title=f"Atrasadas por responsável (análise) — ano {ano_ref}"
    )
    fig_atras.update_layout(xaxis_tickangle=-45, yaxis=dict(dtick=1), coloraxis_showscale=False)

    fig_mot = px.bar(
        df_mot, x="Motivo", y="Ocorrências", color="Ocorrências",
        color_continuous_scale=CONTINUOUS_SCALE,
        title="Ocorrências por motivo (Top 12) — filtro"
    )
    fig_mot.update_layout(xaxis_tickangle=-45, yaxis=dict(dtick=1), coloraxis_showscale=False)

    return fig_mes, fig_resp, fig_atras, fig_mot


# =========================================================
# Resumo Excel (DASHBOARD 2x2 + DADOS + RECORTE)
# =========================================================
def build_resumo_excel_bytes(df_base: pd.DataFrame, df_filtrado: pd.DataFrame, titulo_filtro: str, ano_ref: int) -> bytes:
    dff = df_filtrado.copy()
    theme = _excel_theme()

    total_rec = int(len(dff))
    p_ini = br_date_str(dff[COL_DATA].min()) if total_rec else "-"
    p_fim = br_date_str(dff[COL_DATA].max()) if total_rec else "-"

    if COL_SITUACAO in dff.columns and total_rec:
        sit_rec = dff[COL_SITUACAO].fillna("").apply(normalizar_situacao)
        total_atras_rec = int((sit_rec == "ATRASADA").sum())
    else:
        total_atras_rec = 0

    df_mes, df_resp, df_atras, df_mot = calc_datasets(df_base, dff, ano_ref)

    wb = Workbook()

    ws = wb.active
    ws.title = "DASHBOARD"
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 90
    _set_col_widths(ws, {"A": 2, "B": 28, "C": 28, "D": 28, "E": 28, "F": 2})

    _merge_title(ws, "B2:E2", "RESUMO DE OCORRÊNCIAS — DASHBOARD (4 GRÁFICOS)")
    ws.row_dimensions[2].height = 26

    ws["B4"] = "Filtro:"; ws["B4"].font = Font(bold=True)
    ws["C4"] = titulo_filtro
    ws.merge_cells("C4:E4")
    ws["C4"].alignment = Alignment(wrap_text=True, vertical="top")

    ws["B6"] = "Período (recorte):"; ws["B6"].font = Font(bold=True)
    ws["C6"] = f"{p_ini} a {p_fim}"
    ws["D6"] = "Ano ref (atrasadas ano todo):"; ws["D6"].font = Font(bold=True)
    ws["E6"] = ano_ref

    for row in ws["B8:E10"]:
        for c in row:
            c.fill = theme["kpi_fill"]
            c.alignment = Alignment(vertical="center", wrap_text=True)
    _apply_border(ws, "B8:E10")

    ws["B8"] = "TOTAIS (FILTRO ATUAL)"; ws["B8"].font = Font(bold=True, size=12); ws.merge_cells("B8:E8")
    ws["B9"] = "Total de ocorrências"; ws["B9"].font = Font(bold=True)
    ws["C9"] = total_rec; ws["C9"].font = Font(bold=True, size=14)
    ws["D9"] = "Ocorrências em atraso"; ws["D9"].font = Font(bold=True)
    ws["E9"] = total_atras_rec; ws["E9"].font = Font(bold=True, size=14)
    ws["B10"] = "Obs.: tabelas base ficam na aba DADOS."; ws.merge_cells("B10:E10")

    wsd = wb.create_sheet("DADOS")
    wsd.sheet_view.showGridLines = True
    _set_col_widths(wsd, {"A": 2, "B": 34, "C": 20, "D": 34, "E": 20, "F": 2})
    _merge_title(wsd, "B2:E2", "DADOS — NÃO EDITAR (BASE DOS GRÁFICOS)")
    wsd.row_dimensions[2].height = 22

    r = 4
    wsd[f"B{r}"] = "1) Ocorrências por mês (filtro)"; wsd[f"B{r}"].font = Font(bold=True)
    r1s, _, r1e, _, _ = _add_table(wsd, r + 1, 2, df_mes, table_name="T_MES", style="TableStyleMedium9")

    r = r1e + 3
    wsd[f"B{r}"] = "2) Ocorrências por responsável (análise) — filtro"; wsd[f"B{r}"].font = Font(bold=True)
    r2s, _, r2e, _, _ = _add_table(wsd, r + 1, 2, df_resp, table_name="T_RESP", style="TableStyleMedium9")

    r = r2e + 3
    wsd[f"B{r}"] = f"3) Atrasadas por responsável (análise) — ano {ano_ref} (ano todo)"; wsd[f"B{r}"].font = Font(bold=True)
    r3s, _, r3e, _, _ = _add_table(wsd, r + 1, 2, df_atras, table_name="T_ATRAS_ANO", style="TableStyleMedium7")

    try:
        rng = f"C{r3s+1}:C{r3e}"
        wsd.conditional_formatting.add(
            rng, CellIsRule(operator="greaterThan", formula=["0"], fill=PatternFill("solid", fgColor="FFC7CE"))
        )
    except Exception:
        pass

    r = r3e + 3
    wsd[f"B{r}"] = "4) Ocorrências por motivo (Top 12) — filtro"; wsd[f"B{r}"].font = Font(bold=True)
    r4s, _, r4e, _, _ = _add_table(wsd, r + 1, 2, df_mot, table_name="T_MOT", style="TableStyleMedium9")

    _add_bar_chart_from_sheet(wsd, ws, "Ocorrências por mês (filtro)", 2, 3, r1s, r1e, "B12", "Ocorrências", False)
    _add_bar_chart_from_sheet(wsd, ws, "Ocorrências por responsável (análise) — filtro", 2, 3, r2s, r2e, "D12", "Ocorrências", True)
    _add_bar_chart_from_sheet(wsd, ws, f"Atrasadas por responsável (análise) — ano {ano_ref}", 2, 3, r3s, r3e, "B28", "Atrasadas", True)
    _add_bar_chart_from_sheet(wsd, ws, "Ocorrências por motivo (Top 12) — filtro", 2, 3, r4s, r4e, "D28", "Ocorrências", True)

    ws2 = wb.create_sheet("RECORTE")
    ws2.sheet_view.showGridLines = True
    _merge_title(ws2, "A1:H1", "LISTA DE OCORRÊNCIAS — FILTRO ATUAL")
    ws2.row_dimensions[1].height = 26

    cols_doc = [COL_CODIGO, COL_TITULO, COL_STATUS, COL_DATA, COL_MOTIVO, COL_RESP_ANALISE, COL_SITUACAO]
    cols_doc = [c for c in cols_doc if c in dff.columns]

    dff_out = dff.copy()
    if COL_DATA in dff_out.columns:
        dff_out = dff_out.sort_values(COL_DATA, ascending=False).copy()
        dff_out[COL_DATA] = dff_out[COL_DATA].apply(br_date_str)
    if cols_doc:
        dff_out = dff_out[cols_doc].copy()

    _add_table(ws2, 3, 1, dff_out, table_name="T_RECORTE", style="TableStyleMedium9")
    ws2.column_dimensions["A"].width = 14
    ws2.column_dimensions["B"].width = 70
    ws2.column_dimensions["C"].width = 16
    ws2.column_dimensions["D"].width = 16
    ws2.column_dimensions["E"].width = 26
    ws2.column_dimensions["F"].width = 30
    ws2.column_dimensions["G"].width = 14

    if COL_SITUACAO in cols_doc:
        sit_col_idx = cols_doc.index(COL_SITUACAO) + 1
        sit_col_letter = _xl_col(sit_col_idx)
        data_end = 3 + len(dff_out)
        rng = f"{sit_col_letter}{4}:{sit_col_letter}{data_end}"
        ws2.conditional_formatting.add(
            rng, CellIsRule(operator="equal", formula=['"ATRASADA"'], fill=PatternFill("solid", fgColor="FFC7CE"))
        )
        ws2.conditional_formatting.add(
            rng, CellIsRule(operator="equal", formula=['"NO PRAZO"'], fill=PatternFill("solid", fgColor="C6EFCE"))
        )

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out.read()


# =========================================================
# UI Streamlit
# =========================================================
st.set_page_config(page_title=APP_NAME, page_icon="📊", layout="wide")
require_login()

# Tema plotly base (colorido)
try:
    import plotly.io as pio
    pio.templates.default = "plotly"
except Exception:
    pass

st.title(f"📊 {APP_NAME}")
st.caption("Gráficos por intensidade (quanto maior o valor, mais forte a cor), com exportação para Excel e PDF.")

with st.sidebar:
    st.header("📥 Entrada")
    up = st.file_uploader("Envie o Excel (ex.: Consultas_RNC.xlsx)", type=["xlsx", "xlsm", "xls"])
    sheet = st.text_input("Nome da aba (sheet)", value=DEFAULT_SHEET)
    st.divider()
    st.caption("Senha do app: QualidadeRS")

if not up:
    st.info("Envie o arquivo Excel para começar.")
    st.stop()

try:
    df_base = carregar_df(up.getvalue(), sheet)
except Exception as e:
    st.error(f"Erro ao carregar: {e}")
    st.stop()

anos = sorted(df_base[COL_DATA].dt.year.dropna().unique().tolist())
c1, c2, c3, c4 = st.columns([1, 1, 1.6, 1.2])

with c1:
    ano_sel = st.selectbox("Ano", ["(Todos)"] + [str(a) for a in anos], index=0)
with c2:
    mes_sel = st.selectbox("Mês", ["(Todos)"] + [MESES_ABREV[m] for m in range(1, 13)], index=0)
with c3:
    if COL_RESP_OCORRENCIA in df_base.columns:
        resp_vals = sorted(df_base[COL_RESP_OCORRENCIA].dropna().astype(str).replace("nan", "").unique().tolist())
        resp_vals = [v for v in resp_vals if v != ""]
    else:
        resp_vals = []
    resp_occ_sel = st.selectbox("Resp. ocorrência", ["(Todos)"] + resp_vals, index=0)
with c4:
    show_table = st.toggle("Mostrar tabela", value=True)

with st.expander("Filtros por marcar (clique para abrir)", expanded=False):
    cols = st.columns(4)
    multi_filters = {}
    for i, col in enumerate(FILTROS_COLS):
        if col not in df_base.columns:
            continue
        vals = sorted(df_base[col].dropna().astype(str).replace("nan", "").unique().tolist())
        vals = [v for v in vals if v != ""]
        with cols[i % 4]:
            sel = st.multiselect(col, options=vals, default=vals)
        multi_filters[col] = sel

df_filtrado = aplicar_filtros(df_base, ano_sel, mes_sel, resp_occ_sel, multi_filters)

total = int(len(df_filtrado))
situ = df_filtrado[COL_SITUACAO].apply(normalizar_situacao) if (COL_SITUACAO in df_filtrado.columns and total) else pd.Series([], dtype=str)
atras = int((situ == "ATRASADA").sum()) if total else 0
p_ini = br_date_str(df_filtrado[COL_DATA].min()) if total else "-"
p_fim = br_date_str(df_filtrado[COL_DATA].max()) if total else "-"
ano_ref = int(ano_sel) if ano_sel != "(Todos)" else ano_padrao_para_relatorios(df_base, COL_DATA)

k1, k2, k3, k4 = st.columns(4)
k1.metric("Total ocorrências", total)
k2.metric("Em atraso", atras)
k3.metric("Período", f"{p_ini} → {p_fim}")
k4.metric("Ano ref (atrasadas ano todo)", ano_ref)

st.divider()
tab1, tab2 = st.tabs(["📈 Dashboard", "📦 Exportações (Excel/PDF)"])

with tab1:
    if not total:
        st.warning("Sem registros no filtro atual.")
        st.stop()

    df_mes, df_resp, df_atras, df_mot = calc_datasets(df_base, df_filtrado, ano_ref)
    fig_mes, fig_resp, fig_atras, fig_mot = build_intensity_figs(df_mes, df_resp, df_atras, df_mot, ano_ref)

    g1, g2 = st.columns(2)
    g3, g4 = st.columns(2)

    with g1:
        st.plotly_chart(fig_mes, use_container_width=True)
    with g2:
        st.plotly_chart(fig_resp, use_container_width=True)
    with g3:
        st.plotly_chart(fig_atras, use_container_width=True)
    with g4:
        st.plotly_chart(fig_mot, use_container_width=True)

    if show_table:
        st.subheader("Recorte (tabela)")
        st.dataframe(df_filtrado.sort_values(COL_DATA, ascending=False), use_container_width=True, height=380)

with tab2:
    if not total:
        st.info("Quando houver registros no filtro, as exportações ficarão disponíveis.")
        st.stop()

    filtro_txt = _titulo_filtro(ano_sel, mes_sel, resp_occ_sel)
    kpis_pdf = {"total": total, "atras": atras, "periodo": f"{p_ini} → {p_fim}", "ano_ref": ano_ref}

    st.subheader("📄 PDF do Dashboard (1 página, 4 gráficos)")
    try:
        df_mes, df_resp, df_atras, df_mot = calc_datasets(df_base, df_filtrado, ano_ref)
        fig_mes, fig_resp, fig_atras, fig_mot = build_intensity_figs(df_mes, df_resp, df_atras, df_mot, ano_ref)

        pdf_bytes = build_dashboard_pdf_bytes(
            app_name=APP_NAME,
            filtro_txt=filtro_txt,
            kpis=kpis_pdf,
            figs_plotly=[fig_mes, fig_resp, fig_atras, fig_mot],
        )

        st.download_button(
            label="📄 Baixar PDF do Dashboard",
            data=pdf_bytes,
            file_name=f"Dashboard_{APP_NAME.replace(' ', '_')}.pdf",
            mime="application/pdf",
        )
    except Exception as e:
        st.error(f"Erro ao gerar PDF. Detalhe: {e}")
        st.caption("Se citar 'kaleido', reinstale requirements e reinicie: py -m pip install -r requirements.txt")

    st.divider()
    st.subheader("📊 Resumo Excel (DASHBOARD + DADOS + RECORTE)")

    titulo_filtro = f"Reclamações — Filtro atual | {filtro_txt}"
    resumo_bytes = build_resumo_excel_bytes(df_base, df_filtrado, titulo_filtro, ano_ref)

    st.download_button(
        label="📥 Baixar Resumo Excel",
        data=resumo_bytes,
        file_name=f"Resumo_{APP_NAME.replace(' ', '_')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

