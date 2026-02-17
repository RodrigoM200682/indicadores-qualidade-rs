# =========================================================
# INDICADORES QUALIDADE RS - versão web profissional
# =========================================================
import io
import pandas as pd
import streamlit as st
import plotly.express as px

from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.chart import BarChart
from openpyxl.chart.reference import Reference

# =========================================================
APP_NAME = "INDICADORES QUALIDADE RS"
APP_PASSWORD = "QualidadeRS"
DEFAULT_SHEET = "Sheet1"

COL_DATA = "Data de emissão"
COL_MOTIVO = "Motivo Reclamação"
COL_RESP_ANALISE = "Responsável da análise de causa"
COL_CATEGORIA = "Categoria"
COL_SITUACAO = "Situação"

# =========================================================
# LOGIN
# =========================================================
def login():
    if "ok" not in st.session_state:
        st.session_state.ok = False

    if st.session_state.ok:
        return

    st.title(APP_NAME)
    senha = st.text_input("Senha", type="password")

    if st.button("Entrar"):
        if senha == APP_PASSWORD:
            st.session_state.ok = True
            st.rerun()
        else:
            st.error("Senha inválida")

    st.stop()

login()

# =========================================================
st.set_page_config(layout="wide")
st.title("📊 INDICADORES QUALIDADE RS")

file = st.sidebar.file_uploader("Enviar Excel", type=["xlsx"])

if not file:
    st.stop()

df = pd.read_excel(file, sheet_name=DEFAULT_SHEET)
df[COL_DATA] = pd.to_datetime(df[COL_DATA], errors="coerce", dayfirst=True)
df = df.dropna(subset=[COL_DATA])

# =========================================================
# FILTROS
# =========================================================
anos = sorted(df[COL_DATA].dt.year.unique())
ano = st.sidebar.selectbox("Ano", ["Todos"] + anos)

meses = ["Todos","Jan","Fev","Mar","Abr","Mai","Jun","Jul","Ago","Set","Out","Nov","Dez"]
mes = st.sidebar.selectbox("Mês", meses)

resp = st.sidebar.multiselect(
    "Responsável",
    df[COL_RESP_ANALISE].dropna().unique(),
    default=df[COL_RESP_ANALISE].dropna().unique()
)

# NOVO FILTRO CATEGORIA
if COL_CATEGORIA in df.columns:
    categoria = st.sidebar.multiselect(
        "Categoria",
        df[COL_CATEGORIA].dropna().unique(),
        default=df[COL_CATEGORIA].dropna().unique()
    )
else:
    categoria = []

# =========================================================
# APLICA FILTROS
# =========================================================
dff = df.copy()

if ano != "Todos":
    dff = dff[dff[COL_DATA].dt.year == int(ano)]

if mes != "Todos":
    num = meses.index(mes)
    dff = dff[dff[COL_DATA].dt.month == num]

if resp:
    dff = dff[dff[COL_RESP_ANALISE].isin(resp)]

if categoria:
    dff = dff[dff[COL_CATEGORIA].isin(categoria)]

# =========================================================
# KPIs
# =========================================================
col1,col2,col3 = st.columns(3)
col1.metric("Ocorrências", len(dff))
col2.metric("Atrasadas", (dff[COL_SITUACAO]=="ATRASADA").sum())
col3.metric("Ano referência", ano if ano!="Todos" else dff[COL_DATA].dt.year.max())

# =========================================================
# GRÁFICOS COLORIDOS
# =========================================================
st.divider()

c1,c2 = st.columns(2)
c3,c4 = st.columns(2)

# MÊS
mes_df = dff.groupby(dff[COL_DATA].dt.month).size().reset_index(name="Ocorrências")
mes_df["Mês"] = mes_df[COL_DATA].map({
1:"Jan",2:"Fev",3:"Mar",4:"Abr",5:"Mai",6:"Jun",
7:"Jul",8:"Ago",9:"Set",10:"Out",11:"Nov",12:"Dez"})

fig1 = px.bar(mes_df,x="Mês",y="Ocorrências",color="Ocorrências",
              color_continuous_scale="Turbo")
c1.plotly_chart(fig1,use_container_width=True)

# RESPONSÁVEL
resp_df = dff.groupby(COL_RESP_ANALISE).size().reset_index(name="Ocorrências")
fig2 = px.bar(resp_df,x=COL_RESP_ANALISE,y="Ocorrências",color="Ocorrências",
              color_continuous_scale="Turbo")
fig2.update_layout(xaxis_tickangle=-45)
c2.plotly_chart(fig2,use_container_width=True)

# CATEGORIA
if COL_CATEGORIA in dff.columns:
    cat_df = dff.groupby(COL_CATEGORIA).size().reset_index(name="Ocorrências")
    fig3 = px.bar(cat_df,x=COL_CATEGORIA,y="Ocorrências",color="Ocorrências",
                  color_continuous_scale="Turbo")
    c3.plotly_chart(fig3,use_container_width=True)

# MOTIVO
mot_df = dff.groupby(COL_MOTIVO).size().reset_index(name="Ocorrências")
fig4 = px.bar(mot_df,x=COL_MOTIVO,y="Ocorrências",color="Ocorrências",
              color_continuous_scale="Turbo")
fig4.update_layout(xaxis_tickangle=-45)
c4.plotly_chart(fig4,use_container_width=True)

# =========================================================
# TABELA
# =========================================================
st.dataframe(dff,use_container_width=True)