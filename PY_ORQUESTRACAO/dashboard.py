# -*- coding: utf-8 -*-
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import streamlit as st
import pandas as pd
import plotly.express as px
from bim_paths import XLSX_DASH

st.set_page_config(page_title="Maturidade BIM", layout="wide")
st.title("Dashboard de Maturidade BIM - MODELOS ATIVOS")

df  = pd.read_excel(XLSX_DASH, sheet_name="Detalhado")
res = pd.read_excel(XLSX_DASH, sheet_name="Resumo")

c1, c2, c3 = st.columns(3)
c1.metric("Modelos", len(res))
c2.metric("Nota media", round(res["Nota Final"].mean(), 2))
c3.metric("Melhor", res.iloc[0]["Arquivo"])

st.subheader("Ranking")
st.dataframe(res, width='stretch')

st.subheader("Heatmap por check")
pivot = df.pivot_table(index="Arquivo", columns="Descricao",
                       values="Nota (0-10)", aggfunc="mean")
st.plotly_chart(px.imshow(pivot, color_continuous_scale="RdYlGn", aspect="auto"),
                width='stretch')

st.subheader("Radar")
st.plotly_chart(px.line_polar(df, r="Nota (0-10)", theta="Descricao",
                              color="Arquivo", line_close=True),
                width='stretch')
