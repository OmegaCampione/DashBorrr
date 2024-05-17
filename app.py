import pandas as pd
import plotly_express as px
import streamlit as st

st.set_page_config(page_title="Dashboard epico",
                   page_icon=":bar_chart:",
                   layout="wide")

@st.cache_data
def pegar_dados_do_excel():
    df = pd.read_excel(
        io='Arquivo final.xlsx',
        engine='openpyxl',
        sheet_name='Arquivo final',
        nrows=7225
    )
    
    #df['ANO'] = pd.to_datetime(df['ANO'], format='%d-%m-%Y')
    df['ANO'] = df['ANO'].astype(int)
    return df


df = pegar_dados_do_excel()

#---sidebar---#
st.sidebar.header("Filtro Aqui")
gerente = st.sidebar.multiselect(
    "Selecione o Gerente",
    options=df["GERENTE"].unique(),
    default=df["GERENTE"].unique()
)

ano = st.sidebar.multiselect(
    "Selecione o Ano",
    options=df["ANO"].unique(),
    default=df["ANO"].unique()
)

categoria = st.sidebar.multiselect(
    "Selecione a Categoria",
    options=df["CATEGORIA"].unique(),
    default=df["CATEGORIA"].unique()
)

df_selection = df.query(
    "GERENTE == @gerente & ANO == @ano & CATEGORIA ==@categoria"
)

#---main---#
st.title(":bar_chart: Dashboard Épico")
st.markdown("##")

#kdi#
vendas_totais = int(df_selection["VENDA"].sum())
vendas_media = int(df_selection["VENDA"].mean())

left_column, middle_column, right_column = st.columns (3)
with left_column:
    st.subheader("Vendas Totais:")
    st.subheader(f"R$ {vendas_totais:,}")
with middle_column:
    st.subheader("Média das Vendas:")
    st.subheader(f"R$ {vendas_media:,}")

st.markdown("---")

#Vendas por Categoria (bar chart)
vendas_por_cat = (
    df_selection.groupby(by=["CATEGORIA"])[["VENDA"]].sum().sort_values(by="VENDA")
)
fig_prod_vendas = px.bar(
    vendas_por_cat,
    x="VENDA",
    y=vendas_por_cat.index,
    orientation="h",
    title="<b>Vendas por Categoria</b>",
    color_discrete_sequence=["#00083B"] * len(vendas_por_cat),
    template="plotly_white",
)

fig_prod_vendas.update_layout(
    plot_bgcolor="rgba(0,0,0,0)",
    xaxis=(dict(showgrid=False)),
)

st.plotly_chart(fig_prod_vendas)

#Vendas por ANO
vendas_por_ano = df_selection.groupby(by=["ANO"])[["VENDA"]].sum().sort_values(by="ANO")  # 12
fif_ano_vendas = px.bar(
    vendas_por_ano,
    x=vendas_por_ano.index,  # 13
    y="VENDA",
    orientation="h",
    title="<b>Vendas por Ano</b>",
    color_discrete_sequence=["#00083B"] * len(vendas_por_ano),
    template="plotly_white",
)

fif_ano_vendas.update_layout(
    xaxis=dict(tickmode="linear"),
    plot_bgcolor="rgba(0,0,0,0)",
    yaxis=(dict(showgrid=False)),
)

st.plotly_chart(fif_ano_vendas)

