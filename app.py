import pandas as pd
import plotly.express as px
import streamlit as st

st.set_page_config(page_title="Dashboard épico",
                   page_icon=":bar_chart:",
                   layout="wide")

@st.cache_data
def pegar_dados_do_excel():
    df = pd.read_excel(
        io='Arquivo final.xlsx',
        engine='openpyxl',
        sheet_name='Arquivo final',
        nrows=7225,  # Limite o número de linhas se possível
        usecols=["GERENTE", "ANO", "CATEGORIA", "VENDA"]  # Leia apenas as colunas necessárias
    )
    
    # Convertendo a coluna ANO para inteiros
    df['ANO'] = df['ANO'].astype(int)
    
    return df

df = pegar_dados_do_excel()

# ---sidebar---#
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
    "GERENTE == @gerente & ANO == @ano & CATEGORIA == @categoria"
)

# Convertendo a coluna ANO para inteiros após filtrar
df_selection['ANO'] = df_selection['ANO'].astype(int)

# ---main---#
st.title(":bar_chart: Dashboard Épico")
st.markdown("##")

# kpi #
vendas_totais = int(df_selection["VENDA"].sum())
vendas_media = int(df_selection["VENDA"].mean())

left_column, middle_column, right_column = st.columns(3)
with left_column:
    st.subheader("Vendas Totais:")
    st.subheader(f"R$ {vendas_totais:,}")
with middle_column:
    st.subheader("Média das Vendas:")
    st.subheader(f"R$ {vendas_media:,}")

st.markdown("---")

# Vendas por Categoria (bar chart)
vendas_por_cat = df_selection.groupby(by=["CATEGORIA"])[["VENDA"]].sum().sort_values(by="VENDA")
fig_prod_vendas = px.bar(
    vendas_por_cat,
    x="VENDA",
    y=vendas_por_cat.index,
    orientation="h",
    title="<b>Vendas por Categoria</b>",
    color_discrete_sequence=["#4CAF50"] * len(vendas_por_cat),
    template="plotly_white",
)

fig_prod_vendas.update_layout(
    plot_bgcolor="rgba(0,0,0,0)",
    xaxis=dict(showgrid=False),
)


# Vendas por ANO (refatorado)
vendas_por_ano = df_selection.groupby(by=["ANO"])[["VENDA"]].sum().sort_values(by="VENDA")

fig_ano_vendas = px.bar(
    vendas_por_ano,
    x=vendas_por_ano.index,
    y="VENDA",
    title="<b>Vendas por Ano</b>",
    color_discrete_sequence=["#4CAF50"] * len(vendas_por_ano),
    template="plotly_white",
)

fig_ano_vendas.update_layout(
    xaxis=dict(tickmode="linear"),
    plot_bgcolor="rgba(0,0,0,0)",
    yaxis=dict(showgrid=False, tickformat="d"),  # Formatando o eixo y para inteiros
)

# Vendas por Categoria e Ano (stacked bar chart)
vendas_por_cat_ano = df_selection.groupby(by=["ANO", "CATEGORIA"])[["VENDA"]].sum().reset_index()

custom_colors = ["#E694FF", "#FFA07A", "#20B2AA", "#4CAF50", "#FF6347"]
fig_cat_ano_vendas = px.bar(
    vendas_por_cat_ano,
    x="ANO",
    y="VENDA",
    color="CATEGORIA",
    title="<b>Vendas por Categoria e Ano</b>",
    template="plotly_white",
    barmode="stack",
    color_discrete_sequence=custom_colors
)

fig_cat_ano_vendas.update_layout(
    plot_bgcolor="rgba(0,0,0,0)",
    yaxis=dict(showgrid=False),
)

left_column, right_column = st.columns(2)
left_column.plotly_chart(fig_prod_vendas, use_container_width=True)
right_column.plotly_chart(fig_ano_vendas, use_container_width=True)

# Vendas por Gerente (bar chart)
vendas_por_gerente = df_selection.groupby(by=["GERENTE"])[["VENDA"]].sum().sort_values(by="VENDA")
fig_gerente_vendas = px.bar(
    vendas_por_gerente,
    x=vendas_por_gerente.index,
    y="VENDA",
    title="<b>Vendas por Gerente</b>",
    color_discrete_sequence=["#4CAF50"] * len(vendas_por_gerente),
    template="plotly_white",
)

fig_gerente_vendas.update_layout(
    plot_bgcolor="rgba(0,0,0,0)",
    yaxis=dict(showgrid=False),
)

left_column, right_column = st.columns(2)
left_column.plotly_chart(fig_cat_ano_vendas, use_container_width=True)
right_column.plotly_chart(fig_gerente_vendas, use_container_width=True)

# ---Esconder Streamlit.Style ---
esconder_st_style = """
            <style>
            #MainMenu {visibility: hidden;}
            footer {visibility: hidden;}
            header {visibility: hidden;}
            </style>
            """
st.markdown(esconder_st_style, unsafe_allow_html=True)
