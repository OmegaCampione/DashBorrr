import pandas as pd
import plotly.express as px
import streamlit as st
from io import BytesIO

st.set_page_config(page_title="Dashboard épico",
                   page_icon=":chart_with_upwards_trend:",
                   layout="wide")

@st.cache_data
def pegar_dados_do_excel():
    df = pd.read_excel(
        io='Arquivo final.xlsx',
        engine='openpyxl',
        sheet_name='Arquivo final',
        nrows=7225,  # Limite o número de linhas se possível
        usecols=["GERENTE", "MES", "DIA", "ANO", "CATEGORIA", "VENDA"]
    )
    
    # Convertendo a coluna ANO para inteiros
    df['ANO'] = df['ANO'].astype(int)
    
    return df

df = pegar_dados_do_excel()

# Convert 'DIA' p/ datetime.date
df['DIA'] = pd.to_datetime(df['DIA']).dt.date

# ---sidebar---#
st.sidebar.header("Filtre Aqui:")
gerente = st.sidebar.multiselect(
    "Selecione o Gerente",
    options=df["GERENTE"].unique(),
    default=df["GERENTE"].unique()
)

mes = st.sidebar.multiselect(
    "Selecione o Mês:",
    options=df["MES"].unique(),
    default=df["MES"].unique()
)

ano_range = st.sidebar.slider(
    "Selecione o Ano:",
    min_value=int(df["ANO"].min()),
    max_value=int(df["ANO"].max()),
    value=(int(df["ANO"].min()), int(df["ANO"].max()))
)

categoria = st.sidebar.multiselect(
    "Selecione a Categoria:",
    options=df["CATEGORIA"].unique(),
    default=df["CATEGORIA"].unique()
)

venda_range = st.sidebar.slider(
    "Selecione o Intervalo de Vendas:",
    min_value=int(df["VENDA"].min()),
    max_value=int(df["VENDA"].max()),
    value=(int(df["VENDA"].min()), int(df["VENDA"].max()))
)

# Apply filters
df_selection = df.query(
    "GERENTE == @gerente & MES == @mes & ANO >= @ano_range[0] & ANO <= @ano_range[1] & CATEGORIA == @categoria & VENDA >= @venda_range[0] & VENDA <= @venda_range[1]"
)

# Convertendo a coluna ANO para inteiros após filtrar
df_selection['ANO'] = df_selection['ANO'].astype(int)

# ---main---#
st.title(":bar_chart: Dashboard Épico")
st.markdown("##")

# kpi #
vendas_totais = int(df_selection["VENDA"].sum())
vendas_media = int(df_selection["VENDA"].mean())
total_vendas = df_selection.shape[0]

left_column, middle_column, right_column = st.columns(3)
with left_column:
    st.subheader("Vendas Totais:")
    st.subheader(f"R$ {vendas_totais:,}")
with middle_column:
    st.subheader("Média das Vendas:")
    st.subheader(f"R$ {vendas_media:,}")
with right_column:
    st.subheader("Número de Vendas:")
    st.subheader(f"{total_vendas}")

st.markdown("---")

# Paleta de cores customizada
custom_color_map = {
    "CALÇA": "#f0e1ff",  
    "CAMISA": "#e2c8ff",  
    "RELÓGIO": "#d5b0ff",  
    "TÊNIS": "#c999ff"  
}

# Vendas por Categoria (bar chart)
vendas_por_cat = df_selection.groupby(by=["CATEGORIA"])[["VENDA"]].sum().sort_values(by="VENDA")
fig_prod_vendas = px.bar(
    vendas_por_cat,
    x="VENDA",
    y=vendas_por_cat.index,
    orientation="h",
    title="<b>Vendas por Categoria</b>",
    color=vendas_por_cat.index,
    color_discrete_map=custom_color_map,
    template="plotly_dark",
    labels={"VENDA": "Vendas", "CATEGORIA": "Categoria"},
    hover_data={"VENDA": ":.2f"}
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
    color_discrete_sequence=["#bc81ff"] * len(vendas_por_ano),  # 730943
    template="plotly_dark",
    labels={"ANO": "Ano", "VENDA": "Vendas"},
    hover_data={"VENDA": ":.2f"}
)

fig_ano_vendas.update_layout(
    xaxis=dict(tickmode="linear"),
    plot_bgcolor="rgba(0,0,0,0)",
    yaxis=dict(showgrid=False, tickformat="d"),  # Formatação do eixo y para inteiros
)

# Vendas por Categoria e Ano (bar chart de stack)
vendas_por_cat_ano = df_selection.groupby(by=["ANO", "CATEGORIA"])[["VENDA"]].sum().reset_index()

fig_cat_ano_vendas = px.bar(
    vendas_por_cat_ano,
    x="ANO",
    y="VENDA",
    color="CATEGORIA",
    title="<b>Vendas por Categoria e Ano</b>",
    template="plotly_dark",
    barmode="stack",
    color_discrete_map=custom_color_map,
    labels={"ANO": "Ano", "VENDA": "Vendas", "CATEGORIA": "Categoria"},
    hover_data={"VENDA": ":.2f"}
)

fig_cat_ano_vendas.update_layout(
    plot_bgcolor="rgba(0,0,0,0)",
    yaxis=dict(showgrid=False),
    xaxis=dict(tickmode='linear', tick0=2022, dtick=1),  # Exibição dos anos em decimal (fixed)
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
    color_discrete_sequence=["#ae69ff"] * len(vendas_por_gerente),  # 75-5
    template="plotly_dark",
    labels={"GERENTE": "Gerente", "VENDA": "Vendas"},
    hover_data={"VENDA": ":.2f"}
)

fig_gerente_vendas.update_layout(
    plot_bgcolor="rgba(0,0,0,0)",
    yaxis=dict(showgrid=False),
)

left_column, right_column = st.columns(2)
left_column.plotly_chart(fig_cat_ano_vendas, use_container_width=True)
right_column.plotly_chart(fig_gerente_vendas, use_container_width=True)

# Botão de Downloadd
st.markdown("---")
st.subheader("Download da Base de Dados Filtrada")
buffer = BytesIO()
with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
    df_selection.to_excel(writer, sheet_name='Dados Filtrados', index=False)
    worksheet = writer.sheets['Dados Filtrados']
    for col_num, value in enumerate(df_selection.columns.values):
        if value == "DIA":
            worksheet.set_column(col_num, col_num, 15) #coluna peq/fix
buffer.seek(0)
st.download_button(
    label="Baixar Dados Filtrados",
    data=buffer,
    file_name="dados_filtrados.xlsx",
    mime="application/vnd.ms-excel"
)

# ---Esconder Streamlit.Style ---
esconder_st_style = """
            <style>
            #MainMenu {visibility: hidden;}
            footer {visibility: hidden;}
            header {visibility: hidden;}
            </style>
            """
st.markdown(esconder_st_style, unsafe_allow_html=True)
