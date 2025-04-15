import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.colors as pc
from io import BytesIO

# ========== CONFIGURAÇÕES ==========
st.set_page_config(layout="wide")
st.title("📊 Relatório Interativo de ICMS e Análise Contábil")

# ========== PLANILHAS ==========
caminho_icms = "notas_processadas1.xlsx"
caminho_contab = "Contabilidade.xlsx"

# ========== LEITURA ICMS ==========
entradas = pd.read_excel(caminho_icms, sheet_name="Todas Entradas", skiprows=1)
entradas = entradas.loc[:, ~entradas.columns.to_series().isna()]
entradas.columns = [str(col).strip() for col in entradas.columns]
entradas = entradas.loc[:, ~entradas.columns.str.contains("Unnamed|^\\d+$", na=False)]
saidas = pd.read_excel(caminho_icms, sheet_name="Todas Saídas")
entradas.columns = entradas.columns.str.strip()
saidas.columns = saidas.columns.str.strip()
entradas['Mês'] = pd.to_datetime(entradas['Mês'], errors='coerce')
saidas['Mês'] = pd.to_datetime(saidas['Mês'], errors='coerce')
for df in [entradas, saidas]:
    for col in ['Valor ICMS', 'Valor Total', 'Alíquota ICMS']:
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

# ========== LEITURA PLANILHA CONTABILIDADE ==========
caixa = pd.read_excel(caminho_contab, sheet_name="Caixa")
pis = pd.read_excel(caminho_contab, sheet_name="PIS")
cofins = pd.read_excel(caminho_contab, sheet_name="COFINS")
dre = pd.read_excel(caminho_contab, sheet_name="DRE 1º Trimestre")

# Padronizar colunas caixa
caixa.columns = caixa.columns.str.strip()
caixa['Data'] = pd.to_datetime(caixa['Data'], errors='coerce')
caixa['Mês'] = caixa['Data'].dt.to_period("M")

# ========== FILTROS ==========
st.sidebar.header("🎛️ Filtros")
periodos = {
    "Janeiro/2025": [1],
    "Fevereiro/2025": [2],
    "Março/2025": [3],
    "1º Trimestre/2025": [1, 2, 3]
}
filtro_periodo = st.sidebar.selectbox("Selecione o período:", list(periodos.keys()))
filtro_aba = st.sidebar.selectbox("Tipo de Análise:", [
    "Mapa por UF",
    "Comparativo de Crédito x Débito",
    "Apuração com Crédito Acumulado",
    "Relatórios Detalhados",
    "Contabilidade e Caixa"
])
meses_filtrados = periodos[filtro_periodo]

# ========== MAPAS DE CORES ==========
ufs = sorted(set(entradas['UF do Emitente'].dropna().unique().tolist() + saidas['UF do Destinatário'].dropna().unique().tolist()))
palette = pc.qualitative.Alphabet
uf_cores = {uf: palette[i % len(palette)] for i, uf in enumerate(ufs)}

aliq_cores = {
    0: '#636EFA', 4: '#EF553B', 7: '#00CC96', 12: '#AB63FA', 19: '#FFA15A'
}

# ========== NOVA ABA: CONTABILIDADE E CAIXA ==========
if filtro_aba == "Contabilidade e Caixa":
    st.header("📘 Contabilidade e Caixa")

    caixa_filtrada = caixa[caixa['Data'].dt.month.isin(meses_filtrados)]

    receita_total = caixa_filtrada[caixa_filtrada['Tipo'] == 'Saída']['Valor'].sum()
    despesa_total = caixa_filtrada[caixa_filtrada['Tipo'] == 'Entrada']['Valor'].sum()
    saldo_final = receita_total - despesa_total
    margem_lucro = (saldo_final / receita_total) * 100 if receita_total > 0 else 0

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("💵 Receita Total", f"R$ {receita_total:,.2f}")
    col2.metric("📤 Despesas Totais", f"R$ {despesa_total:,.2f}")
    col3.metric("📌 Saldo Final", f"R$ {saldo_final:,.2f}")
    col4.metric("📈 Margem de Lucro", f"{margem_lucro:.2f}%")

    st.subheader("📊 Gráfico Receitas vs Despesas")
    graf_df = caixa_filtrada.copy()
    graf_df['Mês'] = graf_df['Data'].dt.to_period('M')
    graf_df_group = graf_df.groupby(['Mês', 'Tipo'])['Valor'].sum().reset_index()
    fig1 = px.bar(graf_df_group, x='Mês', y='Valor', color='Tipo', barmode='group', text_auto='.2s')
    st.plotly_chart(fig1, use_container_width=True)

    st.subheader("📈 Evolução do Saldo Acumulado")
    caixa_filtrada = caixa_filtrada.sort_values('Data')
    caixa_filtrada['Saldo'] = caixa_filtrada.apply(lambda row: row['Valor'] if row['Tipo'] == 'Saída' else -row['Valor'], axis=1).cumsum()
    fig2 = px.line(caixa_filtrada, x='Data', y='Saldo', title='Evolução do Saldo Acumulado')
    st.plotly_chart(fig2, use_container_width=True)

    st.subheader("📉 Despesas por Categoria")
    categoria_pizza = caixa_filtrada[caixa_filtrada['Tipo'] == 'Entrada'].groupby('Categoria')['Valor'].sum().reset_index()
    fig3 = px.pie(categoria_pizza, names='Categoria', values='Valor', title='Distribuição das Despesas por Categoria')
    fig3.update_traces(textinfo='label+percent')
    st.plotly_chart(fig3, use_container_width=True)

    st.subheader("🧾 Tabela Completa com Filtros")
    st.dataframe(caixa_filtrada[['Data', 'Categoria', 'Tipo', 'Valor']], use_container_width=True)
