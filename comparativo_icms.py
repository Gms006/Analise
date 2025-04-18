import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.colors as pc
from io import BytesIO
import base64
import unicodedata

# =========================
# 1. FONT AWESOME & CSS GLOBAL
# =========================
st.set_page_config(
    layout="wide",
    page_title="Relatório GH Sistemas",
    page_icon="📊",
    initial_sidebar_state="expanded"
)

st.markdown("""
<!-- Font Awesome -->
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0-beta3/css/all.min.css">
<style>
body, .stApp {
    background-color: #1E2B3D !important;
}
h1, h3, h4, h5, h6 {
    color: #C89D4A !important;
    font-family: 'Segoe UI', sans-serif;
}
p, li, th, td, label, .markdown-text-container {
    color: #CCCCCC !important;
    font-family: 'Segoe UI', sans-serif;
}
hr {
    border: 1px solid #C89D4A !important;
}
button[kind="primary"], .stButton>button {
    background-color: #C89D4A !important;
    color: #1E2B3D !important;
    border: none !important;
    border-radius: 5px !important;
    font-weight: bold;
}
.stDataFrame, .stTable {
    background-color: #22304A !important;
    border-radius: 8px !important;
}
.stExpanderHeader {
    color: #C89D4A !important;
}
.info-bloco {
    background-color: #2D3B50;
    border-left: 5px solid #C89D4A;
    border-radius: 8px;
    padding: 12px 18px;
    margin-bottom: 18px;
}
.info-bloco i {
    color: #C89D4A;
    margin-right: 8px;
}
.rodape {
    margin-top: 40px;
    padding: 18px 0 0 0;
    text-align: center;
    color: #C89D4A;
    font-size: 16px;
    border-top: 1px solid #C89D4A;
    letter-spacing: 1px;
}
</style>
""", unsafe_allow_html=True)

# =========================
# 2. FUNÇÃO DE PLANO DE FUNDO
# =========================
def set_background(path):
    with open(path, "rb") as f:
        encoded = base64.b64encode(f.read()).decode()
    st.markdown(
        f"""
        <style>
        .stApp {{
            background-image: url("data:image/png;base64,{encoded}");
            background-position: top right;
            background-repeat: no-repeat;
            background-size: 300px;
            background-attachment: fixed;
            background-color: #1E2B3D;
        }}
        </style>
        """,
        unsafe_allow_html=True
    )
set_background("logo.png")

# =========================
# 3. SIDEBAR: LOGO E IDENTIDADE
# =========================
with st.sidebar:
    st.image("logo.png", use_column_width=True)
    st.markdown(
        "<h3 style='text-align: center; color: #C89D4A;'>Neto Contabilidade</h3>",
        unsafe_allow_html=True
    )

# =========================
# 4. TÍTULO PRINCIPAL UNIFORMIZADO
# =========================
st.markdown("""
<h1 style='text-align: center; font-size: 42px;'>
    <i class="fas fa-chart-bar"></i> Relatório Gerencial - Neto Contabilidade
</h1>
<hr>
""", unsafe_allow_html=True)

# =========================
# 5. FUNÇÃO REUTILIZÁVEL: BLOCO VISUAL
# =========================
def bloco_visual(titulo, icone, descricao):
    st.markdown(f"""
    <div class="info-bloco">
        <h3 style="margin:0;">
            <i class="fas fa-{icone}"></i> {titulo}
        </h3>
        <p style="margin:5px 0 0; font-size:15px;">{descricao}</p>
    </div>
    """, unsafe_allow_html=True)

# =========================
# 6. LEITURA DE DADOS
# =========================
caminho_planilha = "notas_processadas1.xlsx"
entradas = pd.read_excel(caminho_planilha, sheet_name="Todas Entradas", skiprows=1)
entradas = entradas.loc[:, ~entradas.columns.to_series().isna()]
entradas.columns = [str(col).strip() for col in entradas.columns]
entradas = entradas.loc[:, ~entradas.columns.str.contains("Unnamed|^\\d+$", na=False)]
saidas = pd.read_excel(caminho_planilha, sheet_name="Todas Saídas")
entradas.columns = entradas.columns.str.strip()
saidas.columns = saidas.columns.str.strip()
entradas['Mês'] = pd.to_datetime(entradas['Mês'], errors='coerce')
saidas['Mês'] = pd.to_datetime(saidas['Mês'], errors='coerce')
for df in [entradas, saidas]:
    for col in ['Valor ICMS', 'Valor Total', 'Alíquota ICMS']:
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

with pd.ExcelFile("Contabilidade.xlsx") as xls:
    if "Caixa" in xls.sheet_names:
        caixa_df = pd.read_excel(xls, sheet_name="Caixa")
    else:
        st.warning("Aba 'Caixa' não encontrada.")

try:
    piscofins_df = pd.read_excel("Contabilidade.xlsx", sheet_name="PISCOFINS")
    dre_df = pd.read_excel("Contabilidade.xlsx", sheet_name="DRE 1º Trimestre")
except KeyError as e:
    st.error(f"Erro: Aba não encontrada - {e}")

@st.cache_data
def carregar_dados():
    return entradas, saidas, caixa_df, piscofins_df, dre_df

# =========================
# 7. FUNÇÕES AUXILIARES
# =========================
def calcular_saldo_com_acumulado(df, meses_filtrados):
    df = df.sort_values("Data").copy()
    df["Mês"] = df["Data"].dt.month
    df["Ano"] = df["Data"].dt.year
    df["Valor Líquido"] = df["Entradas"] - df["Saídas"]
    primeiro_mes = min(meses_filtrados)
    saldo_anterior = df[df["Mês"] < primeiro_mes]["Valor Líquido"].sum()
    df_filtrado = df[df["Mês"].isin(meses_filtrados)].copy()
    df_filtrado["Saldo Acumulado"] = df_filtrado["Valor Líquido"].cumsum() + saldo_anterior
    return df_filtrado

def plotar_saldo_mensal(caixa_df, meses_selecionados):
    df = caixa_df.copy()
    df['Data'] = pd.to_datetime(df['Data'], errors='coerce')
    df = df.sort_values('Data').reset_index(drop=True)
    df['Mês'] = df['Data'].dt.month
    df['Ano'] = df['Data'].dt.year
    df['Valor Líquido'] = df['Entradas'] - df['Saídas']
    pontos = []
    for mes in meses_selecionados:
        df_mes = df[df['Mês'] == mes]
        if df_mes.empty:
            continue
        ano = df_mes['Ano'].dropna().iloc[0] if not df_mes['Ano'].dropna().empty else None
        if ano is None:
            continue
        try:
            data_limite = pd.Timestamp(f"{int(ano)}-{mes:02d}-01")
        except Exception as e:
            st.warning(f"Erro ao gerar data para o mês {mes} e ano {ano}: {e}")
            continue
        df_ant = df[df['Data'] < data_limite]
        saldo_ant = df_ant['Valor Líquido'].sum() if not df_ant.empty else 0
        if len(meses_selecionados) > 1:
            saldo_fim = df_mes['Valor Líquido'].cumsum().iloc[-1] + saldo_ant
            data_fim = df_mes['Data'].iloc[-1]
            pontos.append({'Data': data_fim, 'Saldo Acumulado': saldo_fim, 'Mês': mes})
        else:
            data_ant = df_ant['Data'].iloc[-1] if not df_ant.empty else (data_limite - pd.Timedelta(days=1))
            pontos.append({'Data': data_ant, 'Saldo Acumulado': saldo_ant, 'Mês': mes})
            data_15 = pd.Timestamp(f"{int(ano)}-{mes:02d}-15")
            df_mes_15 = df_mes[df_mes['Data'] <= data_15]
            if not df_mes_15.empty:
                saldo_15 = df_mes_15['Valor Líquido'].cumsum().iloc[-1] + saldo_ant
                pontos.append({'Data': data_15, 'Saldo Acumulado': saldo_15, 'Mês': mes})
            saldo_fim = df_mes['Valor Líquido'].cumsum().iloc[-1] + saldo_ant
            data_fim = df_mes['Data'].iloc[-1]
            pontos.append({'Data': data_fim, 'Saldo Acumulado': saldo_fim, 'Mês': mes})
    df_pontos = pd.DataFrame(pontos)
    fig = px.line(df_pontos, x="Data", y="Saldo Acumulado", markers=True, title="Evolução  Saldo de caixa ")
    st.plotly_chart(fig, use_container_width=True)

def normalize_str(s):
    # Remove acentos e deixa minúsculo
    return ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn').lower()

def get_dre_val(desc, flexible=False):
    if flexible:
        # Busca flexível: ignora acentos, maiúsculas/minúsculas e aceita parte do texto
        desc_norm = normalize_str(desc)
        dre_df['desc_norm'] = dre_df['Descrição'].astype(str).apply(normalize_str)
        row = dre_df[dre_df['desc_norm'].str.contains(desc_norm)]
    else:
        row = dre_df[dre_df['Descrição'].str.strip().str.upper() == desc.strip().upper()]
    if not row.empty:
        val = str(row['Saldo'].values[0])
        val = val.replace("R$", "").replace(" ", "")
        if "," in val:
            val_float = float(val.replace(".", "").replace(",", "."))
        else:
            try:
                val_float = float(val)
            except Exception:
                val_float = 0.0
        return f"R$ {val_float:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return "R$ 0,00"

# =========================
# 8. FILTROS DINÂMICOS
# =========================
st.sidebar.markdown("""
<h3 style="color:#C89D4A; margin-bottom: 0;">
    <i class="fas fa-sliders-h"></i> Filtros
</h3>
""", unsafe_allow_html=True)
periodos = {
    "Janeiro/2025": [1],
    "Fevereiro/2025": [2],
    "Março/2025": [3],
    "1º Trimestre/2025": [1, 2, 3]
}
filtro_periodo = st.sidebar.selectbox(
    "📅 Período:",
    ["Janeiro/2025", "Fevereiro/2025", "Março/2025", "1º Trimestre/2025"],
    key="periodo"
)
meses_filtrados = periodos[filtro_periodo]
entradas_filtradas = entradas[entradas['Mês'].dt.month.isin(meses_filtrados)]
saidas_filtradas = saidas[saidas['Mês'].dt.month.isin(meses_filtrados)]

# =========================
# 9. DEMONSTRATIVO DO PERÍODO FILTRADO
# =========================
creditos = entradas.groupby(entradas['Mês'].dt.to_period('M'))['Valor ICMS'].sum().reset_index(name='ICMS Crédito')
debitos = saidas.groupby(saidas['Mês'].dt.to_period('M'))['Valor ICMS'].sum().reset_index(name='ICMS Débito')
comparativo = pd.merge(creditos, debitos, on='Mês', how='outer').fillna(0).sort_values(by='Mês')
comparativo['Crédito Acumulado'] = 0.0
comparativo['ICMS Apurado Corrigido'] = 0.0
credito_acumulado = 0
for i, row in comparativo.iterrows():
    credito_total = row['ICMS Crédito'] + credito_acumulado
    apurado = row['ICMS Débito'] - credito_total
    comparativo.at[i, 'Crédito Acumulado'] = credito_acumulado
    comparativo.at[i, 'ICMS Apurado Corrigido'] = apurado
    credito_acumulado = max(0, -apurado)
comparativo['Mês'] = comparativo['Mês'].astype(str)
comparativo_filtrado = comparativo[comparativo['Mês'].apply(lambda x: int(x[5:7]) in meses_filtrados)]

# =========================
# 10. MAPA DE CORES
# =========================
ufs = sorted(set(entradas['UF do Emitente'].dropna().unique().tolist() + saidas['UF do Destinatário'].dropna().unique().tolist()))
palette = pc.qualitative.Alphabet
uf_cores = {uf: palette[i % len(palette)] for i, uf in enumerate(ufs)}
aliq_cores = {0: '#636EFA', 4: '#EF553B', 7: '#00CC96', 12: '#AB63FA', 19: '#FFA15A'}

# =========================
# =========================
# 11. GRÁFICOS E RELATÓRIOS
# === CATEGORIAS DE RELATÓRIOS ===
aba = st.sidebar.radio(
    "📁 Tipo de Relatório:",
    ["📂 Fiscal", "📊 Contábil"]
)

# === OPÇÕES DINÂMICAS DEPENDENDO DA CATEGORIA ===
if aba == "📂 Fiscal":
    filtro_grafico = st.sidebar.selectbox(
        "📄 Relatórios Fiscais:",
        [
            "Mapa por UF",
            "Comparativo de Crédito x Débito",
            "Apuração com Crédito Acumulado",
            "Relatórios Detalhados"
        ]
    )
else:
    filtro_grafico = st.sidebar.selectbox(
        "📘 Relatórios Contábeis:",
        [
            "📘 Contabilidade e Caixa",
            "📗 PIS e COFINS",
            "📘 DRE Trimestral",
            "📑 Tabelas Contabilidade"
        ]
    )

if filtro_grafico == "Mapa por UF":
    bloco_visual(
        "Distribuição de Compras e Vendas por Estado (UF)",
        "map-marker-alt",
        "Visualize o volume total de compras e vendas por unidade federativa, tanto em barras quanto em pizza. <i class='fas fa-info-circle'></i>"
    )
    col1, col2 = st.columns(2)
    with col1:
        uf_compras = entradas_filtradas.groupby('UF do Emitente')['Valor Total'].sum().reset_index()
        fig = px.bar(uf_compras, x='UF do Emitente', y='Valor Total', text_auto='.2s', title="Compras por UF (Volume Total)")
        st.plotly_chart(fig, use_container_width=True)
        fig_pie = px.pie(uf_compras, names='UF do Emitente', values='Valor Total', title='Distribuição % por UF - Compras',
                         color='UF do Emitente', color_discrete_map=uf_cores, hole=0.3)
        fig_pie.update_traces(textinfo='label+value')
        st.plotly_chart(fig_pie, use_container_width=True)
    with col2:
        uf_vendas = saidas_filtradas.groupby('UF do Destinatário')['Valor Total'].sum().reset_index()
        fig = px.bar(uf_vendas, x='UF do Destinatário', y='Valor Total', text_auto='.2s', title="Saídas por UF (Volume Total)")
        st.plotly_chart(fig, use_container_width=True)
        fig_pie2 = px.pie(uf_vendas, names='UF do Destinatário', values='Valor Total', title='Distribuição % por UF - Faturamento',
                          color='UF do Destinatário', color_discrete_map=uf_cores, hole=0.3)
        fig_pie2.update_traces(textinfo='label+value')
        st.plotly_chart(fig_pie2, use_container_width=True)

elif filtro_grafico == "Comparativo de Crédito x Débito":
    bloco_visual(
        "Comparativo Mensal de ICMS",
        "balance-scale",
        "Compare créditos e débitos de ICMS mês a mês, além da distribuição por faixa de alíquota. <i class='fas fa-info-circle'></i>"
    )
    df_bar = comparativo_filtrado.melt(id_vars='Mês', value_vars=['ICMS Crédito', 'ICMS Débito'])
    fig_bar = px.bar(df_bar, x='Mês', y='value', color='variable', barmode='group', text_auto='.2s')
    st.plotly_chart(fig_bar, use_container_width=True)
    bloco_visual(
        "Distribuição de ICMS por Faixa de Alíquota",
        "percent",
        "Veja como os créditos e débitos de ICMS se distribuem entre diferentes faixas de alíquota. <i class='fas fa-info-circle'></i>"
    )
    df_aliq = entradas_filtradas.copy()
    df_aliq['Aliquota'] = (df_aliq['Alíquota ICMS'] * 100).round(0).astype(int)
    df_aliq['Crédito ICMS Estimado'] = df_aliq['Valor ICMS']
    df_saida = saidas_filtradas.copy()
    df_saida['Aliquota'] = (df_saida['Alíquota ICMS'] * 100).round(0).astype(int)
    total_compras = df_aliq.groupby('Aliquota').agg({'Valor Total': 'sum', 'Crédito ICMS Estimado': 'sum'}).reset_index()
    total_debitos = df_saida.groupby('Aliquota')['Valor ICMS'].sum().reset_index(name='Débito ICMS')
    df_final = pd.merge(total_compras, total_debitos, on='Aliquota', how='outer').fillna(0)
    df_dual = df_final.melt(id_vars='Aliquota', value_vars=['Valor Total', 'Crédito ICMS Estimado', 'Débito ICMS'],
                            var_name='Tipo', value_name='Valor')
    fig_aliq_bar = px.bar(df_dual, x='Aliquota', y='Valor', color='Tipo', barmode='group', text_auto='.2s',
                          title="Comparativo por Alíquota: Compras, Crédito e Débito")
    fig_aliq_bar.update_layout(xaxis=dict(tickmode='array', tickvals=[0, 4, 7, 12, 19]))
    st.plotly_chart(fig_aliq_bar, use_container_width=True)
    fig_pie_credito = px.pie(df_final, names='Aliquota', values='Crédito ICMS Estimado', title='% de Crédito por Alíquota',
                             color='Aliquota', color_discrete_map=aliq_cores, hole=0.3)
    fig_pie_credito.update_traces(textinfo='label+value')
    fig_pie_debito = px.pie(df_final, names='Aliquota', values='Débito ICMS', title='% de Débito por Alíquota',
                            color='Aliquota', color_discrete_map=aliq_cores, hole=0.3)
    fig_pie_debito.update_traces(textinfo='label+value')
    col3, col4 = st.columns(2)
    with col3:
        st.plotly_chart(fig_pie_credito, use_container_width=True)
    with col4:
        st.plotly_chart(fig_pie_debito, use_container_width=True)

elif filtro_grafico == "Relatórios Detalhados":
    bloco_visual(
        "Dados Fiscais Detalhados (.xlsx)",
        "file-excel",
        "Visualize e baixe todas as notas fiscais e apurações do período selecionado. <i class='fas fa-info-circle'></i>"
    )
    st.markdown("<h3><i class='fas fa-download'></i> Notas Fiscais de Entrada</h3>", unsafe_allow_html=True)
    st.dataframe(entradas_filtradas.fillna("").astype(str), use_container_width=True)
    st.markdown("<h3><i class='fas fa-upload'></i> Notas Fiscais de Saída</h3>", unsafe_allow_html=True)
    st.dataframe(saidas_filtradas.fillna("").astype(str), use_container_width=True)
    st.markdown("<h3><i class='fas fa-balance-scale'></i> Comparativo de Crédito x Débito com Crédito Acumulado</h3>", unsafe_allow_html=True)
    st.dataframe(comparativo_filtrado.style.format({
        'ICMS Crédito': 'R$ {:,.2f}',
        'ICMS Débito': 'R$ {:,.2f}',
        'Crédito Acumulado': 'R$ {:,.2f}',
        'ICMS Apurado Corrigido': 'R$ {:,.2f}'
    }), use_container_width=True)
    def to_excel():
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            entradas_filtradas.to_excel(writer, sheet_name="Entradas", index=False)
            saidas_filtradas.to_excel(writer, sheet_name="Saídas", index=False)
            comparativo_filtrado.to_excel(writer, sheet_name="Apuracao", index=False)
        processed_data = output.getvalue()
        return processed_data
    excel_bytes = to_excel()
    st.download_button(
        label="Baixar Relatórios Completos (.xlsx)",
        data=excel_bytes,
        file_name="Relatorio_ICMS_Completo.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

elif filtro_grafico == "📘 Contabilidade e Caixa":
    bloco_visual(
        "Caixa Contábil no Período",
        "cash-register",
        "Acompanhe entradas, saídas e saldo acumulado do caixa contábil. <i class='fas fa-info-circle'></i>"
    )
    caixa_df['Entradas'] = pd.to_numeric(caixa_df['Entradas'], errors='coerce').fillna(0)
    caixa_df['Saídas'] = pd.to_numeric(caixa_df['Saídas'], errors='coerce').fillna(0)
    caixa_df['Entrada'] = caixa_df['Entradas']
    caixa_df['Saída'] = -caixa_df['Saídas']
    caixa_df['Valor Líquido'] = caixa_df['Entrada'] + caixa_df['Saída']
    caixa_df['Data'] = pd.to_datetime(caixa_df['Data'], errors='coerce')
    caixa_df['Mês'] = caixa_df['Data'].dt.month
    caixa_df['Ano'] = caixa_df['Data'].dt.year
    meses_selecionados = periodos[filtro_periodo]
    caixa_ordenado = caixa_df.sort_values('Data').copy()
    caixa_filtrado = calcular_saldo_com_acumulado(caixa_df, meses_selecionados)
    receita_total = caixa_filtrado['Entradas'].sum()
    despesa_total = caixa_filtrado['Saídas'].sum()
    saldo_final = receita_total - despesa_total
    margem = (saldo_final / receita_total * 100) if receita_total != 0 else 0
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Total de Entradas", f"R$ {receita_total:,.2f}")
    col2.metric("Total de Saídas", f"R$ {despesa_total:,.2f}")
    col3.metric("Saldo Final", f"R$ {saldo_final:,.2f}")
    col4.metric("Margem (%)", f"{margem:.2f}%")
    df_bar = caixa_filtrado[['Entradas', 'Saídas']].sum().reset_index()
    df_bar.columns = ['Tipo', 'Valor']
    fig_bar = px.bar(df_bar, x='Tipo', y='Valor', text_auto='.2s', color='Tipo', title="Entradas x Saídas no Período")
    st.plotly_chart(fig_bar, use_container_width=True)
    plotar_saldo_mensal(caixa_df, meses_selecionados)

elif filtro_grafico == "📗 PIS e COFINS":
    bloco_visual(
        "Situação Fiscal de PIS e COFINS",
        "file-invoice-dollar",
        "Veja créditos, débitos e saldo acumulado de PIS e COFINS no período. <i class='fas fa-info-circle'></i>"
    )
    ordem_meses = {"Janeiro": 1, "Fevereiro": 2, "Março": 3}
    meses_filtro = {
        "Janeiro/2025": ["Janeiro"],
        "Fevereiro/2025": ["Fevereiro"],
        "Março/2025": ["Março"],
        "1º Trimestre/2025": ["Janeiro", "Fevereiro", "Março"]
    }
    meses_selecionados = meses_filtro[filtro_periodo]
    piscofins_ordenado = piscofins_df.copy()
    piscofins_ordenado['Ordem'] = piscofins_ordenado['Mês'].map(ordem_meses)
    piscofins_ordenado = piscofins_ordenado.sort_values(by="Ordem")
    piscofins_filtrado = piscofins_ordenado[piscofins_ordenado['Mês'].isin(meses_selecionados)]
    credito_total = piscofins_filtrado['Crédito'].sum()
    debito_total = piscofins_filtrado['Débito'].sum()
    saldo_final = credito_total - debito_total
    col1, col2, col3 = st.columns(3)
    col1.metric("Total Créditos", f"R$ {credito_total:,.2f}")
    col2.metric("Total Débitos", f"R$ {debito_total:,.2f}")
    col3.metric("Saldo Final", f"R$ {saldo_final:,.2f}")
    df_bar = pd.DataFrame({
        'Tipo': ['Crédito', 'Débito'],
        'Valor': [credito_total, debito_total]
    })
    fig_bar = px.bar(df_bar, x='Tipo', y='Valor', text_auto='.2s', color='Tipo', title="Créditos x Débitos no Período")
    st.plotly_chart(fig_bar, use_container_width=True)
    pontos = []
    if len(meses_selecionados) == 1:
        mes_nome = meses_selecionados[0]
        mes_num = ordem_meses[mes_nome]
        saldo_anterior = piscofins_ordenado[piscofins_ordenado['Ordem'] < mes_num]['Saldo']
        saldo_anterior = saldo_anterior.iloc[-1] if not saldo_anterior.empty else 0
        pontos.append({'Mês': f"{mes_nome} - Início", 'Saldo': -saldo_anterior})
        saldo_fim = piscofins_ordenado[piscofins_ordenado['Ordem'] == mes_num]['Saldo']
        saldo_fim = saldo_fim.iloc[-1] if not saldo_fim.empty else saldo_anterior
        pontos.append({'Mês': f"{mes_nome} - Fim", 'Saldo': -saldo_fim})
    else:
        for mes_nome in meses_selecionados:
            saldo_fim = piscofins_ordenado[piscofins_ordenado['Mês'] == mes_nome]['Saldo']
            if saldo_fim.empty:
                continue
            pontos.append({'Mês': mes_nome, 'Saldo': -saldo_fim.iloc[-1]})
    df_pontos = pd.DataFrame(pontos)
    fig_saldo_pis = px.line(
    df_pontos, x='Mês', y='Saldo',
        title='Evolução do Saldo Acumulado - PIS e COFINS'
    )
    st.plotly_chart(fig_saldo_pis, use_container_width=True)

elif filtro_grafico == "📘 DRE Trimestral":
    bloco_visual(
        "Demonstração do Resultado do Exercício (DRE)",
        "file-contract",
        "Resumo do resultado do exercício, com receitas, deduções, custos, despesas e lucro/prejuízo final. <i class='fas fa-info-circle'></i>"
    )

    receita_bruta = get_dre_val("VENDA DE MERCADORIAS A VISTA", flexible=True)
    deducoes = get_dre_val("(-) Deduções das Receitas Operacionais", flexible=True)
    receita_liquida = get_dre_val("Total das Receitas Operacionais Líquidas", flexible=True)
    custo_mercadorias = get_dre_val("Custos das Mercadorias Vendidas", flexible=True)
    lucro_bruto = get_dre_val("Lucro e/ou Prejuízo Operacional Bruto", flexible=True)
    despesas = get_dre_val("Despesas Administrativas", flexible=True)

    resultado_liquido = get_dre_val("Resultado Líquido do Exercício", flexible=True)
    if resultado_liquido == "R$ 0,00":
        resultado_liquido = get_dre_val("Lucro/Prejuízo Líquido do Exercício", flexible=True)
    if resultado_liquido == "R$ 0,00":
        resultado_liquido = get_dre_val("Lucro/Prejuízo Líquido Antes da CSLL", flexible=True)

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Receita Bruta", receita_bruta)
    col2.metric("Receita Líquida", receita_liquida)
    col3.metric("Lucro Bruto", lucro_bruto)
    col4.metric("Resultado Líquido", resultado_liquido)

    st.markdown(f"""
        <div style="background:#22304A; border-radius:10px; padding:24px; margin-top:20px;">
          <ul style="list-style:none; padding-left:0; font-size:18px;">
            <li style="padding:8px 0; color:#C89D4A;">
              <i class="fas fa-cash-register"></i> <b>Receita Operacional Bruta:</b>
              <span style="color:#00FFAA;">{receita_bruta}</span>
            </li>
            <li style="padding:8px 0; color:#C89D4A; margin-left:24px;">
              <i class="fas fa-minus-circle"></i> (-) Deduções:
              <span style="color:#FFA500;">{deducoes}</span>
            </li>
            <hr style="border: 1px dashed #C89D4A; margin: 12px 0;">
            <li style="padding:8px 0; color:#C89D4A;">
              <i class="fas fa-file-invoice-dollar"></i> <b>Receita Líquida:</b>
              <span style="color:#00FFAA;">{receita_liquida}</span>
            </li>
            <li style="padding:8px 0; color:#C89D4A; margin-left:24px;">
              <i class="fas fa-box"></i> (-) Custos das Mercadorias Vendidas:
              <span style="color:#FFA500;">{custo_mercadorias}</span>
            </li>
            <hr style="border: 1px dashed #C89D4A; margin: 12px 0;">
            <li style="padding:8px 0; color:#C89D4A;">
              <i class="fas fa-chart-line"></i> <b>Lucro Bruto:</b>
              <span style="color:{'#00FFAA' if lucro_bruto.replace('R$', '').replace('.', '').replace(',', '.').strip() and float(lucro_bruto.replace('R$', '').replace('.', '').replace(',', '.')) >= 0 else '#FF5555'};">{lucro_bruto}</span>
            </li>
            <li style="padding:8px 0; color:#C89D4A; margin-left:24px;">
              <i class="fas fa-money-bill-wave"></i> (-) Despesas Administrativas:
              <span style="color:#FFA500;">{despesas}</span>
            </li>
            <hr style="border: 1px dashed #C89D4A; margin: 12px 0;">
            <li style="padding:8px 0; color:#C89D4A;">
              <i class="fas fa-coins"></i> <b>Resultado Líquido do Exercício:</b>
              <span style="color:{'#00FFAA' if resultado_liquido.replace('R$', '').replace('.', '').replace(',', '.').strip() and float(resultado_liquido.replace('R$', '').replace('.', '').replace(',', '.')) >= 0 else '#FF5555'};">{resultado_liquido}</span>
            </li>
          </ul>
        </div>
        """, unsafe_allow_html=True)

# =========================
# 13. RODAPÉ INSTITUCIONAL
# =========================
st.markdown("""
<div class="rodape">
    <i class="fas fa-building"></i> Neto Contabilidade &nbsp;|&nbsp; Powered by GH Sistemas
</div>
""", unsafe_allow_html=True)