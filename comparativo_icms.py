import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.colors as pc
from io import BytesIO

# ========== CONFIGURAÇÕES ==========
st.set_page_config(layout="wide")
st.title("📊 Relatório Trimestral GH Sistemas")
caminho_planilha = "notas_processadas1.xlsx"

# ========== LEITURA ==========
# LEITURA COM TRATAMENTO DE COLUNAS INVÁLIDAS
entradas = pd.read_excel(caminho_planilha, sheet_name="Todas Entradas", skiprows=1)
entradas = entradas.loc[:, ~entradas.columns.to_series().isna()]  # Remove colunas sem nome
entradas.columns = [str(col).strip() for col in entradas.columns]  # Remove espaços em branco nos nomes das colunas
entradas = entradas.loc[:, ~entradas.columns.str.contains("Unnamed|^\\d+$", na=False)]  # Remove colunas "Unnamed" ou com nomes numéricos

saidas = pd.read_excel(caminho_planilha, sheet_name="Todas Saídas")

# LIMPEZA E FORMATOS
entradas.columns = entradas.columns.str.strip()
saidas.columns = saidas.columns.str.strip()
entradas['Mês'] = pd.to_datetime(entradas['Mês'], errors='coerce')
saidas['Mês'] = pd.to_datetime(saidas['Mês'], errors='coerce')

# CONVERSÕES
for df in [entradas, saidas]:
    for col in ['Valor ICMS', 'Valor Total', 'Alíquota ICMS']:
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

# ========== LEITURA DA PLANILHA CONTABILIDADE ==========
planilha_contabil = pd.read_excel("Contabilidade.xlsx", sheet_name=None)
try:
    caixa_df = planilha_contabil['Caixa']
    piscofins_df = planilha_contabil['PISCOFINS']
    dre_df = planilha_contabil['DRE 1º Trimestre']
except KeyError as e:
    st.error(f"Erro: Aba não encontrada - {e}")

# ========== FILTROS DINÂMICOS ==========
st.sidebar.header("🎛️ Filtros")
periodos = {
    "Janeiro/2025": [1],
    "Fevereiro/2025": [2],
    "Março/2025": [3],
    "1º Trimestre/2025": [1, 2, 3]
}
filtro_periodo = st.sidebar.selectbox("Selecione o período:", list(periodos.keys()))

# Separação Fiscal x Contabilidade
aba = st.sidebar.radio("Selecione a área:", ["Fiscal", "Contabilidade"])

if aba == "Fiscal":
    filtro_grafico = st.sidebar.selectbox("Tipo de gráfico Fiscal:", [
        "Mapa por UF",
        "Comparativo de Crédito x Débito",
        "Apuração com Crédito Acumulado",
        "Relatórios Detalhados",
    ])
else:
    filtro_grafico = st.sidebar.selectbox("Tipo de gráfico Contabilidade:", [
        "📘 Contabilidade e Caixa",
        "📗 PIS e COFINS",
        "📘 DRE Trimestral"
    ])

meses_filtrados = periodos[filtro_periodo]
entradas_filtradas = entradas[entradas['Mês'].dt.month.isin(meses_filtrados)]
saidas_filtradas = saidas[saidas['Mês'].dt.month.isin(meses_filtrados)]

# ========== DEMONSTRATIVO DO PERÍODO FILTRADO ==========
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

# ========== MAPA DE CORES ==========
ufs = sorted(set(entradas['UF do Emitente'].dropna().unique().tolist() + saidas['UF do Destinatário'].dropna().unique().tolist()))
palette = pc.qualitative.Alphabet
uf_cores = {uf: palette[i % len(palette)] for i, uf in enumerate(ufs)}

aliq_cores = {
    0: '#636EFA', 4: '#EF553B', 7: '#00CC96', 12: '#AB63FA', 19: '#FFA15A'
}

# ========== GERAÇÃO DOS GRÁFICOS ==========
if filtro_grafico == "Mapa por UF":
    st.subheader("📍 Mapa de Apuração por UF")
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
    st.subheader("📊 Comparativo de Crédito x Débito")
    df_bar = comparativo_filtrado.melt(id_vars='Mês', value_vars=['ICMS Crédito', 'ICMS Débito'])
    fig_bar = px.bar(df_bar, x='Mês', y='value', color='variable', barmode='group', text_auto='.2s')
    st.plotly_chart(fig_bar, use_container_width=True)

    st.subheader("📊 Compras e Apuração por Alíquota de ICMS")
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
    st.subheader("📄 Relatórios Detalhados e Download de Tabelas")

    # Exibir Entradas
    st.subheader("📥 Entradas Filtradas")
    st.dataframe(entradas_filtradas.replace({pd.NA: "", None: "", float("nan"): ""}), use_container_width=True)

    # Exibir Saídas
    st.subheader("📤 Saídas Filtradas")
    st.dataframe(saidas_filtradas.replace({pd.NA: "", None: "", float("nan"): ""}), use_container_width=True)

    # Exibir Apuração com crédito acumulado
    st.write("### 📊 Comparativo de Crédito x Débito com Crédito Acumulado")
    st.dataframe(comparativo_filtrado.style.format({
        'ICMS Crédito': 'R$ {:,.2f}',
        'ICMS Débito': 'R$ {:,.2f}',
        'Crédito Acumulado': 'R$ {:,.2f}',
        'ICMS Apurado Corrigido': 'R$ {:,.2f}'
    }), use_container_width=True)

    # Função para gerar Excel
    def to_excel():
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            entradas_filtradas.to_excel(writer, sheet_name="Entradas", index=False)
            saidas_filtradas.to_excel(writer, sheet_name="Saídas", index=False)
            comparativo_filtrado.to_excel(writer, sheet_name="Apuracao", index=False)
        processed_data = output.getvalue()
        return processed_data

    # Botão para baixar o Excel completo
    excel_bytes = to_excel()
    st.download_button("⬇️ Baixar Relatórios Completos (.xlsx)",
                       data=excel_bytes,
                       file_name="Relatorio_ICMS_Completo.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

elif filtro_grafico == "📘 Contabilidade e Caixa":
    st.subheader("📘 Contabilidade e Caixa")

    # Tratamento das colunas para garantir sinal correto
    caixa_df['Entradas'] = pd.to_numeric(caixa_df['Entradas'], errors='coerce').fillna(0)
    caixa_df['Saídas'] = pd.to_numeric(caixa_df['Saídas'], errors='coerce').fillna(0)
    caixa_df['Entrada'] = caixa_df['Entradas']
    caixa_df['Saída'] = -caixa_df['Saídas']  # saídas negativas explicitamente
    caixa_df['Valor Líquido'] = caixa_df['Entrada'] + caixa_df['Saída']

    caixa_df['Data'] = pd.to_datetime(caixa_df['Data'], errors='coerce')
    caixa_df['Mês'] = caixa_df['Data'].dt.month
    caixa_df['Ano'] = caixa_df['Data'].dt.year

    meses_selecionados = periodos[filtro_periodo]
    caixa_filtrado = caixa_df[caixa_df['Mês'].isin(meses_selecionados)]

    # Tabela detalhada PRIMEIRO
    st.subheader("🗃️ Tabela Detalhada de Caixa")
    st.dataframe(caixa_filtrado[['Data', 'Descricao', 'Entradas', 'Saídas', 'Valor Líquido']],
                 use_container_width=True)

    caixa_resumo = caixa_filtrado.groupby('Mês').agg({
        'Entradas': 'sum',
        'Saídas': 'sum',
        'Valor Líquido': 'sum'
    }).reset_index()

    caixa_resumo['Saldo Acumulado'] = caixa_resumo['Valor Líquido'].cumsum()
    nomes_meses = {1:'Janeiro', 2:'Fevereiro', 3:'Março'}
    caixa_resumo['Mês'] = caixa_resumo['Mês'].map(nomes_meses)

    fig = px.bar(caixa_resumo, x='Mês', y=['Entradas', 'Saídas'], barmode='group',
                 title="Entradas vs Saídas Mensais")
    st.plotly_chart(fig, use_container_width=True)

    fig_saldo = px.line(
        caixa_resumo, x='Mês', y='Saldo Acumulado',
        title='Evolução Mensal do Saldo Acumulado - Caixa',
        markers=True
    )
    st.plotly_chart(fig_saldo, use_container_width=True)

    if 'Descricao' in caixa_filtrado.columns:
        categoria_resumo = caixa_filtrado.groupby('Descricao')['Valor Líquido'].sum().reset_index()
        fig_categoria = px.pie(categoria_resumo, names='Descricao', values='Valor Líquido',
                               title='Distribuição de Gastos/Receitas por Categoria')
        st.plotly_chart(fig_categoria, use_container_width=True)

    receita_total = caixa_filtrado['Entradas'].sum()
    despesa_total = caixa_filtrado['Saídas'].sum()
    saldo_final = receita_total - despesa_total
    margem = (saldo_final / receita_total * 100) if receita_total != 0 else 0

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("📈 Receita Total", f"R$ {receita_total:,.2f}")
    col2.metric("📉 Despesa Total", f"R$ {despesa_total:,.2f}")
    col3.metric("💰 Saldo Final", f"R$ {saldo_final:,.2f}")
    col4.metric("📌 Margem (%)", f"{margem:.2f}%")

elif filtro_grafico == "📗 PIS e COFINS":
    st.subheader("📗 Apuração PIS e COFINS")

    # Ordenação correta dos meses
    ordem_meses = {"Janeiro": 1, "Fevereiro": 2, "Março": 3}
    meses_filtro = {
        "Janeiro/2025": ["Janeiro"],
        "Fevereiro/2025": ["Fevereiro"],
        "Março/2025": ["Março"],
        "1º Trimestre/2025": ["Janeiro", "Fevereiro", "Março"]
    }
    meses_selecionados = meses_filtro[filtro_periodo]
    piscofins_filtrado = piscofins_df[piscofins_df['Mês'].isin(meses_selecionados)]
    piscofins_filtrado = piscofins_filtrado.sort_values(by="Mês", key=lambda x: x.map(ordem_meses))

    # Tabela detalhada PRIMEIRO
    st.subheader("📋 Tabela Detalhada PIS e COFINS")
    st.dataframe(piscofins_filtrado[['Mês', 'Crédito', 'Débito', 'Saldo']],
                 use_container_width=True)

    # Gráfico de barras Créditos vs Débitos
    fig_pis = px.bar(piscofins_filtrado, x='Mês', y=['Crédito', 'Débito'], barmode='group',
                     title='Créditos vs Débitos PIS e COFINS')
    st.plotly_chart(fig_pis, use_container_width=True)

    # Gráfico de linha do Saldo acumulado (mensal)
    piscofins_filtrado['Saldo Acumulado'] = piscofins_filtrado['Saldo'].cumsum()
    fig_saldo_pis = px.line(
        piscofins_filtrado, x='Mês', y='Saldo Acumulado',
        title='Evolução Mensal do Saldo Acumulado - PIS e COFINS',
        markers=True
    )
    st.plotly_chart(fig_saldo_pis, use_container_width=True)

    # Cards de resumo financeiro para PIS/COFINS
    credito_total = piscofins_filtrado['Crédito'].sum()
    debito_total = piscofins_filtrado['Débito'].sum()
    saldo_final = credito_total - debito_total

    col1, col2, col3 = st.columns(3)
    col1.metric("💳 Total Créditos", f"R$ {credito_total:,.2f}")
    col2.metric("📌 Total Débitos", f"R$ {debito_total:,.2f}")
    col3.metric("💰 Saldo Final", f"R$ {saldo_final:,.2f}")

elif filtro_grafico == "📘 DRE Trimestral":
    st.subheader("📘 DRE Trimestral")
    dre_df['Valor'] = pd.to_numeric(dre_df['Valor'], errors='coerce').fillna(0)
    dre_total = dre_df.groupby('Descrição')['Valor'].sum().reset_index()

    # Tabela formatada
    st.dataframe(dre_total, use_container_width=True)

    # Gráfico de barras: Receita vs Resultado Líquido
    grupo = dre_total[dre_total['Descrição'].str.contains("Receita|Resultado", case=False)]
    fig_dre = px.bar(grupo, x='Descrição', y='Valor', title="Receita vs Resultado Líquido")
    st.plotly_chart(fig_dre, use_container_width=True)

    # Gráfico de pizza: despesas sobre o resultado
    despesas = dre_total[dre_total['Descrição'].str.contains("Despesa", case=False)]
    if not despesas.empty:
        fig_pizza_desp = px.pie(despesas, names='Descrição', values='Valor', title="Composição das Despesas", hole=0.3)
        fig_pizza_desp.update_traces(textinfo='label+percent')
        st.plotly_chart(fig_pizza_desp, use_container_width=True)

    # Destaque visual de prejuízo
    resultado = dre_total[dre_total['Descrição'].str.contains("Resultado Líquido", case=False)]['Valor'].sum()
    if resultado < 0:
        st.error(f"❌ Prejuízo apurado no período: R$ {abs(resultado):,.2f}")
    else:
        st.success(f"✅ Lucro apurado no período: R$ {resultado:,.2f}")

# Função para gerar Excel
def to_excel():
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        entradas_filtradas.to_excel(writer, sheet_name="Entradas", index=False)
        saidas_filtradas.to_excel(writer, sheet_name="Saídas", index=False)
        comparativo_filtrado.to_excel(writer, sheet_name="Apuracao", index=False)

        caixa_df.to_excel(writer, sheet_name="Caixa", index=False)
        piscofins_df.to_excel(writer, sheet_name="PISCOFINS", index=False)
        dre_df.to_excel(writer, sheet_name="DRE", index=False)
    processed_data = output.getvalue()
    return processed_data

# Botão para baixar o Excel completo
excel_bytes = to_excel()
st.download_button("⬇️ Baixar Relatórios Completos (.xlsx)",
                   data=excel_bytes,
                   file_name="Relatorio_ICMS_Completo.xlsx",
                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
