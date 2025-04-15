import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.colors as pc
from io import BytesIO

# ========== CONFIGURAÇÕES ==========
st.set_page_config(layout="wide")
st.title("📊 Relatório Interativo de ICMS")
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
caixa_df = planilha_contabil['Caixa']
pis_df = planilha_contabil['PIS']
cofins_df = planilha_contabil['COFINS']
dre_df = planilha_contabil['DRE 1º Trimestre']

# ========== FILTROS DINÂMICOS ==========
st.sidebar.header("🎛️ Filtros")
periodos = {
    "Janeiro/2025": [1],
    "Fevereiro/2025": [2],
    "Março/2025": [3],
    "1º Trimestre/2025": [1, 2, 3]
}
filtro_periodo = st.sidebar.selectbox("Selecione o período:", list(periodos.keys()))
filtro_grafico = st.sidebar.selectbox("Tipo de gráfico:", [
    "Mapa por UF",
    "Comparativo de Crédito x Débito",
    "Apuração com Crédito Acumulado",
    "Relatórios Detalhados",
    "📘 Contabilidade e Caixa",
    "📗 PIS",
    "📙 COFINS",
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

    # Conversão de data e extração do mês/ano
    caixa_df['Data'] = pd.to_datetime(caixa_df['Data'], errors='coerce')
    caixa_df = caixa_df.dropna(subset=['Data', 'Tipo', 'Valor'])
    caixa_df['Valor'] = pd.to_numeric(caixa_df['Valor'], errors='coerce').fillna(0)
    caixa_df['Mês'] = caixa_df['Data'].dt.strftime('%b/%Y')

    # ========== CARDS RESUMO ==========
    total_receitas = caixa_df[caixa_df['Tipo'].str.lower() == 'entrada']['Valor'].sum()
    total_despesas = caixa_df[caixa_df['Tipo'].str.lower() == 'saida']['Valor'].sum()
    saldo = total_receitas - total_despesas
    margem = (saldo / total_receitas) * 100 if total_receitas > 0 else 0

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Receitas Totais", f"R$ {total_receitas:,.2f}")
    col2.metric("Despesas Totais", f"R$ {total_despesas:,.2f}")
    col3.metric("Saldo Final", f"R$ {saldo:,.2f}")
    col4.metric("Margem (%)", f"{margem:.2f}%")

    # ========== GRÁFICO BARRAS RECEITA x DESPESAS ==========
    graf_df = caixa_df.groupby(['Mês', 'Tipo'])['Valor'].sum().reset_index()
    fig_bar = px.bar(graf_df, x='Mês', y='Valor', color='Tipo', barmode='group',
                    title='Comparativo de Entradas e Saídas por Mês')
    st.plotly_chart(fig_bar, use_container_width=True)

    # ========== SALDO ACUMULADO ==========
    caixa_df = caixa_df.sort_values(by='Data')
    caixa_df['Saldo Acumulado'] = caixa_df.apply(
        lambda row: row['Valor'] if row['Tipo'].lower() == 'entrada' else -row['Valor'], axis=1
    ).cumsum()
    fig_saldo = px.line(caixa_df, x='Data', y='Saldo Acumulado', title='Saldo Acumulado ao Longo do Tempo')
    st.plotly_chart(fig_saldo, use_container_width=True)

    # ========== GRÁFICO PIZZA CATEGORIAS ==========
    if 'Categoria' in caixa_df.columns:
        categoria_df = caixa_df[caixa_df['Tipo'].str.lower() == 'saida'].groupby('Categoria')['Valor'].sum().reset_index()
        fig_pizza = px.pie(categoria_df, names='Categoria', values='Valor', title='Distribuição das Despesas por Categoria', hole=0.3)
        fig_pizza.update_traces(textinfo='label+percent')
        st.plotly_chart(fig_pizza, use_container_width=True)

    # ========== TABELA DETALHADA ==========
    st.write("\n### 📜 Tabela Detalhada de Caixa")
    st.dataframe(caixa_df, use_container_width=True)


# ========== 2. PIS e COFINS ==========

# Objetivo: Apurar e exibir os gráficos de crédito e débito do PIS/COFINS e mostrar créditos a transportar
# Justificativa: As seções estavam com texto fixo sem processar os dados lidos

for tributo, df in zip(["PIS", "COFINS"], [pis_df, cofins_df]):
    if filtro_grafico == f"📗 {tributo}":
        st.subheader(f"📗 Apuração de {tributo}")

        df['Mês'] = pd.to_datetime(df['Mês'], errors='coerce')
        df['Mês Ref'] = df['Mês'].dt.strftime('%b/%Y')

        df['Crédito'] = pd.to_numeric(df['Crédito'], errors='coerce').fillna(0)
        df['Débito'] = pd.to_numeric(df['Débito'], errors='coerce').fillna(0)
        df['A Transportar'] = pd.to_numeric(df['A Transportar'], errors='coerce').fillna(0)

        # Cards
        total_cred = df['Crédito'].sum()
        total_deb = df['Débito'].sum()
        total_transp = df['A Transportar'].sum()

        c1, c2, c3 = st.columns(3)
        c1.metric("Total de Créditos", f"R$ {total_cred:,.2f}")
        c2.metric("Total de Débitos", f"R$ {total_deb:,.2f}")
        c3.metric("Créditos a Transportar", f"R$ {total_transp:,.2f}")

        # Gráfico
        graf = df[['Mês Ref', 'Crédito', 'Débito']].melt(id_vars='Mês Ref', var_name='Tipo', value_name='Valor')
        fig = px.bar(graf, x='Mês Ref', y='Valor', color='Tipo', barmode='group', title=f'{tributo} - Crédito x Débito')
        st.plotly_chart(fig, use_container_width=True)

        st.write("### 📊 Detalhamento dos Créditos a Transportar")
        st.dataframe(df[['Mês Ref', 'A Transportar']], use_container_width=True)


# ========== 3. DRE ==========

# Objetivo: Apresentar o resultado gerencial do período e destacar prejuízos
# Justificativa: A seção estava com texto placeholder, sem exibir os dados da aba "DRE 1º Trimestre"

if filtro_grafico == "📘 DRE Trimestral":
    st.subheader("📘 Demonstração do Resultado do Exercício")
    dre_df.columns = dre_df.columns.str.strip()

    if 'Valor' in dre_df.columns:
        dre_df['Valor'] = pd.to_numeric(dre_df['Valor'], errors='coerce').fillna(0)

        # DRE formatada
        st.dataframe(dre_df, use_container_width=True)

        # Destacar prejuízo
        resultado = dre_df[dre_df['Conta'].str.contains("Resultado Líquido", case=False)]['Valor'].sum()
        if resultado < 0:
            st.error(f"❌ Prejuízo apurado no período: R$ {abs(resultado):,.2f}")
        else:
            st.success(f"✅ Lucro apurado no período: R$ {resultado:,.2f}")

        # Gráfico barras Receita vs Resultado
        grupo = dre_df[dre_df['Conta'].str.contains("Receita|Resultado", case=False)]
        fig_dre = px.bar(grupo, x='Conta', y='Valor', title="Receita vs Resultado Líquido")
        st.plotly_chart(fig_dre, use_container_width=True)

        # Gráfico de pizza de despesas (se houver)
        despesas = dre_df[dre_df['Conta'].str.contains("Despesa", case=False)]
        if not despesas.empty:
            fig_pizza_desp = px.pie(despesas, names='Conta', values='Valor', title="Composição das Despesas", hole=0.3)
            fig_pizza_desp.update_traces(textinfo='label+percent')
            st.plotly_chart(fig_pizza_desp, use_container_width=True)

# Função para gerar Excel
def to_excel():
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        entradas_filtradas.to_excel(writer, sheet_name="Entradas", index=False)
        saidas_filtradas.to_excel(writer, sheet_name="Saídas", index=False)
        comparativo_filtrado.to_excel(writer, sheet_name="Apuracao", index=False)

        caixa_df.to_excel(writer, sheet_name="Caixa", index=False)
        pis_df.to_excel(writer, sheet_name="PIS", index=False)
        cofins_df.to_excel(writer, sheet_name="COFINS", index=False)
        dre_df.to_excel(writer, sheet_name="DRE", index=False)
    processed_data = output.getvalue()
    return processed_data

# Botão para baixar o Excel completo
excel_bytes = to_excel()
st.download_button("⬇️ Baixar Relatórios Completos (.xlsx)",
                   data=excel_bytes,
                   file_name="Relatorio_ICMS_Completo.xlsx",
                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
