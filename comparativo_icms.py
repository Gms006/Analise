import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.colors as pc
from io import BytesIO

# ========== CONFIGURAÇÕES ==========
st.set_page_config(layout="wide")
st.title("📊 Relatório GH Sistemas")
caminho_planilha = "notas_processadas1.xlsx"

# ========== LEITURA ==========
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

# ========== LEITURA DA PLANILHA CONTABILIDADE ==========
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

# ========== FUNÇÕES AUXILIARES ==========

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
            # TRIMESTRE: Apenas saldo final do mês
            saldo_fim = df_mes['Valor Líquido'].cumsum().iloc[-1] + saldo_ant
            data_fim = df_mes['Data'].iloc[-1]
            pontos.append({'Data': data_fim, 'Saldo Acumulado': saldo_fim, 'Mês': mes})
        else:
            # MENSAL: Saldo anterior, saldo do dia 15, saldo final
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
    fig = px.line(df_pontos, x="Data", y="Saldo Acumulado", markers=True, title="Evolução Decacional do Saldo Acumulado - Caixa")
    st.plotly_chart(fig, use_container_width=True)

# ========== FILTROS DINÂMICOS ==========
st.sidebar.header("🎛️ Filtros")
periodos = {
    "Janeiro/2025": [1],
    "Fevereiro/2025": [2],
    "Março/2025": [3],
    "1º Trimestre/2025": [1, 2, 3]
}
filtro_periodo = st.sidebar.selectbox("Selecione o período:", list(periodos.keys()))

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
        "📘 DRE Trimestral",
        "📑 Tabelas Contabilidade"  # Nova opção
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

    st.subheader("📥 Entradas Filtradas")
    st.dataframe(entradas_filtradas.fillna("").astype(str), use_container_width=True)

    st.subheader("📤 Saídas Filtradas")
    st.dataframe(saidas_filtradas.fillna("").astype(str), use_container_width=True)

    st.write("### 📊 Comparativo de Crédito x Débito com Crédito Acumulado")
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
    st.download_button("⬇️ Baixar Relatórios Completos (.xlsx)",
                       data=excel_bytes,
                       file_name="Relatorio_ICMS_Completo.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

elif filtro_grafico == "📘 Contabilidade e Caixa":
    st.subheader("📘 Contabilidade e Caixa")

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
    col1.metric("📈 Total de Entradas", f"R$ {receita_total:,.2f}")
    col2.metric("📉 Total de Saídas", f"R$ {despesa_total:,.2f}")
    col3.metric("💰 Saldo Final", f"R$ {saldo_final:,.2f}")
    col4.metric("📌 Margem (%)", f"{margem:.2f}%")

    plotar_saldo_mensal(caixa_df, meses_selecionados)

elif filtro_grafico == "📗 PIS e COFINS":
    st.subheader("📗 Apuração PIS e COFINS")

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

    # Relatório fixo no início
    credito_total = piscofins_filtrado['Crédito'].sum()
    debito_total = piscofins_filtrado['Débito'].sum()
    saldo_final = credito_total - debito_total

    col1, col2, col3 = st.columns(3)
    col1.metric("💳 Total Créditos", f"R$ {credito_total:,.2f}")
    col2.metric("📌 Total Débitos", f"R$ {debito_total:,.2f}")
    col3.metric("💰 Saldo Final", f"R$ {saldo_final:,.2f}")

    # MANTENHA AQUI OS GRÁFICOS DE BARRA E PIZZA ORIGINAIS, por exemplo:
    if 'Crédito' in piscofins_filtrado.columns and 'Débito' in piscofins_filtrado.columns:
        fig_bar = px.bar(
            piscofins_filtrado, x='Mês', y=['Crédito', 'Débito'],
            barmode='group', title='Crédito x Débito por Mês'
        )
        st.plotly_chart(fig_bar, use_container_width=True)
# Gráfico de linha: apenas saldo final de cada mês (igual ao trimestre do caixa)
    pontos = []
    for mes_nome in meses_selecionados:
        df_mes = piscofins_filtrado[piscofins_filtrado['Mês'] == mes_nome]
        if df_mes.empty:
            continue
        saldo_fim = df_mes['Saldo'].iloc[-1]
        pontos.append({'Mês': mes_nome, 'Saldo Acumulado': saldo_fim})

    df_pontos = pd.DataFrame(pontos)
    fig_saldo_pis = px.line(
        df_pontos, x='Mês', y='Saldo Acumulado',
        title='Evolução Mensal do Saldo Acumulado - PIS e COFINS',
        markers=True
    )
    st.plotly_chart(fig_saldo_pis, use_container_width=True)

elif filtro_grafico == "📘 DRE Trimestral":
    st.subheader("📘 DRE Trimestral")
    dre_df['Valor'] = pd.to_numeric(dre_df['Valor'], errors='coerce').fillna(0)
    dre_total = dre_df.groupby('Descrição')['Valor'].sum().reset_index()

    # Tabela dinâmica e fácil de visualizar
    st.markdown("### 📋 Tabela Completa DRE")
    st.dataframe(dre_df, use_container_width=True)

elif filtro_grafico == "📑 Tabelas Contabilidade":
    st.subheader("📑 Todas as Tabelas de Contabilidade")

    st.markdown("### Caixa")
    st.dataframe(caixa_df, use_container_width=True)

    st.markdown("### PIS e COFINS")
    st.dataframe(piscofins_df, use_container_width=True)

    st.markdown("### DRE 1º Trimestre")
    st.dataframe(dre_df, use_container_width=True)

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

excel_bytes = to_excel()
st.download_button("⬇️ Baixar Relatórios Completos (.xlsx)",
                   data=excel_bytes,
                   file_name="Relatorio_ICMS_Completo.xlsx",
                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
