import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO

st.set_page_config(page_title="Dashboard Sky Group", layout="wide")

# --- CSS personalizado ---
st.markdown("""
    <style>
    .main, .block-container {
        background-color: #f2f2f2;
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    }
    .metric-container {
        background-color: white;
        border-radius: 15px;
        padding: 20px;
        text-align: center;
        box-shadow: 0px 4px 12px rgba(0, 0, 0, 0.1);
        margin-bottom: 10px;
        border: 1px solid #e0e0e0;
        transition: transform 0.2s ease;
    }
    .metric-container:hover {
        transform: translateY(-2px);
        box-shadow: 0px 6px 16px rgba(0, 0, 0, 0.15);
    }
    .metric-value {
        font-size: 28px;
        font-weight: bold;
        color: #333333;
        margin-bottom: 5px;
    }
    .metric-label {
        font-size: 14px;
        color: #666666;
        font-weight: 500;
    }
    .header-title {
        color: #ef5b25;
        text-align: center;
        margin-bottom: 30px;
        font-size: 2.5em;
        font-weight: 700;
    }
    .section-header {
        color: #333333;
        margin-top: 30px;
        margin-bottom: 20px;
        font-size: 1.5em;
        font-weight: 600;
        border-left: 5px solid #ef5b25;
        padding-left: 15px;
    }
    .footer {
        text-align: center;
        padding: 20px;
        background: #e0e0e0;
        margin-top: 40px;
        border-radius: 10px;
        color: #666666;
        font-weight: 500;
    }
    .legenda-status {
        background: white;
        padding: 15px;
        border-radius: 10px;
        margin: 20px 0;
        box-shadow: 0px 2px 8px rgba(0, 0, 0, 0.1);
    }
    </style>
""", unsafe_allow_html=True)

# --- Carregamento de dados com cache ---
@st.cache_data
def carregar_dados():
    try:
        df = pd.read_excel("Limite geral.xlsx", sheet_name="Limite geral")
        
        # Limpeza e normalização dos nomes das colunas
        df.columns = (df.columns.str.strip()
                     .str.normalize('NFKD')
                     .str.encode('ascii', errors='ignore')
                     .str.decode('utf-8'))
        
        # Renomear colunas para nomes consistentes
        df.rename(columns={
            'Fi': 'Filial',
            'CNPJ/CPF': 'CNPJ_CPF',
            'Limite': 'Limite',
            'Risk Scor': 'Risk_Score',
            'Status': 'Status',
            'Media,Ser': 'Media_Ser',
            'Media,Pa': 'Media_Pa',
            'Lim,Sug,': 'Lim_Sug',
            'Lim,Real': 'Lim_Real',
            'Vencido': 'Vencido',
            'A vencer': 'A_Vencer',
            'Divida': 'Divida',
            'Disponivel': 'Disponivel',
            'Vendedor': 'Vendedor',
            'Receita': 'Receita',
            'Fundacao': 'Fundacao',
            'T.M': 'TM'
        }, inplace=True)
        
        # Limpeza e tratamento de dados
        df['Filial'] = df['Filial'].astype(str).str.strip()
        
        # Tratar coluna Vendedor - converter para string
        df['Vendedor'] = df['Vendedor'].astype(str).str.strip()
        
        # Preencher valores nulos nas colunas numéricas
        colunas_numericas = ['Vencido', 'A_Vencer', 'Disponivel', 'Limite', 'Risk_Score']
        for coluna in colunas_numericas:
            if coluna in df.columns:
                df[coluna] = pd.to_numeric(df[coluna], errors='coerce').fillna(0)
        
        return df
        
    except FileNotFoundError:
        st.error("❌ Arquivo 'Limite geral.xlsx' não encontrado. Por favor, verifique se o arquivo está na pasta correta.")
        st.stop()
    except Exception as e:
        st.error(f"❌ Erro ao carregar o arquivo: {e}")
        st.stop()

# --- Título e Logo ---
col_logo, col_titulo = st.columns([1, 4])
with col_titulo:
    st.markdown("<div class='header-title'>📊 Dashboard Financeiro - Sky Group</div>", unsafe_allow_html=True)

# --- Botão para recarregar dados ---
if st.button("🔄 Atualizar Dados"):
    st.cache_data.clear()
    st.rerun()

# --- Carregar dados ---
df = carregar_dados()

# --- Filtros ---
st.markdown("<div class='section-header'>🔍 Filtros</div>", unsafe_allow_html=True)

col1, col2 = st.columns(2)
with col1:
    filiais = sorted(df['Filial'].unique().tolist())
    filiais_select = st.multiselect(
        "**Filial**", 
        options=filiais, 
        default=filiais,
        help="Selecione uma ou mais filiais"
    )

with col2:
    vendedores = sorted(df['Vendedor'].unique().tolist())
    vendedor_select = st.selectbox(
        "**Vendedor**", 
        options=['Todos'] + vendedores,
        help="Selecione um vendedor específico ou 'Todos'"
    )

# --- Aplicar filtros ---
df_filtrado = df.copy()
if filiais_select:
    df_filtrado = df_filtrado[df_filtrado['Filial'].isin(filiais_select)]
if vendedor_select != 'Todos':
    df_filtrado = df_filtrado[df_filtrado['Vendedor'] == vendedor_select]

# --- Métricas em Cards ---
st.markdown("<div class='section-header'>📈 Métricas Principais</div>", unsafe_allow_html=True)

vencido_total = df_filtrado['Vencido'].sum()
avencer_total = df_filtrado['A_Vencer'].sum()
disponivel_total = df_filtrado['Disponivel'].sum()
total_carteira = vencido_total + avencer_total
inad_pct = (vencido_total / total_carteira * 100) if total_carteira else 0

# Função para formatar valores monetários
def formatar_moeda(valor):
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

col1, col2, col3, col4 = st.columns(4)

with col1:
    st.markdown(f"""
        <div class='metric-container'>
            <div class='metric-value' style='color: #ef5b25;'>{formatar_moeda(vencido_total)}</div>
            <div class='metric-label'>💰 Total Vencido</div>
        </div>
    """, unsafe_allow_html=True)

with col2:
    st.markdown(f"""
        <div class='metric-container'>
            <div class='metric-value' style='color: #ff9800;'>{formatar_moeda(avencer_total)}</div>
            <div class='metric-label'>📆 Total a Vencer</div>
        </div>
    """, unsafe_allow_html=True)

with col3:
    st.markdown(f"""
        <div class='metric-container'>
            <div class='metric-value' style='color: #4caf50;'>{formatar_moeda(disponivel_total)}</div>
            <div class='metric-label'>✅ Crédito Disponível</div>
        </div>
    """, unsafe_allow_html=True)

with col4:
    cor_inad = "#ef5b25" if inad_pct > 10 else "#ff9800" if inad_pct > 5 else "#4caf50"
    st.markdown(f"""
        <div class='metric-container'>
            <div class='metric-value' style='color: {cor_inad};'>{inad_pct:.1f}%</div>
            <div class='metric-label'>⚠️ Inadimplência</div>
        </div>
    """, unsafe_allow_html=True)

# --- Legenda dos Status ---
st.markdown("<div class='section-header'>📋 Legenda de Classificação</div>", unsafe_allow_html=True)

st.markdown("""
<div class='legenda-status'>
    <h4>🎯 Classificação por Risk Score:</h4>
    <div style='display: grid; grid-template-columns: repeat(2, 1fr); gap: 10px; margin-top: 10px;'>
        <div style='display: flex; align-items: center;'>
            <div style='width: 20px; height: 20px; background-color: #ef5b25; border-radius: 50%; margin-right: 10px;'></div>
            <span><strong>🟥 Ruim</strong> → Score até 201</span>
        </div>
        <div style='display: flex; align-items: center;'>
            <div style='width: 20px; height: 20px; background-color: #ff9800; border-radius: 50%; margin-right: 10px;'></div>
            <span><strong>🟧 Regular</strong> → Score até 500</span>
        </div>
        <div style='display: flex; align-items: center;'>
            <div style='width: 20px; height: 20px; background-color: #ffc107; border-radius: 50%; margin-right: 10px;'></div>
            <span><strong>🟨 Bom</strong> → Score até 700</span>
        </div>
        <div style='display: flex; align-items: center;'>
            <div style='width: 20px; height: 20px; background-color: #4caf50; border-radius: 50%; margin-right: 10px;'></div>
            <span><strong>🟩 Ótimo</strong> → Score acima de 700</span>
        </div>
    </div>
</div>
""", unsafe_allow_html=True)

# --- Gráfico comparativo por Status ---
st.markdown("<div class='section-header'>📊 Gráfico Comparativo por Status</div>", unsafe_allow_html=True)

ordem_status = ['Ruim', 'Regular', 'Bom', 'Ótimo']
df_grafico = (df_filtrado.groupby('Status')[['Vencido', 'A_Vencer', 'Disponivel']]
              .sum()
              .reindex(ordem_status)
              .fillna(0)
              .reset_index())

df_long = pd.melt(df_grafico, id_vars='Status', var_name='Tipo', value_name='Valor')

fig = px.bar(df_long, 
             x='Status', 
             y='Valor', 
             color='Tipo', 
             barmode='group',
             color_discrete_map={
                 "Vencido": "#ef5b25", 
                 "A_Vencer": "#ff9800", 
                 "Disponivel": "#4caf50"
             },
             height=500,
             title="Comparação de Valores por Status do Cliente")

fig.update_layout(
    plot_bgcolor='white',
    paper_bgcolor='white',
    font=dict(size=12),
    xaxis_title="Status do Cliente",
    yaxis_title="Valor (R$)",
    legend_title="Tipo de Valor"
)

# Formatar eixos Y em R$
fig.update_yaxes(tickprefix="R$ ", tickformat=",.2f")

st.plotly_chart(fig, use_container_width=True)

# --- Tabelas Detalhadas ---
st.markdown("<div class='section-header'>📂 Tabelas Detalhadas</div>", unsafe_allow_html=True)

def gerar_tabela_formatada(coluna, nome_exibicao):
    serie = df_filtrado.groupby('Status')[coluna].sum().reindex(ordem_status).fillna(0)
    total = serie.sum()
    percentual = (serie / total * 100).round(2) if total else 0
    
    df_tabela = pd.DataFrame({
        nome_exibicao: serie,
        'Percentual (%)': percentual
    })
    
    return df_tabela

col1, col2 = st.columns(2)

with col1:
    st.markdown("**💰 Vencido por Status**")
    tabela_vencido = gerar_tabela_formatada('Vencido', 'Vencido')
    st.dataframe(
        tabela_vencido.style.format({
            "Vencido": "R$ {:,.2f}",
            "Percentual (%)": "{:,.2f}%"
        }),
        use_container_width=True
    )

with col2:
    st.markdown("**📆 A Vencer por Status**")
    tabela_avencer = gerar_tabela_formatada('A_Vencer', 'A Vencer')
    st.dataframe(
        tabela_avencer.style.format({
            "A Vencer": "R$ {:,.2f}",
            "Percentual (%)": "{:,.2f}%"
        }),
        use_container_width=True
    )

col3, col4 = st.columns(2)

with col3:
    st.markdown("**✅ Disponível por Status**")
    tabela_disponivel = gerar_tabela_formatada('Disponivel', 'Disponível')
    st.dataframe(
        tabela_disponivel.style.format({
            "Disponível": "R$ {:,.2f}",
            "Percentual (%)": "{:,.2f}%"
        }),
        use_container_width=True
    )

with col4:
    st.markdown("**👥 Distribuição de Clientes por Status**")
    contagem = df_filtrado['Status'].value_counts().reindex(ordem_status, fill_value=0)
    total = contagem.sum()
    percentual = (contagem / total * 100).round(2) if total else 0
    df_clientes = pd.DataFrame({
        "Quantidade": contagem,
        "Percentual (%)": percentual
    })
    st.dataframe(
        df_clientes.style.format({
            "Quantidade": "{:,.0f}",
            "Percentual (%)": "{:,.2f}%"
        }),
        use_container_width=True
    )

# --- Análise por Tempo de Mercado ---
st.markdown("<div class='section-header'>⏳ Análise por Tempo de Mercado</div>", unsafe_allow_html=True)

if 'TM' in df_filtrado.columns:
    df_tm = df_filtrado.copy()

    # Extrair apenas números e converter para float
    df_tm['TM'] = df_tm['TM'].astype(str).str.extract(r'(\d+(?:,\d+)?)')[0].str.replace(',', '.').astype(float)

    # Criar faixas de tempo de mercado
    def classificar_tempo_mercado(valor):
        if pd.isna(valor):
            return "Não informado"
        elif valor <= 2:
            return "0–2 anos"
        elif valor <= 5:
            return "3–5 anos"
        elif valor <= 10:
            return "6–10 anos"
        elif valor <= 20:
            return "11–20 anos"
        else:
            return "Acima de 20 anos"

    df_tm['Faixa_Tempo'] = df_tm['TM'].apply(classificar_tempo_mercado)

    # Criar flags para clientes com valores vencidos e a vencer
    df_tm['Cliente_Vencido'] = df_tm['Vencido'] > 0
    df_tm['Cliente_Avencer'] = df_tm['A_Vencer'] > 0

    # Agrupar e somar
    resumo_tm = (df_tm.groupby('Faixa_Tempo')
                 .agg({
                     'Vencido': 'sum',
                     'A_Vencer': 'sum',
                     'CNPJ_CPF': 'count',
                     'Cliente_Vencido': 'sum',
                     'Cliente_Avencer': 'sum'
                 })
                 .rename(columns={
                     'CNPJ_CPF': 'Clientes_Totais',
                     'Cliente_Vencido': 'Clientes_Vencido',
                     'Cliente_Avencer': 'Clientes_Avencer'
                 })
                 .reindex(["0–2 anos", "3–5 anos", "6–10 anos", "11–20 anos", "Acima de 20 anos", "Não informado"])
                 .fillna(0)
                 .reset_index())

    # Calcular percentuais DENTRO DE CADA FAIXA
    resumo_tm['Total_Faixa'] = resumo_tm['Vencido'] + resumo_tm['A_Vencer']
    resumo_tm['% Vencido'] = (resumo_tm['Vencido'] / resumo_tm['Total_Faixa'] * 100).round(2).fillna(0)
    resumo_tm['% A_Vencer'] = (resumo_tm['A_Vencer'] / resumo_tm['Total_Faixa'] * 100).round(2).fillna(0)

    # Exibir os resultados faixa a faixa
    for _, row in resumo_tm.iterrows():
        faixa = row['Faixa_Tempo']
        vencido = formatar_moeda(row['Vencido'])
        avencer = formatar_moeda(row['A_Vencer'])
        clientes_total = int(row['Clientes_Totais'])
        clientes_venc = int(row['Clientes_Vencido'])
        clientes_aven = int(row['Clientes_Avencer'])
        pct_venc = row['% Vencido']
        pct_aven = row['% A_Vencer']

        st.markdown(f"""
        <div style='background:white; padding:15px; border-radius:10px; margin-bottom:10px;
                    box-shadow:0px 2px 8px rgba(0,0,0,0.1);'>
            <strong>📅 {faixa}</strong><br>
            💰 <strong>Vencido:</strong> {vencido} ({pct_venc:.2f}%) — <strong>{clientes_venc}</strong> clientes<br>
            📆 <strong>A Vencer:</strong> {avencer} ({pct_aven:.2f}%) — <strong>{clientes_aven}</strong> clientes<br>
            👥 <strong>Total na Faixa:</strong> {clientes_total} clientes
        </div>
        """, unsafe_allow_html=True)

    # --- Gráfico de Linha ---
    st.markdown("### 📈 Gráfico de Vencido e A Vencer por Faixa de Tempo de Mercado")

    import plotly.graph_objects as go

    fig_tm = go.Figure()

    fig_tm.add_trace(go.Scatter(
        x=resumo_tm['Faixa_Tempo'],
        y=resumo_tm['Vencido'],
        mode='lines+markers',
        name='Vencido',
        line=dict(width=3, color='#ef5b25')
    ))

    fig_tm.add_trace(go.Scatter(
        x=resumo_tm['Faixa_Tempo'],
        y=resumo_tm['A_Vencer'],
        mode='lines+markers',
        name='A Vencer',
        line=dict(width=3, color='#ff9800')
    ))

    fig_tm.update_layout(
        title="Evolução de Valores por Tempo de Mercado",
        xaxis_title="Faixa de Tempo de Mercado",
        yaxis_title="Valor (R$)",
        yaxis_tickprefix="R$ ",
        plot_bgcolor='white',
        paper_bgcolor='white',
        font=dict(size=12),
        hovermode="x unified"
    )

    st.plotly_chart(fig_tm, use_container_width=True)

else:
    st.warning("⚠️ Coluna 'T.M' não encontrada no arquivo.")

# --- Rankings ---
st.markdown("<div class='section-header'>🏆 Rankings</div>", unsafe_allow_html=True)

col_rank1, col_rank2 = st.columns(2)

with col_rank1:
    st.markdown("**🏆 Top 10 Clientes com Maior Valor Vencido**")
    if 'CNPJ_CPF' in df_filtrado.columns:
        ranking_clientes = (df_filtrado.groupby('CNPJ_CPF')
                           .agg({'Vencido': 'sum', 'Status': 'first'})
                           .sort_values('Vencido', ascending=False)
                           .head(10)
                           .reset_index())
        
        ranking_clientes['Posição'] = range(1, len(ranking_clientes) + 1)
        ranking_clientes = ranking_clientes[['Posição', 'CNPJ_CPF', 'Status', 'Vencido']]
        
        st.dataframe(
            ranking_clientes.style.format({
                "Vencido": "R$ {:,.2f}"
            }).hide(axis='index'),
            use_container_width=True
        )

with col_rank2:
    st.markdown("**🥇 Ranking de Filiais por Inadimplência**")
    
    df_ranking_filial = df_filtrado.groupby('Filial').agg(
        Vencido=('Vencido', 'sum'),
        A_Vencer=('A_Vencer', 'sum')
    ).reset_index()
    
    df_ranking_filial['Total Carteira'] = df_ranking_filial['Vencido'] + df_ranking_filial['A_Vencer']
    df_ranking_filial['Inadimplência (%)'] = (
        df_ranking_filial['Vencido'] / df_ranking_filial['Total Carteira'] * 100
    ).fillna(0).replace([float('inf'), -float('inf')], 0).round(2)
    
    df_ranking_filial = df_ranking_filial.sort_values('Inadimplência (%)', ascending=False)
    df_ranking_filial['Posição'] = range(1, len(df_ranking_filial) + 1)
    df_ranking_filial = df_ranking_filial[['Posição', 'Filial', 'Inadimplência (%)', 'Vencido', 'A_Vencer', 'Total Carteira']]
    
    st.dataframe(
        df_ranking_filial.style.format({
            "Vencido": "R$ {:,.2f}",
            "A_Vencer": "R$ {:,.2f}",
            "Total Carteira": "R$ {:,.2f}",
            "Inadimplência (%)": "{:,.2f}%"
        }).hide(axis='index'),
        use_container_width=True
    )

# --- Exportação para Excel ---
st.markdown("<div class='section-header'>📥 Exportação de Dados</div>", unsafe_allow_html=True)

st.markdown("""
<div style='background: white; padding: 20px; border-radius: 10px; box-shadow: 0px 2px 8px rgba(0, 0, 0, 0.1);'>
    <p>Exporte os dados filtrados para análise em Excel:</p>
</div>
""", unsafe_allow_html=True)

output = BytesIO()
with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
    df_filtrado.to_excel(writer, index=False, sheet_name='Dados_Filtrados')
    
    # Adicionar aba com resumo
    resumo = pd.DataFrame({
        'Métrica': ['Total Vencido', 'Total a Vencer', 'Crédito Disponível', 'Inadimplência (%)'],
        'Valor': [vencido_total, avencer_total, disponivel_total, inad_pct]
    })
    resumo.to_excel(writer, index=False, sheet_name='Resumo')

data = output.getvalue()

col_export1, col_export2, col_export3 = st.columns([1, 2, 1])
with col_export2:
    st.download_button(
        label="📤 Baixar Dados Filtrados em Excel",
        data=data,
        file_name=f"dashboard_financeiro_skygroup_{pd.Timestamp.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

# --- Rodapé ---
st.markdown("""
<div class='footer'>
    Desenvolvido por Daniel Mendes | Sky Group Brasil
</div>
""", unsafe_allow_html=True)