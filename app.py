import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
from simple_salesforce import Salesforce
import io

# ==========================================
# 0. CONFIGURAÇÃO E DESIGN EXECUTIVO
# ==========================================
st.set_page_config(
    page_title="Gestão de Baixas | Financeiro", 
    page_icon="⚖️", 
    layout="wide"
)

st.markdown("""
    <style>
    [data-testid="stAppViewBlockContainer"] {
        padding-top: 1.5rem;
        padding-bottom: 1rem;
    }
    header {visibility: hidden;}
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}

    .main { background-color: #F0F2F6; }
    
    .metric-card {
        background-color: #ffffff;
        border-radius: 12px;
        padding: 20px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.05);
        border-left: 5px solid #1E3A8A;
        margin-bottom: 10px;
    }
    .metric-label {
        color: #64748B;
        font-size: 0.9rem;
        font-weight: 600;
        text-transform: uppercase;
        margin-bottom: 5px;
    }
    .metric-value {
        color: #1E293B;
        font-size: 1.8rem;
        font-weight: 700;
    }
    
    .main-title {
        color: #1E3A8A;
        font-size: 24px;
        font-weight: 800;
        margin-bottom: 25px;
    }
    </style>
    """, unsafe_allow_html=True)

def kpi_card(label, value):
    st.markdown(f"""
        <div class="metric-card">
            <div class="metric-label">{label}</div>
            <div class="metric-value">{value}</div>
        </div>
    """, unsafe_allow_html=True)

def formatar_moeda_br(valor):
    """Formata valores para o padrão R$ 1.234,56"""
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

st.markdown('<p class="main-title">Monitoramento: Baixas Manuais de Ativos</p>', unsafe_allow_html=True)

# ==========================================
# 1. MOTOR DE DADOS (SEM CACHE PARA DADOS REAIS)
# ==========================================
# Removido @st.cache_data para garantir conexão a cada acesso
def carregar_dados_full():
    usuario = st.secrets["sf_user"]
    senha = st.secrets["sf_pass"]
    token = st.secrets["sf_token"]
    
    sf = Salesforce(username=usuario, password=senha, security_token=token)
    
    query = """
    SELECT 
        FOZ_CodigoItem__c,
        SerialNumber,
        Status,
        FOZ_Data_Ativacao_Inativacao_Manual__c,
        FOZ_Motivo_Inativacao_Manual__c,
        FOZ_ValorTotal__c,
        FOZ_ContaRecebedora__r.Name
    FROM Asset
    WHERE FOZ_Motivo_Desinstalacao__c = 'BAIXA MANUAL DE ATIVO'
    """
    
    res = sf.query_all(query)
    df = pd.json_normalize(res['records'])
    
    if df.empty: 
        # Ajuste de fuso horário (-3h para Brasília)
        return pd.DataFrame(), datetime.now() - timedelta(hours=3)

    df.drop(columns=[c for c in df.columns if 'attributes' in c], inplace=True, errors='ignore')
    
    df['FOZ_CodigoItem__c'] = df['FOZ_CodigoItem__c'].apply(
        lambda x: str(x).split('.')[0].zfill(8) if pd.notna(x) else ""
    )
    
    df['FOZ_Data_Ativacao_Inativacao_Manual__c'] = pd.to_datetime(df['FOZ_Data_Ativacao_Inativacao_Manual__c']).dt.tz_localize(None)
    df['FOZ_ValorTotal__c'] = pd.to_numeric(df['FOZ_ValorTotal__c'], errors='coerce').fillna(0.0)
    
    df['Ano'] = df['FOZ_Data_Ativacao_Inativacao_Manual__c'].dt.year
    df['Mes_Num'] = df['FOZ_Data_Ativacao_Inativacao_Manual__c'].dt.month
    
    df = df.rename(columns={
        'FOZ_ContaRecebedora__r.Name': 'Cliente',
        'FOZ_CodigoItem__c': 'Item do Contrato',
        'SerialNumber': 'Nº de Série',
        'FOZ_Data_Ativacao_Inativacao_Manual__c': 'Data Baixa',
        'FOZ_Motivo_Inativacao_Manual__c': 'Motivo',
        'FOZ_ValorTotal__c': 'Valor (R$)'
    })
    
    # Ajuste de fuso horário na hora da sincronização
    return df, datetime.now() - timedelta(hours=3)

# ==========================================
# 2. CONTROLES E FILTROS
# ==========================================
st.sidebar.markdown("### 🏛️ Governança Financeira")
df_base, data_ref = carregar_dados_full()

if not df_base.empty:
    anos = sorted(df_base['Ano'].unique().tolist(), reverse=True)
    ano_sel = st.sidebar.selectbox("Exercício (Ano)", anos)

    meses_map = {1:"Jan", 2:"Fev", 3:"Mar", 4:"Abr", 5:"Mai", 6:"Jun", 
                 7:"Jul", 8:"Ago", 9:"Set", 10:"Out", 11:"Nov", 12:"Dez"}
    
    meses_disponiveis = sorted(df_base[df_base['Ano'] == ano_sel]['Mes_Num'].unique().tolist())
    mes_sel = st.sidebar.selectbox("Mês de Referência", meses_disponiveis, format_func=lambda x: meses_map[x])

    df_filtrado = df_base[(df_base['Ano'] == ano_sel) & (df_base['Mes_Num'] == mes_sel)]

    st.sidebar.divider()
    if st.sidebar.button("🔄 Forçar Atualização"):
        st.rerun()

    st.sidebar.caption(f"Última Sincronização: \n{data_ref.strftime('%d/%m/%Y %H:%M:%S')}")

# ==========================================
# 3. DASHBOARD E KPIs
# ==========================================
if not df_base.empty:
    c1, c2, c3 = st.columns(3)
    
    total_vol = len(df_filtrado)
    total_fin = df_filtrado['Valor (R$)'].sum()
    media = total_fin / total_vol if total_vol > 0 else 0

    with c1:
        kpi_card("Volume de Baixas", f"{total_vol} unidades")
    with c2:
        kpi_card("Total Financeiro", formatar_moeda_br(total_fin))
    with c3:
        kpi_card("Ticket Médio", formatar_moeda_br(media))
    
    st.markdown("<br>", unsafe_allow_html=True)

    # Tabela
    cols_exibicao = ['Cliente', 'Item do Contrato', 'Nº de Série', 'Status', 'Data Baixa', 'Motivo', 'Valor (R$)']
    df_show = df_filtrado[cols_exibicao].copy()
    
    # Formatação para exibição na tabela
    df_tabela = df_show.copy()
    df_tabela['Data Baixa'] = df_tabela['Data Baixa'].dt.strftime('%d/%m/%Y')
    df_tabela['Valor (R$)'] = df_tabela['Valor (R$)'].apply(formatar_moeda_br)

    st.dataframe(df_tabela, use_container_width=True, hide_index=True)

    # Exportação
    if not df_filtrado.empty:
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            # Exporta os dados com valores numéricos para o Excel permitir cálculos
            df_show.to_excel(writer, index=False, sheet_name='Extrato')
        
        st.download_button(
            label="📄 Exportar Extrato Excel",
            data=buffer.getvalue(),
            file_name=f"Baixas_{meses_map[mes_sel]}_{ano_sel}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.info("Buscando dados no Salesforce...")
