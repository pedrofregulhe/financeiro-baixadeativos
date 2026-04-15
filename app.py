import streamlit as st
import pandas as pd
from datetime import datetime
from simple_salesforce import Salesforce

# ==========================================
# 0. CONFIGURAÇÃO DA PÁGINA
# ==========================================
st.set_page_config(
    page_title="Dashboard - Baixa de Ativos", 
    page_icon="📉", 
    layout="wide"
)

# Título e Subtítulo
st.title("📉 Relatório Financeiro: Baixa Manual de Ativos")

# ==========================================
# 1. FUNÇÃO DE EXTRAÇÃO (COM CACHE E TIMESTAMPS)
# ==========================================
@st.cache_data(ttl=3600)
def carregar_dados_sf():
    # Busca credenciais dos Secrets
    usuario = st.secrets["sf_user"]
    senha = st.secrets["sf_pass"]
    token = st.secrets["sf_token"]
    
    sf = Salesforce(username=usuario, password=senha, security_token=token)
    
    query_ativos = """
    SELECT 
        FOZ_CodigoItem__c,
        Status,
        FOZ_Data_Ativacao_Inativacao_Manual__c,
        FOZ_Motivo_Inativacao_Manual__c,
        FOZ_ValorTotal__c,
        FOZ_ContaRecebedora__r.Name
    FROM Asset
    WHERE FOZ_Motivo_Desinstalacao__c = 'BAIXA MANUAL DE ATIVO'
    """
    
    resultado = sf.query_all(query_ativos)
    df = pd.json_normalize(resultado['records'])
    
    if df.empty:
        return pd.DataFrame(), datetime.now()

    # Tratamentos de dados
    df.drop(columns=[col for col in df.columns if 'attributes' in col], inplace=True, errors='ignore')
    df['FOZ_Data_Ativacao_Inativacao_Manual__c'] = pd.to_datetime(df['FOZ_Data_Ativacao_Inativacao_Manual__c']).dt.tz_localize(None)
    df['FOZ_ValorTotal__c'] = pd.to_numeric(df['FOZ_ValorTotal__c'], errors='coerce').fillna(0.0)
    
    # Renomeação
    df = df.rename(columns={
        'FOZ_ContaRecebedora__r.Name': 'Cliente',
        'FOZ_CodigoItem__c': 'Número de Série',
        'Status': 'Status',
        'FOZ_Data_Ativacao_Inativacao_Manual__c': 'Data da Baixa',
        'FOZ_Motivo_Inativacao_Manual__c': 'Motivo da Baixa',
        'FOZ_ValorTotal__c': 'Valor do Contrato (R$)'
    })
    
    # Retorna o DataFrame e o momento exato da carga
    return df, datetime.now()

# ==========================================
# 2. CONTROLES DA BARRA LATERAL (SIDEBAR)
# ==========================================
st.sidebar.header("⚙️ Controles")

# Botão para atualizar resultados (limpa o cache e recarrega a página)
if st.sidebar.button("🔄 Atualizar Dados"):
    st.cache_data.clear()
    st.rerun()

# Carregamento dos dados
with st.spinner("Puxando dados do Salesforce..."):
    df_base, data_atualizacao = carregar_dados_sf()

# Exibição do Texto de Última Atualização
st.sidebar.write(f"**Última atualização:** \n{data_atualizacao.strftime('%d/%m/%Y %H:%M:%S')}")

# ==========================================
# 3. FILTROS E VISUALIZAÇÃO
# ==========================================
if not df_base.empty:
    # Filtro de Período
    min_d = df_base['Data da Baixa'].min().date()
    max_d = df_base['Data da Baixa'].max().date()
    
    periodo = st.sidebar.date_input("Filtrar Período", [min_d, max_d])
    
    # Lógica de Filtro
    if len(periodo) == 2:
        df_filtrado = df_base[
            (df_base['Data da Baixa'].dt.date >= periodo[0]) & 
            (df_base['Data da Baixa'].dt.date <= periodo[1])
        ]
    else:
        df_filtrado = df_base

    # Métricas Principais
    st.divider()
    c1, c2 = st.columns(2)
    c1.metric("Ativos Baixados", len(df_filtrado))
    c2.metric("Impacto Financeiro", f"R$ {df_filtrado['Valor do Contrato (R$)'].sum():,.2f}")
    
    # Tabela de Dados
    st.subheader("Detalhamento")
    df_visualizacao = df_filtrado.copy()
    df_visualizacao['Data da Baixa'] = df_visualizacao['Data da Baixa'].dt.strftime('%d/%m/%Y')
    st.dataframe(df_visualizacao, use_container_width=True, hide_index=True)
else:
    st.info("Nenhuma baixa manual encontrada.")