import streamlit as st
import pandas as pd
from datetime import datetime
from simple_salesforce import Salesforce
import io

# ==========================================
# 0. CONFIGURAÇÃO E ESTILO CORPORATIVO (HEADER REMOVED)
# ==========================================
st.set_page_config(
    page_title="Relatório de Baixa de Ativos", 
    page_icon="📊", 
    layout="wide"
)

# CSS para visual corporativo e REMOÇÃO do cabeçalho/menu do Streamlit
st.markdown("""
    <style>
    /* Esconde o cabeçalho, menu e rodapé padrão */
    #MainMenu {visibility: hidden;}
    header {visibility: hidden;}
    footer {visibility: hidden;}
    
    .main { background-color: #f8f9fa; }
    .stMetric {
        background-color: #ffffff;
        padding: 15px;
        border-radius: 10px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
    }
    h2 { color: #1e3a8a; font-weight: 700; margin-top: 0px; padding-top: 20px; }
    </style>
    """, unsafe_allow_html=True)

st.markdown("## Baixa Manual de Ativos")

# ==========================================
# 1. FUNÇÕES DE SUPORTE
# ==========================================
@st.cache_data(ttl=3600)
def carregar_dados_financeiro():
    usuario = st.secrets["sf_user"]
    senha = st.secrets["sf_pass"]
    token = st.secrets["sf_token"]
    
    sf = Salesforce(username=usuario, password=senha, security_token=token)
    
    # Query atualizada com SerialNumber e FOZ_CodigoItem__c
    query_ativos = """
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
    
    resultado = sf.query_all(query_ativos)
    df = pd.json_normalize(resultado['records'])
    
    if df.empty:
        return pd.DataFrame(), datetime.now()

    df.drop(columns=[col for col in df.columns if 'attributes' in col], inplace=True, errors='ignore')
    
    # Tratamento do Código do Item (Item do Contrato)
    def corrigir_codigo_item(valor):
        if pd.isna(valor) or valor == '': return ''
        numeros = ''.join(filter(str.isdigit, str(valor).split('.')[0]))
        return numeros.zfill(8) if numeros else ''
    
    df['FOZ_CodigoItem__c'] = df['FOZ_CodigoItem__c'].apply(corrigir_codigo_item)
    
    # Datas e Valores
    df['FOZ_Data_Ativacao_Inativacao_Manual__c'] = pd.to_datetime(df['FOZ_Data_Ativacao_Inativacao_Manual__c']).dt.tz_localize(None)
    df['FOZ_ValorTotal__c'] = pd.to_numeric(df['FOZ_ValorTotal__c'], errors='coerce').fillna(0.0)
    
    df['Ano'] = df['FOZ_Data_Ativacao_Inativacao_Manual__c'].dt.year
    df['Mes_Num'] = df['FOZ_Data_Ativacao_Inativacao_Manual__c'].dt.month
    
    df = df.rename(columns={
        'FOZ_ContaRecebedora__r.Name': 'Cliente',
        'FOZ_CodigoItem__c': 'Item do Contrato',
        'SerialNumber': 'Número de Série',
        'Status': 'Status',
        'FOZ_Data_Ativacao_Inativacao_Manual__c': 'Data da Baixa',
        'FOZ_Motivo_Inativacao_Manual__c': 'Motivo da Baixa',
        'FOZ_ValorTotal__c': 'Valor do Contrato (R$)'
    })
    
    return df, datetime.now()

def converter_para_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Baixas_Ativos')
    return output.getvalue()

# ==========================================
# 2. BARRA LATERAL (FILTROS)
# ==========================================
st.sidebar.markdown("### 🎛️ Painel de Controle")
df_base, data_ref = carregar_dados_financeiro()

if not df_base.empty:
    anos_disponiveis = sorted(df_base['Ano'].unique().tolist(), reverse=True)
    ano_selecionado = st.sidebar.selectbox("Ano", options=anos_disponiveis)

    meses_nomes = {
        1: "Janeiro", 2: "Fevereiro", 3: "Março", 4: "Abril", 5: "Maio", 6: "Junho",
        7: "Julho", 8: "Agosto", 9: "Setembro", 10: "Outubro", 11: "Novembro", 12: "Dezembro"
    }
    meses_no_ano = sorted(df_base[df_base['Ano'] == ano_selecionado]['Mes_Num'].unique().tolist())
    mes_selecionado_num = st.sidebar.selectbox(
        "Mês", 
        options=meses_no_ano, 
        format_func=lambda x: meses_nomes[x]
    )

    df_filtrado = df_base[
        (df_base['Ano'] == ano_selecionado) & 
        (df_base['Mes_Num'] == mes_selecionado_num)
    ]

    st.sidebar.divider()
    if st.sidebar.button("🔄 Sincronizar Salesforce"):
        st.cache_data.clear()
        st.rerun()

    st.sidebar.caption(f"Dados atualizados em: \n{data_ref.strftime('%d/%m/%Y %H:%M')}")

# ==========================================
# 3. EXIBIÇÃO E DOWNLOAD
# ==========================================
if not df_base.empty:
    c1, c2, c3 = st.columns(3)
    c1.metric("Volume de Baixas", len(df_filtrado))
    
    total_valor = df_filtrado['Valor do Contrato (R$)'].sum()
    c2.metric("Impacto Financeiro", f"R$ {total_valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
    
    ticket_medio = total_valor / len(df_filtrado) if len(df_filtrado) > 0 else 0
    c3.metric("Ticket Médio", f"R$ {ticket_medio:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
    
    st.divider()

    st.markdown("#### Detalhamento de Ativos")
    df_exibicao = df_filtrado.copy()
    df_exibicao['Data da Baixa'] = df_exibicao['Data da Baixa'].dt.strftime('%d/%m/%Y')
    
    # Incluindo "Item do Contrato" e "Número de Série" no extrato
    colunas_finais = [
        'Cliente', 'Item do Contrato', 'Número de Série', 
        'Status', 'Data da Baixa', 'Motivo da Baixa', 'Valor do Contrato (R$)'
    ]
    df_para_tabela = df_exibicao[colunas_finais]

    st.dataframe(df_para_tabela, use_container_width=True, hide_index=True)

    if not df_filtrado.empty:
        excel_data = converter_para_excel(df_para_tabela)
        nome_arquivo = f"Extrato_Baixas_{meses_nomes[mes_selecionado_num]}_{ano_selecionado}.xlsx"
        
        st.download_button(
            label="📥 Baixar Extrato em Excel",
            data=excel_data,
            file_name=nome_arquivo,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.info("Aguardando carregamento de dados...")
