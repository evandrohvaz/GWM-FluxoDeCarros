import streamlit as st
import pandas as pd
from io import BytesIO

# --- 1. CONFIGURAÇÃO INICIAL E ESTILOS ---
st.set_page_config(layout="wide") 

# CSS para um visual mais limpo e profissional das "vagas"
st.markdown(
    """
    <style>
    .vaga-card {
        border: 2px solid #007BFF; /* Cor primária do tema */
        border-radius: 8px;
        padding: 10px;
        margin-bottom: 10px;
        text-align: center;
        background-color: #F8F9FA; /* Fundo leve */
        box-shadow: 2px 2px 5px rgba(0, 0, 0, 0.1);
        transition: transform 0.2s;
        height: 100%; /* Garante que todos os cards tenham a mesma altura */
    }
    .vaga-card:hover {
        transform: translateY(-3px); /* Efeito de hover */
        box-shadow: 4px 4px 10px rgba(0, 0, 0, 0.2);
    }
    .vaga-title {
        font-size: 1.1em;
        font-weight: bold;
        color: #007BFF;
        margin-bottom: 5px;
    }
    .vaga-subtitle {
        font-size: 0.8em;
        color: #6C757D;
    }
    .empty-vaga {
        border: 2px dashed #CED4DA;
        border-radius: 8px;
        padding: 10px;
        margin-bottom: 10px;
        text-align: center;
        color: #ADB5BD;
        background-color: #FFF;
        height: 100%;
        display: flex;
        align-items: center;
        justify-content: center;
    }
    </style>
    """,
    unsafe_allow_html=True
)

st.title('📊 Dashboard de Rastreamento de Linha de Montagem')
st.markdown('Acompanhe a ocupação em tempo real e exporte a sequência de produção.')
st.markdown('---')

# --- 2. CONFIGURAÇÕES DE DADOS E Mapeamento de Colunas ---
VAGAS_POR_ESTACAO = {
    'PBS_Off': 6,
    'BAIN': 15,
    'BAOFF': 8,
    'AF-IN': 9
}
ordem_estacoes = list(VAGAS_POR_ESTACAO.keys())

# Nomes de colunas conforme o seu arquivo Excel
NOME_BODY = 'Body number'
NOME_ESTACAO = 'Estação de aquisição'
NOME_TEMPO = 'Tempo de aquisição'
NOME_LOTE = 'Número Lote'

# --- Função para gerar o link de download do Excel (USANDO OPENPYXL) ---
def to_excel(df):
    """Converte o DataFrame para um objeto BytesIO em formato Excel usando openpyxl."""
    output = BytesIO()
    # Usa o motor 'openpyxl' para escrita
    writer = pd.ExcelWriter(output, engine='openpyxl') 
    df.to_excel(writer, index=False, sheet_name='SequenciaMontagem')
    writer.close() 
    processed_data = output.getvalue()
    return processed_data

# --- Função para processar e exibir cada estação ---
def exibir_estacao(df_estacao, estacao, vagas):
    """Filtra, exibe as vagas mais recentes, o restante e métricas."""
    
    total_carros = len(df_estacao)
    df_vagas = df_estacao.head(vagas).reset_index(drop=True)
    df_restante = df_estacao.iloc[vagas:]
    
    # --- Métrica de Ocupação ---
    col1, col2 = st.columns([1, 4])
    with col1:
        st.metric(
            label=f"Ocupação {estacao}",
            value=f"{len(df_vagas)}/{vagas}",
            delta=f"Fila: {len(df_restante)} carros",
            delta_color="off" if len(df_restante) == 0 else "inverse"
        )
    with col2:
        st.subheader(f'🏭 Estação: {estacao} ({total_carros} Carros no Total)')

    # --- Exibição Visual (Vagas Enfileiradas) ---
    st.markdown(f"**{vagas} Vagas Mais Recentes:**")
    
    # Define 6 colunas para ser amigável em telas menores, Streamlit gerencia quebras
    cols = st.columns(6) 
    
    for i in range(vagas):
        col = cols[i % 6] # Reutiliza as 6 colunas
        
        if i < len(df_vagas):
            body = df_vagas.loc[i, NOME_BODY]
            lote = df_vagas.loc[i, NOME_LOTE]
            
            # --- ATUALIZAÇÃO: Inclui a data e a hora na formatação ---
            tempo = df_vagas.loc[i, NOME_TEMPO].strftime('%d/%m %H:%M:%S') if pd.notna(df_vagas.loc[i, NOME_TEMPO]) else 'S/ Tempo'
            
            # HTML estilizado para a vaga ocupada (Lote e Data/Hora adicionados)
            vaga_html = f"""
            <div class="vaga-card">
                <div class="vaga-title">{body}</div>
                <div class="vaga-subtitle">Lote: **{lote}**</div>
                <div class="vaga-subtitle">Entrada: {tempo}</div>
            </div>
            """
            col.markdown(vaga_html, unsafe_allow_html=True)
        else:
            # HTML estilizado para a vaga vazia
            col.markdown('<div class="empty-vaga">Vaga Vazia</div>', unsafe_allow_html=True)

    # --- Exibição dos Carros Mais Antigos (Fila) ---
    if not df_restante.empty:
        with st.expander(f"➕ Fila de Espera (Carros mais antigos - {len(df_restante)})"):
            st.dataframe(
                df_restante[[NOME_BODY, NOME_TEMPO, NOME_LOTE]],
                use_container_width=True
            )
    
    st.markdown('***') # Separador visual

# --- UPLOAD DO ARQUIVO EXCEL ---
uploaded_file = st.file_uploader(
    "📥 Escolha seu arquivo Excel (xlsx)",
    type=['xlsx']
)

df_sequenciado_final = None # Inicializa a variável para o botão de exportação

if uploaded_file is not None:
    try:
        # 3. LEITURA, CONVERSÃO E ORDENAÇÃO
        df = pd.read_excel(uploaded_file)
        
        # 3.1. Validação de Colunas
        for col in [NOME_BODY, NOME_ESTACAO, NOME_TEMPO, NOME_LOTE]:
            if col not in df.columns:
                raise KeyError(f"Coluna '{col}' não encontrada no arquivo Excel.")

        # 3.2. Pré-processamento
        df[NOME_TEMPO] = pd.to_datetime(df[NOME_TEMPO], errors='coerce')
        df.dropna(subset=[NOME_TEMPO], inplace=True)
        
        df_filtrado = df[df[NOME_ESTACAO].isin(ordem_estacoes)].copy()
        
        # 3.3. Força a ordem categórica das estações (chave para a sequência)
        df_filtrado[NOME_ESTACAO] = pd.Categorical(
            df_filtrado[NOME_ESTACAO], 
            categories=ordem_estacoes, 
            ordered=True
        )

        # 3.4. Ordenação Final: Estação (Fixa) -> Tempo (Mais Novo)
        df_sequenciado = df_filtrado.sort_values(
            by=[NOME_ESTACAO, NOME_TEMPO], 
            ascending=[True, False]
        ).reset_index(drop=True)
        
        df_sequenciado_final = df_sequenciado[[NOME_BODY, NOME_ESTACAO, NOME_TEMPO, NOME_LOTE]]

        st.success('Dados processados e prontos para visualização! ✅')
        st.markdown('---')

        # --- 4. EXIBIÇÃO VISUAL E EXPORTAÇÃO ---

        # Botão de Exportação (Visível apenas após o processamento)
        st.download_button(
            label="⬇️ Exportar Sequência Completa (.xlsx)",
            data=to_excel(df_sequenciado_final),
            file_name=f'Sequencia_Montagem_{pd.Timestamp.now().strftime("%Y%m%d_%H%M")}.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            key='download_excel_button'
        )
        st.markdown('***')

        # Exibe as estações
        for estacao in ordem_estacoes:
            df_estacao = df_sequenciado[df_sequenciado[NOME_ESTACAO] == estacao]
            vagas = VAGAS_POR_ESTACAO.get(estacao, 0)
            
            if not df_estacao.empty:
                exibir_estacao(df_estacao, estacao, vagas)
            else:
                st.info(f'Nenhum carro encontrado na estação: {estacao}.')


    except KeyError as e:
        st.error(f"Erro: Coluna {e} não encontrada. Verifique se os cabeçalhos são exatamente: '{NOME_BODY}', '{NOME_ESTACAO}', '{NOME_TEMPO}' e '{NOME_LOTE}'.")
    except Exception as e:
        st.error(f"Ocorreu um erro inesperado: {e}")

else:
    st.info('Aguardando o upload do arquivo Excel para iniciar o processamento.')