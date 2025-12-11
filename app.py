import streamlit as st
import pandas as pd
import re
import io
import requests
import concurrent.futures
from typing import Optional, List, Dict, Any

# ==============================================================================
# CONFIGURAÇÃO DA PÁGINA E ESTILOS
# ==============================================================================
st.set_page_config(
    page_title="Portal de Compliance & Controles Internos",
    page_icon="🛡️",
    layout="wide"
)

# Constantes de API e Configuração
API_PEP_URL = "https://api.portaldatransparencia.gov.br/api-de-dados/peps"
TIMEOUT_API_SEGUNDOS = 10
MAX_THREADS_API = 20

# ==============================================================================
# FUNÇÕES UTILITÁRIAS (HELPER FUNCTIONS)
# ==============================================================================

def extrair_numero_contrato(texto: str) -> Optional[str]:
    """
    Extrai o número do contrato de uma string padrão do relatório.
    Padrão esperado: 'Contrato: 12345 ('
    """
    if isinstance(texto, str):
        match_regex = re.search(r'Contrato:\s*(\d+)\s*\(', texto)
        return match_regex.group(1) if match_regex else None
    return None

def extrair_nome_empreendimento(texto: str) -> Optional[str]:
    """
    Remove o código numérico inicial do nome do empreendimento.
    Padrão esperado: '00000000000 - NOME DO EMPREENDIMENTO'
    """
    if isinstance(texto, str):
        match_regex = re.search(r'^\d{11} - ', texto)
        return texto.split(' - ')[1] if match_regex else None
    return None

def extrair_id_fornecedor(texto: str) -> Optional[str]:
    """
    Captura o ID do fornecedor no início da string de descrição.
    Padrão esperado: 'Fornecedor: 12345 -'
    """
    if isinstance(texto, str):
        match_regex = re.search(r'Fornecedor:\s*(\d+)\s*-', texto)
        return match_regex.group(1) if match_regex else None
    return None

def sanitizar_documento(valor: Any) -> str:
    """
    Remove caracteres não numéricos de CPF/CNPJ para padronização.
    """
    if pd.isna(valor):
        return ""
    return re.sub(r'\D', '', str(valor))

def consultar_api_pep(cpf: str) -> str:
    """
    Consulta o CPF na API de Pessoas Expostas Politicamente (PEP) do Portal da Transparência.
    Requer chave de API configurada no st.secrets.
    """
    # Validações iniciais básicas
    if not cpf or str(cpf).lower() == 'nan':
        return "CPF Não Informado"
    
    cpf_limpo = re.sub(r'\D', '', str(cpf))
    
    if cpf_limpo == '00000000000':
        return "CPF Zerado (Inválido)"
    if len(cpf_limpo) != 11:
        return "CPF Inválido/Incompleto"

    # Busca a chave de API de forma segura
    try:
        api_key = st.secrets["API_KEY"]
    except Exception:
        return "Erro: API_KEY não configurada"

    headers = {"chave-api-dados": api_key}

    try:
        response = requests.get(
            API_PEP_URL, 
            headers=headers,
            params={'cpf': cpf_limpo, 'pagina': 1}, 
            timeout=TIMEOUT_API_SEGUNDOS
        )

        if response.status_code == 200:
            # Se a lista retornada não for vazia, é PEP
            return "Sim - PEP Identificado" if response.json() else "Não consta"
        elif response.status_code == 404:
            return "Não consta" 
        elif response.status_code == 429:
            return "Erro: Limite de requisições excedido"
        elif response.status_code in [401, 403]:
            return "Erro: Falha de Autenticação API"
        else:
            return f"Erro HTTP {response.status_code}"

    except requests.exceptions.RequestException:
        return "Erro de Conexão"

def processar_parser_relatorio_complexo(arquivo_obj) -> pd.DataFrame:
    """
    Lê um relatório Excel hierárquico (não tabular) e transforma em DataFrame estruturado.
    Identifica Empreendimento, Contrato e Partes (Cliente, Cônjuge, Participante).
    """
    lista_dados_estruturados = []
    
    try:
        # Lê sem cabeçalho pois o arquivo é posicional/hierárquico
        df_bruto = pd.read_excel(arquivo_obj, header=None)
    except Exception as e:
        st.error(f"Falha na leitura do Excel: {e}")
        return pd.DataFrame()

    # Variáveis de estado para manter contexto durante a iteração das linhas
    empresa_atual = None
    contrato_atual = None
    data_contrato_atual = None

    for index, linha in df_bruto.iterrows():
        primeira_celula = str(linha[0]).strip() if pd.notna(linha[0]) else ""

        # Detecção de mudança de contexto: Empreendimento
        if "Empreendimento:" in primeira_celula:
            if len(linha) > 1 and pd.notna(linha[1]):
                empresa_atual = str(linha[1]).strip()
            continue
        
        # Detecção de mudança de contexto: Contrato (identificado por ID numérico na col 0)
        if primeira_celula.isdigit():
            if len(linha) >= 5:
                contrato_atual = str(linha[3]).strip() if pd.notna(linha[3]) else ""
                
                # Tratamento de data
                data_bruta = linha[4]
                if isinstance(data_bruta, pd.Timestamp):
                    data_contrato_atual = data_bruta.strftime('%Y-%m-%d')
                else:
                    data_contrato_atual = str(data_bruta).strip() if pd.notna(data_bruta) else ""

                nome_cliente = str(linha[1]).strip() if pd.notna(linha[1]) else ""
                
                lista_dados_estruturados.append({
                    "Numero Contrato": contrato_atual,
                    "Empreendimento": empresa_atual,
                    "Nome da Parte": nome_cliente,
                    "Tipo Parte": "CLIENTE PRINCIPAL",
                    "CPF": None, # Será preenchido nas linhas subsequentes
                    "Data Cadastro": data_contrato_atual
                })
            continue

        # Detecção de partes relacionadas
        if "Cônjuge:" in primeira_celula:
            nome_conjuge = str(linha[1]).strip() if (len(linha) > 1 and pd.notna(linha[1])) else ""
            lista_dados_estruturados.append({
                "Numero Contrato": contrato_atual,
                "Empreendimento": empresa_atual,
                "Nome da Parte": nome_conjuge,
                "Tipo Parte": "CONJUGE",
                "CPF": None,
                "Data Cadastro": data_contrato_atual
            })
            continue

        if "Participante:" in primeira_celula:
            nome_participante = str(linha[1]).strip() if (len(linha) > 1 and pd.notna(linha[1])) else ""
            lista_dados_estruturados.append({
                "Numero Contrato": contrato_atual,
                "Empreendimento": empresa_atual,
                "Nome da Parte": nome_participante,
                "Tipo Parte": "PARTICIPANTE",
                "CPF": None,
                "Data Cadastro": data_contrato_atual
            })
            continue
        
        # Associação do CPF à última parte adicionada
        if primeira_celula.startswith("CPF"):
            if lista_dados_estruturados:
                cpf_valor = str(linha[1]).strip() if (len(linha) > 1 and pd.notna(linha[1])) else ""
                lista_dados_estruturados[-1]["CPF"] = cpf_valor
            continue

    df_estruturado = pd.DataFrame(lista_dados_estruturados)

    if df_estruturado.empty:
        return df_estruturado

    # Limpeza final de dados
    df_estruturado = df_estruturado[df_estruturado['Nome da Parte'] != ""]
    
    # Remoção de linhas de lixo/totais que podem ter sido capturadas
    termos_ignorar = ["Clientes no Bloco", "Clientes no Empreendimento", "Clientes", "Total Clientes", "Total Geral"]
    df_estruturado = df_estruturado[~df_estruturado['Nome da Parte'].astype(str).str.strip().isin(termos_ignorar)]
    
    # Lógica de Dedup: Mantém apenas um registro por CPF válido para economizar requisições API
    # Mas mantém registros com CPFs inválidos/nulos para reportar erro manual
    df_estruturado['cpf_temp_sanitizado'] = df_estruturado['CPF'].astype(str).str.replace(r'\D', '', regex=True)
    
    mask_cpf_problematico = (
        (df_estruturado['cpf_temp_sanitizado'] == '') | 
        (df_estruturado['cpf_temp_sanitizado'].isnull()) | 
        (df_estruturado['cpf_temp_sanitizado'] == '00000000000') |
        (df_estruturado['CPF'].astype(str).str.lower().str.contains('nan'))
    )

    df_invalidos = df_estruturado[mask_cpf_problematico] 
    df_validos = df_estruturado[~mask_cpf_problematico]    
    
    # Remove duplicatas apenas dos válidos
    df_validos = df_validos.drop_duplicates(subset=['cpf_temp_sanitizado'], keep='first')
    
    df_final = pd.concat([df_validos, df_invalidos], ignore_index=True)
    df_final = df_final.drop(columns=['cpf_temp_sanitizado'])

    return df_final

# ==============================================================================
# INTERFACE PRINCIPAL
# ==============================================================================

abas_navegacao = st.tabs(["📊 Fornecedores sem Diligência", "🔎 Verificação de PEPs"])

# ------------------------------------------------------------------------------
# ABA 1: FORNECEDORES
# ------------------------------------------------------------------------------
with abas_navegacao[0]:
    st.title("📊 Controle Interno - Auditoria de Fornecedores")
    st.markdown("Cruza dados financeiros com a base de *Due Diligence* para identificar pagamentos a fornecedores não homologados.")

    col1, col2, col3, col4 = st.columns(4)
    with col1:
        arquivo_contratos = st.file_uploader("Relatório de Contratos (.xlsx)", type="xlsx")
    with col2:
        arquivo_itens = st.file_uploader("Relatório de Itens (.xlsx)", type="xlsx")
    with col3:
        arquivo_agentes = st.file_uploader("Base de Agentes (.xlsx)", type="xlsx")
    with col4:
        arquivo_diligencias = st.file_uploader("Controle de Diligências (.xlsx)", type="xlsx")

    if st.button("Executar Cruzamento de Dados"):
        if arquivo_contratos and arquivo_itens and arquivo_agentes and arquivo_diligencias:
            try:
                with st.spinner("Normalizando dados e cruzando bases..."):
                    # Carregamento
                    df_contratos = pd.read_excel(arquivo_contratos)
                    df_itens = pd.read_excel(arquivo_itens)
                    df_agentes = pd.read_excel(arquivo_agentes)
                    df_diligencias = pd.read_excel(arquivo_diligencias, sheet_name="DDR e BCK")

                    # Padronização de Colunas (De: Layout Sistema ERP -> Para: Layout Interno)
                    mapa_colunas_erp = {
                        '(a)': 'item_codigo',
                        '(b)': 'descricao_item',
                        '(c)': 'unid_medida',
                        '(d)': 'quantidade',
                        '(e)': 'valor_unitario',
                        '(f)': 'valor_total',
                        '(k = g - j)': 'quantidade_saldo',
                        '(l = i + k)': 'valor_saldo'
                    }
                    df_contratos.rename(columns=mapa_colunas_erp, inplace=True)

                    # Limpeza de linhas de cabeçalho/rodapé do relatório
                    filtro_lixo = 'Total Geral:|Total do Projeto:|Exibir Apropriações|Parâmetros Selecionados|OCMEG'
                    df_contratos = df_contratos[~df_contratos['item_codigo'].astype(str).str.contains(filtro_lixo, na=False)]

                    # Preenchimento de dados hierárquicos (Forward Fill)
                    df_contratos['Contrato_Ref'] = df_contratos['item_codigo'].apply(extrair_numero_contrato).ffill()
                    df_contratos['Empreendimento_Ref'] = df_contratos['item_codigo'].apply(extrair_nome_empreendimento).ffill()
                    df_contratos['ID_Fornecedor'] = df_contratos['descricao_item'].apply(extrair_id_fornecedor).ffill()

                    # Remoção de colunas de cálculo do Excel que não são necessárias
                    cols_descarte = ['(g = e - f)', '(h)', '(i)', '(j)', 'Unnamed: 12', 'Unnamed: 13', 'Unnamed: 14', 'Unnamed: 15', 'Unnamed: 16']
                    df_contratos = df_contratos.drop(columns=[c for c in cols_descarte if c in df_contratos.columns])
                    
                    # Remove linhas divisórias
                    df_contratos = df_contratos[~df_contratos['item_codigo'].astype(str).str.contains('-', na=False)]

                    # Tipagem
                    df_contratos['ID_Fornecedor'] = df_contratos['ID_Fornecedor'].astype(str)
                    df_agentes['Código'] = df_agentes['Código'].astype(str)

                    # Criação de chaves de cruzamento (CNPJ limpo)
                    df_agentes['CNPJ_Sanitizado'] = df_agentes['CNPJ'].apply(sanitizar_documento)
                    df_diligencias['CNPJ_Sanitizado'] = df_diligencias['CNPJ/CPF'].apply(sanitizar_documento)
                    
                    # Cria base única de diligências para evitar duplicatas no merge
                    df_diligencias_unicas = df_diligencias[['CNPJ_Sanitizado']].drop_duplicates()

                    # 1. Enrich com descrição do Item
                    df_contratos = pd.merge(
                        df_contratos,
                        df_itens[['Cód.Item', 'Definição Item']],
                        left_on='item_codigo',
                        right_on='Cód.Item',
                        how='left'
                    )

                    # 2. Enrich com dados do Agente (Fornecedor)
                    df_contratos = pd.merge(
                        df_contratos,
                        df_agentes[['Código', 'CNPJ', 'CNPJ_Sanitizado', 'Nome fantasia']],
                        left_on='ID_Fornecedor',
                        right_on='Código',
                        how='left'
                    )

                    # Conversão numérica para agregações
                    cols_monetarias = ['valor_unitario', 'valor_total', 'valor_saldo', 'quantidade', 'quantidade_saldo']
                    for col in cols_monetarias:
                        df_contratos[col] = pd.to_numeric(df_contratos[col], errors='coerce').fillna(0)

                    # --- GERAÇÃO DE VISÕES (TABELAS) ---

                    # Visão 1: Agrupado por Contrato
                    cols_group_contrato = ['Contrato_Ref', 'Empreendimento_Ref', 'ID_Fornecedor', 'Nome fantasia', 'CNPJ']
                    df_visao_contratos = df_contratos.groupby(cols_group_contrato)[cols_monetarias].sum().reset_index()
                    df_visao_contratos.drop(columns=['valor_unitario', 'quantidade_saldo', 'quantidade'], inplace=True)

                    # Visão 2: Agrupado por Fornecedor (Foco do Compliance)
                    cols_group_fornecedor = ['ID_Fornecedor', 'Nome fantasia', 'CNPJ', 'CNPJ_Sanitizado']
                    df_visao_fornecedores = df_contratos.groupby(cols_group_fornecedor)[cols_monetarias].sum().reset_index()

                    # Cálculo específico: Total gasto apenas com SERVIÇOS
                    filtro_servicos = df_contratos[df_contratos['Definição Item'] == 'Serviços']
                    soma_apenas_servicos = filtro_servicos.groupby('ID_Fornecedor')['valor_total'].sum().reset_index()
                    soma_apenas_servicos.rename(columns={'valor_total': 'Total_Apenas_Servicos'}, inplace=True)

                    df_visao_fornecedores = pd.merge(df_visao_fornecedores, soma_apenas_servicos, left_on='ID_Fornecedor', right_on='ID_Fornecedor', how='left')
                    df_visao_fornecedores['Total_Apenas_Servicos'] = df_visao_fornecedores['Total_Apenas_Servicos'].fillna(0)

                    # Flag de Criticidade: Gastos > 20k
                    df_visao_fornecedores['Flag_Risco_20k'] = df_visao_fornecedores['Total_Apenas_Servicos'].apply(
                        lambda x: '⚠️ Atinge 20 mil' if x >= 20000 else 'Baixo Valor'
                    )

                    # Cruzamento Final: O fornecedor tem diligência feita?
                    df_visao_fornecedores = pd.merge(
                        df_visao_fornecedores,
                        df_diligencias_unicas,
                        on='CNPJ_Sanitizado',
                        how='left',
                        indicator=True
                    )
                    
                    df_visao_fornecedores['Status_Diligencia'] = df_visao_fornecedores['_merge'].apply(
                        lambda x: '✅ Realizada' if x == 'both' else '❌ Pendente'
                    )

                    # Limpeza final para exportação
                    df_visao_fornecedores.drop(columns=['_merge', 'CNPJ_Sanitizado', 'valor_unitario', 'quantidade_saldo', 'quantidade'], inplace=True)
                    df_contratos.drop(columns=['CNPJ_Sanitizado'], inplace=True)

                    # Buffer de memória para download
                    buffer_excel = io.BytesIO()
                    with pd.ExcelWriter(buffer_excel, engine='openpyxl') as writer:
                        df_contratos.to_excel(writer, sheet_name='Analitico_Detalhado', index=False)
                        df_visao_contratos.to_excel(writer, sheet_name='Sintetico_Contratos', index=False)
                        df_visao_fornecedores.to_excel(writer, sheet_name='Matriz_Compliance', index=False)
                    
                    buffer_excel.seek(0)

                st.success("Auditoria processada com sucesso!")
                st.download_button(
                    label="📥 Baixar Relatório de Compliance Consolidado",
                    data=buffer_excel,
                    file_name="Relatorio_Compliance_Fornecedores.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

                st.markdown("### 📋 Matriz de Riscos (Fornecedores)")
                st.dataframe(df_visao_fornecedores.head(10))

            except Exception as e:
                st.error(f"Erro no processamento dos arquivos: {e}")
        else:
            st.warning("⚠️ Atenção: Todos os 4 arquivos são obrigatórios para o cruzamento de dados.")

# ------------------------------------------------------------------------------
# ABA 2: VERIFICAÇÃO DE PEPs (SEPARADO POR SETOR)
# ------------------------------------------------------------------------------
with abas_navegacao[1]:
    st.title("🔎 Consulta Unificada de PEPs")
    st.info("Utilize as seções abaixo de acordo com a origem do arquivo (Incorporadora ou Automotivo). A consulta é feita na API do Portal da Transparência.")

    # --------------------------------------------------------------------------
    # SEÇÃO 1: INCORPORADORA (RELATÓRIO HIERÁRQUICO)
    # --------------------------------------------------------------------------
    st.subheader("🏢 Setor Incorporadora (Relatório ERP)")
    st.markdown("""
    **Layout esperado:** Relatório hierárquico complexo extraído do sistema ERP.  
    O sistema irá tratar a estrutura, remover duplicatas e consultar os CPFs válidos.
    """)

    arquivo_bruto_pep = st.file_uploader("Upload Relatório Incorporadora (.xlsx)", type=["xlsx"], key="upload_incorp")

    if st.button("Verificar Base Incorporadora", key="btn_incorp"):
        if arquivo_bruto_pep is None:
            st.warning("Por favor, faça o upload do arquivo da incorporadora primeiro.")
        else:
            try:
                # Passo 1: Processamento Local
                st.info("🔹 Fase 1/2: Estruturando dados e sanitizando CPFs...")
                df_clientes_tratado = processar_parser_relatorio_complexo(arquivo_bruto_pep)
                
                if df_clientes_tratado.empty:
                    st.error("O arquivo está vazio ou não segue o padrão esperado.")
                else:
                    total_registros = len(df_clientes_tratado)
                    st.success(f"Base processada: {total_registros} registros únicos identificados.")
                    
                    # Passo 2: Consulta Externa API
                    st.info("🔹 Fase 2/2: Consultando API do Governo Federal...")
                    
                    lista_cpfs_consulta = df_clientes_tratado["CPF"].astype(str).tolist()

                    with concurrent.futures.ThreadPoolExecutor(max_workers=MAX_THREADS_API) as executor:
                        resultados_api = list(executor.map(consultar_api_pep, lista_cpfs_consulta))

                    df_clientes_tratado["Status_PEP"] = resultados_api

                    st.success("✅ Verificação Incorporadora Finalizada!")
                    st.dataframe(df_clientes_tratado)

                    # Exportação
                    buffer_excel_pep = io.BytesIO()
                    with pd.ExcelWriter(buffer_excel_pep, engine='openpyxl') as writer:
                        df_clientes_tratado.to_excel(writer, index=False, sheet_name="Analise_PEP")
                    
                    buffer_excel_pep.seek(0)

                    st.download_button(
                        "📥 Baixar Relatório (Incorporadora)",
                        data=buffer_excel_pep,
                        file_name="Resultado_PEP_Incorporadora.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="dl_incorp"
                    )

            except Exception as e:
                st.error(f"Erro crítico: {e}")

    # Divisor visual entre as seções
    st.markdown("---") 

    # --------------------------------------------------------------------------
    # SEÇÃO 2: AUTOMOTIVO (TABELA SIMPLES)
    # --------------------------------------------------------------------------
    st.subheader("🚗 Setor Automotivo (Planilha Simples)")
    st.markdown("""
    **Layout esperado:** Planilha Excel simples com as colunas obrigatórias:  
    `CLIENTE`, `NOME`, `CPF` (sem formatação).
    """)

    arquivo_bruto_auto = st.file_uploader("Upload Base Automotiva (.xlsx)", type=["xlsx"], key="upload_auto")

    if st.button("Verificar Base Automotiva", key="btn_auto"):
        if arquivo_bruto_auto is None:
            st.warning("Por favor, faça o upload do arquivo automotivo primeiro.")
        else:
            try:
                st.info("🔹 Lendo arquivo e normalizando dados...")
                
                df_auto = pd.read_excel(arquivo_bruto_auto)
                
                # Normalização dos nomes das colunas (para evitar erro de case sensitivity)
                df_auto.columns = [c.upper().strip() for c in df_auto.columns]
                
                colunas_esperadas = ['CLIENTE', 'NOME', 'CPF']
                
                # Validação de Colunas
                if not all(col in df_auto.columns for col in colunas_esperadas):
                    st.error(f"Erro de Layout. Colunas obrigatórias não encontradas: {colunas_esperadas}")
                    st.write("Colunas encontradas:", list(df_auto.columns))
                else:
                    # Sanitização
                    df_auto['CPF_Limpo'] = df_auto['CPF'].apply(sanitizar_documento)
                    
                    # Filtra apenas CPFs que parecem válidos (tamanho 11) para não gastar API à toa
                    # Mas mantemos os inválidos no dataframe final marcados
                    mask_cpf_valido = df_auto['CPF_Limpo'].str.len() == 11
                    
                    cpfs_para_consulta = df_auto.loc[mask_cpf_valido, 'CPF_Limpo'].tolist()
                    
                    st.info(f"🔹 Consultando {len(cpfs_para_consulta)} CPFs na API...")

                    # Consulta API
                    with concurrent.futures.ThreadPoolExecutor(max_workers=MAX_THREADS_API) as executor:
                        resultados_raw = list(executor.map(consultar_api_pep, cpfs_para_consulta))
                    
                    # Mapear resultados de volta para o DataFrame
                    # Criamos um dicionário {cpf: resultado} para mapear com segurança
                    mapa_resultados = dict(zip(cpfs_para_consulta, resultados_raw))
                    
                    df_auto['Status_PEP'] = df_auto['CPF_Limpo'].map(mapa_resultados).fillna("CPF Inválido/Formato Incorreto")

                    st.success("✅ Verificação Automotiva Finalizada!")
                    
                    # Reordena colunas para visualização
                    cols_final = ['CLIENTE', 'NOME', 'CPF', 'Status_PEP']
                    # Adiciona colunas extras se houverem no arquivo original
                    cols_extras = [c for c in df_auto.columns if c not in cols_final and c != 'CPF_Limpo']
                    df_auto_final = df_auto[cols_final + cols_extras]

                    st.dataframe(df_auto_final)

                    # Exportação
                    buffer_excel_auto = io.BytesIO()
                    with pd.ExcelWriter(buffer_excel_auto, engine='openpyxl') as writer:
                        df_auto_final.to_excel(writer, index=False, sheet_name="Analise_PEP_Auto")
                    
                    buffer_excel_auto.seek(0)

                    st.download_button(
                        "📥 Baixar Relatório (Automotivo)",
                        data=buffer_excel_auto,
                        file_name="Resultado_PEP_Automotivo.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="dl_auto"
                    )

            except Exception as e:
                st.error(f"Erro ao processar arquivo automotivo: {e}")


