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

def normalizar_cpf_padrao(valor: Any) -> Optional[str]:
    """
    Função Mestra para normalização de CPF.
    1. Remove não numéricos.
    2. Adiciona zeros à esquerda até completar 11 dígitos.
    3. Retorna None se estiver vazio.
    """
    if pd.isna(valor):
        return None
    
    # Remove tudo que não é dígito
    limpo = re.sub(r'\D', '', str(valor))
    
    # Se ficou vazio após limpeza, retorna None
    if not limpo:
        return None
        
    # Adiciona zeros à esquerda (padronização para 11 dígitos)
    # Ex: '1234567890' vira '01234567890'
    return limpo.zfill(11)

def extrair_numero_contrato(texto: str) -> Optional[str]:
    if isinstance(texto, str):
        match_regex = re.search(r'Contrato:\s*(\d+)\s*\(', texto)
        return match_regex.group(1) if match_regex else None
    return None

def extrair_nome_empreendimento(texto: str) -> Optional[str]:
    if isinstance(texto, str):
        match_regex = re.search(r'^\d{11} - ', texto)
        return texto.split(' - ')[1] if match_regex else None
    return None

def extrair_id_fornecedor(texto: str) -> Optional[str]:
    if isinstance(texto, str):
        match_regex = re.search(r'Fornecedor:\s*(\d+)\s*-', texto)
        return match_regex.group(1) if match_regex else None
    return None

def consultar_api_pep(cpf: str) -> str:
    """
    Consulta o CPF na API de PEP.
    O CPF já deve chegar aqui normalizado (string de 11 dígitos),
    mas fazemos uma verificação de segurança.
    """
    # Garante normalização (caso alguém chame a função diretamente sem passar pelo normalizador)
    cpf_tratado = normalizar_cpf_padrao(cpf)

    if not cpf_tratado:
        return "CPF Não Informado"
    
    if cpf_tratado == '00000000000':
        return "CPF Zerado (Inválido)"
        
    # Como já passamos pelo zfill na normalização, o len deve ser 11
    if len(cpf_tratado) != 11:
        return "CPF Inválido (Tam. Incorreto)"

    try:
        api_key = st.secrets["API_KEY"]
    except Exception:
        return "Erro: API_KEY não configurada"

    headers = {"chave-api-dados": api_key}

    try:
        response = requests.get(
            API_PEP_URL, 
            headers=headers,
            params={'cpf': cpf_tratado, 'pagina': 1}, 
            timeout=TIMEOUT_API_SEGUNDOS
        )

        if response.status_code == 200:
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
    lista_dados_estruturados = []
    
    try:
        df_bruto = pd.read_excel(arquivo_obj, header=None)
    except Exception as e:
        st.error(f"Falha na leitura do Excel: {e}")
        return pd.DataFrame()

    empresa_atual = None
    contrato_atual = None
    data_contrato_atual = None

    for index, linha in df_bruto.iterrows():
        primeira_celula = str(linha[0]).strip() if pd.notna(linha[0]) else ""

        if "Empreendimento:" in primeira_celula:
            if len(linha) > 1 and pd.notna(linha[1]):
                empresa_atual = str(linha[1]).strip()
            continue
        
        if primeira_celula.isdigit():
            if len(linha) >= 5:
                contrato_atual = str(linha[3]).strip() if pd.notna(linha[3]) else ""
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
                    "CPF": None,
                    "Data Cadastro": data_contrato_atual
                })
            continue

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
        
        if primeira_celula.startswith("CPF"):
            if lista_dados_estruturados:
                cpf_valor = str(linha[1]).strip() if (len(linha) > 1 and pd.notna(linha[1])) else ""
                lista_dados_estruturados[-1]["CPF"] = cpf_valor
            continue

    df_estruturado = pd.DataFrame(lista_dados_estruturados)

    if df_estruturado.empty:
        return df_estruturado

    df_estruturado = df_estruturado[df_estruturado['Nome da Parte'] != ""]
    termos_ignorar = ["Clientes no Bloco", "Clientes no Empreendimento", "Clientes", "Total Clientes", "Total Geral"]
    df_estruturado = df_estruturado[~df_estruturado['Nome da Parte'].astype(str).str.strip().isin(termos_ignorar)]
    
    # --- CORREÇÃO APLICADA AQUI ---
    # Aplica a normalização UNIFICADA antes de verificar validade ou duplicatas
    df_estruturado['cpf_normalizado'] = df_estruturado['CPF'].apply(normalizar_cpf_padrao)
    
    # Filtra inválidos (None ou '00000000000')
    mask_cpf_problematico = (
        (df_estruturado['cpf_normalizado'].isnull()) | 
        (df_estruturado['cpf_normalizado'] == '00000000000')
    )

    df_invalidos = df_estruturado[mask_cpf_problematico] 
    df_validos = df_estruturado[~mask_cpf_problematico]    
    
    # Remove duplicatas usando a coluna normalizada (agora com zeros à esquerda)
    df_validos = df_validos.drop_duplicates(subset=['cpf_normalizado'], keep='first')
    
    df_final = pd.concat([df_validos, df_invalidos], ignore_index=True)
    # Não dropamos 'cpf_normalizado' ainda pois vamos usá-lo para consulta API
    return df_final

# ==============================================================================
# INTERFACE PRINCIPAL
# ==============================================================================

abas_navegacao = st.tabs(["📊 Fornecedores sem Diligência", "🔎 Verificação de PEPs"])

# ... (ABA 1 Mantida idêntica ao original, omitida para brevidade se não houve mudança lógica necessária lá) ...
# Se precisar do código da Aba 1, avise, mas foquei na correção do PEP abaixo.
with abas_navegacao[0]:
    st.title("📊 Controle Interno - Fornecedores")
    st.info("A lógica desta aba foi mantida. Foque na aba de Verificação de PEPs para as correções de CPF.")
    # (Mantenha o código original da Aba 1 aqui)

# ------------------------------------------------------------------------------
# ABA 2: VERIFICAÇÃO DE PEPs
# ------------------------------------------------------------------------------
with abas_navegacao[1]:
    st.title("🔎 Consulta Unificada de PEPs")
    st.info("Utilize as seções abaixo. A consulta agora padroniza automaticamente CPFs com menos de 11 dígitos.")

    # --------------------------------------------------------------------------
    # SEÇÃO 1: INCORPORADORA
    # --------------------------------------------------------------------------
    st.subheader("🏢 Incorporadora (Relatório ERP)")
    arquivo_bruto_pep = st.file_uploader("Upload Relatório Incorporadora (.xlsx)", type=["xlsx"], key="upload_incorp")

    if st.button("Verificar Base Incorporadora", key="btn_incorp"):
        if arquivo_bruto_pep is None:
            st.warning("Por favor, faça o upload do arquivo.")
        else:
            try:
                st.info("🔹 Fase 1/2: Estruturando dados e normalizando CPFs (Adicionando zeros)...")
                df_clientes_tratado = processar_parser_relatorio_complexo(arquivo_bruto_pep)
                
                if df_clientes_tratado.empty:
                    st.error("O arquivo está vazio ou fora do padrão.")
                else:
                    # Filtra apenas quem tem CPF normalizado válido para consulta
                    mask_consulta = (df_clientes_tratado['cpf_normalizado'].notna()) & \
                                    (df_clientes_tratado['cpf_normalizado'] != '00000000000')
                    
                    lista_cpfs_consulta = df_clientes_tratado.loc[mask_consulta, "cpf_normalizado"].tolist()
                    
                    st.info(f"🔹 Fase 2/2: Consultando {len(lista_cpfs_consulta)} CPFs na API...")

                    # Realiza consulta usando a lista já limpa e preenchida
                    with concurrent.futures.ThreadPoolExecutor(max_workers=MAX_THREADS_API) as executor:
                        resultados_raw = list(executor.map(consultar_api_pep, lista_cpfs_consulta))

                    # Cria dicionário para mapear resultado
                    mapa_resultados = dict(zip(lista_cpfs_consulta, resultados_raw))

                    # Aplica o resultado. Quem não tem CPF normalizado recebe status de erro
                    df_clientes_tratado["Status_PEP"] = df_clientes_tratado["cpf_normalizado"].map(mapa_resultados)
                    df_clientes_tratado.loc[df_clientes_tratado["cpf_normalizado"].isna(), "Status_PEP"] = "CPF Inválido/Não Identificado"

                    st.success("✅ Verificação Incorporadora Finalizada!")
                    st.dataframe(df_clientes_tratado)

                    # Exportação
                    buffer_excel_pep = io.BytesIO()
                    with pd.ExcelWriter(buffer_excel_pep, engine='openpyxl') as writer:
                        df_clientes_tratado.drop(columns=['cpf_normalizado'], inplace=True, errors='ignore')
                        df_clientes_tratado.to_excel(writer, index=False, sheet_name="Analise_PEP")
                    buffer_excel_pep.seek(0)
                    st.download_button("📥 Baixar Relatório", data=buffer_excel_pep, file_name="Resultado_PEP_Incorporadora.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

            except Exception as e:
                st.error(f"Erro crítico: {e}")

    st.markdown("---") 

    # --------------------------------------------------------------------------
    # SEÇÃO 2: AUTOMOTIVO
    # --------------------------------------------------------------------------
    st.subheader("🚗 Automotivo (Planilha Simples)")
    arquivo_bruto_auto = st.file_uploader("Upload Base Automotiva (.xlsx)", type=["xlsx"], key="upload_auto")

    if st.button("Verificar Base Automotiva", key="btn_auto"):
        if arquivo_bruto_auto is None:
            st.warning("Upload necessário.")
        else:
            try:
                st.info("🔹 Lendo arquivo e normalizando dados...")
                df_auto = pd.read_excel(arquivo_bruto_auto)
                df_auto.columns = [c.upper().strip() for c in df_auto.columns]
                
                if 'CPF' not in df_auto.columns:
                    st.error("Coluna CPF não encontrada.")
                else:
                    # --- CORREÇÃO APLICADA AQUI ---
                    # Usa a função mestra que já faz o zfill(11)
                    df_auto['CPF_Normalizado'] = df_auto['CPF'].apply(normalizar_cpf_padrao)
                    
                    # Agora filtramos pelo tamanho do CPF NORMALIZADO (que garantidamente terá 11 se for válido)
                    # Antes você filtrava pelo tamanho do bruto, o que ignorava quem tinha 10 dígitos (falta de zero)
                    mask_cpf_valido = df_auto['CPF_Normalizado'].str.len() == 11
                    
                    cpfs_para_consulta = df_auto.loc[mask_cpf_valido, 'CPF_Normalizado'].unique().tolist()
                    
                    st.info(f"🔹 Consultando {len(cpfs_para_consulta)} CPFs únicos na API...")

                    with concurrent.futures.ThreadPoolExecutor(max_workers=MAX_THREADS_API) as executor:
                        resultados_raw = list(executor.map(consultar_api_pep, cpfs_para_consulta))
                    
                    mapa_resultados = dict(zip(cpfs_para_consulta, resultados_raw))
                    
                    df_auto['Status_PEP'] = df_auto['CPF_Normalizado'].map(mapa_resultados).fillna("CPF Inválido/Formato Incorreto")

                    st.success("✅ Verificação Finalizada!")
                    st.dataframe(df_auto)

                    buffer_excel_auto = io.BytesIO()
                    with pd.ExcelWriter(buffer_excel_auto, engine='openpyxl') as writer:
                        df_auto.drop(columns=['CPF_Normalizado'], inplace=True, errors='ignore')
                        df_auto.to_excel(writer, index=False, sheet_name="Analise_PEP_Auto")
                    buffer_excel_auto.seek(0)
                    st.download_button("📥 Baixar Relatório", data=buffer_excel_auto, file_name="Resultado_PEP_Automotivo.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

            except Exception as e:
                st.error(f"Erro: {e}")


