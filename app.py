import streamlit as st
import pandas as pd
import re
import io
import requests
import concurrent.futures
import csv


st.set_page_config(page_title="Portal de Compliance", layout="wide")


aba = st.tabs(["📊 Fornecedores sem Diligência", "🔎 Verificação de PEPs Incorporadora"])



with aba[0]:

    st.title("📊 Controle Interno - Fornecedores sem Diligência")
    st.markdown("Faça o upload das quatro planilhas abaixo.")

    col1, col2, col3, col4 = st.columns(4)
    with col1:
        file_contratos = st.file_uploader("Upload Contratos (.xlsx)", type="xlsx")
    with col2:
        file_itens = st.file_uploader("Upload Itens (.xlsx)", type="xlsx")
    with col3:
        file_agentes = st.file_uploader("Upload Agentes (.xlsx)", type="xlsx")
    with col4:
        file_diligencias = st.file_uploader("Upload Diligências (.xlsx)", type="xlsx")


    def extrair_contrato(texto):
        if isinstance(texto, str):
            r = re.search(r'Contrato:\s*(\d+)\s*\(', texto)
            return r.group(1) if r else None
        return None

    def extrair_empreendimento(texto):
        if isinstance(texto, str):
            r = re.search(r'^\d{11} - ', texto)
            return texto.split(' - ')[1] if r else None
        return None

    def extrair_fornecedor(texto):
        if isinstance(texto, str):
            r = re.search(r'Fornecedor:\s*(\d+)\s*-', texto)
            return r.group(1) if r else None
        return None

    def limpar_documento(valor):
        if pd.isna(valor):
            return ""
        return re.sub(r'\D', '', str(valor))


    if st.button("Processar Arquivos"):

        if file_contratos and file_itens and file_agentes and file_diligencias:

            try:
                with st.spinner("Processando arquivos..."):

                    contratos = pd.read_excel(file_contratos)
                    itens = pd.read_excel(file_itens)
                    agentes = pd.read_excel(file_agentes)
                    diligencias = pd.read_excel(file_diligencias, sheet_name="DDR e BCK")

                    contratos.rename(columns={
                        '(a)': 'item_codigo',
                        '(b)': 'descricao_item',
                        '(c)': 'unid_medida',
                        '(d)': 'quantidade',
                        '(e)': 'valor_unitario',
                        '(f)': 'valor_total',
                        '(k = g - j)': 'quantidade_saldo',
                        '(l = i + k)': 'valor_saldo'
                    }, inplace=True)

                    contratos = contratos[~contratos['item_codigo'].astype(str).str.contains(
                        'Total Geral:|Total do Projeto:|Exibir Apropriações dos Itens:|Parâmetros Selecionados|OCMEG',
                        na=False
                    )]

                    contratos['Contrato'] = contratos['item_codigo'].apply(extrair_contrato).ffill()
                    contratos['Empreendimento'] = contratos['item_codigo'].apply(extrair_empreendimento).ffill()
                    contratos['Fornecedor'] = contratos['descricao_item'].apply(extrair_fornecedor).ffill()

                    drop_cols = ['(g = e - f)', '(h)', '(i)', '(j)',
                                'Unnamed: 12', 'Unnamed: 13', 'Unnamed: 14', 'Unnamed: 15', 'Unnamed: 16']

                    contratos = contratos.drop(columns=[c for c in drop_cols if c in contratos.columns])
                    contratos = contratos[~contratos['item_codigo'].astype(str).str.contains('-', na=False)]

                    contratos['Fornecedor'] = contratos['Fornecedor'].astype(str)
                    agentes['Código'] = agentes['Código'].astype(str)

                    agentes['CNPJ_Limpo'] = agentes['CNPJ'].apply(limpar_documento)
                    diligencias['CNPJ_Limpo'] = diligencias['CNPJ/CPF'].apply(limpar_documento)

                    diligencias_unicas = diligencias[['CNPJ_Limpo']].drop_duplicates()

                    # Merge itens
                    contratos = pd.merge(
                        contratos,
                        itens[['Cód.Item', 'Definição Item']],
                        left_on='item_codigo',
                        right_on='Cód.Item',
                        how='left'
                    )

                    # Merge agentes
                    contratos = pd.merge(
                        contratos,
                        agentes[['Código', 'CNPJ', 'CNPJ_Limpo', 'Nome fantasia']],
                        left_on='Fornecedor',
                        right_on='Código',
                        how='left'
                    )

                    cols_valores = ['valor_unitario', 'valor_total', 'valor_saldo',
                                    'quantidade', 'quantidade_saldo']

                    for c in cols_valores:
                        contratos[c] = pd.to_numeric(contratos[c], errors='coerce').fillna(0)

                    # Aba por contrato
                    cols_group_contrato = ['Contrato', 'Empreendimento', 'Fornecedor', 'Nome fantasia', 'CNPJ']
                    aba_contratos = contratos.groupby(cols_group_contrato)[cols_valores].sum().reset_index()
                    aba_contratos.drop(columns=['valor_unitario', 'quantidade_saldo', 'quantidade'], inplace=True)

                    # Aba por fornecedor
                    cols_group_fornecedor = ['Fornecedor', 'Nome fantasia', 'CNPJ', 'CNPJ_Limpo']
                    aba_fornecedores = contratos.groupby(cols_group_fornecedor)[cols_valores].sum().reset_index()

                    filtro_servicos = contratos[contratos['Definição Item'] == 'Serviços']
                    soma_servicos = filtro_servicos.groupby('Fornecedor')['valor_total'].sum().reset_index()
                    soma_servicos.rename(columns={'valor_total': 'Total_Apenas_Servicos'}, inplace=True)

                    aba_fornecedores = pd.merge(aba_fornecedores, soma_servicos, on='Fornecedor', how='left')
                    aba_fornecedores['Total_Apenas_Servicos'] = aba_fornecedores['Total_Apenas_Servicos'].fillna(0)

                    aba_fornecedores['Verificacao_20k_Servicos'] = aba_fornecedores['Total_Apenas_Servicos'].apply(
                        lambda x: 'Atinge 20 mil' if x >= 20000 else 'Não atinge 20 mil'
                    )

                    aba_fornecedores = pd.merge(
                        aba_fornecedores,
                        diligencias_unicas,
                        on='CNPJ_Limpo',
                        how='left',
                        indicator=True
                    )

                    aba_fornecedores['Diligencias_Realizadas'] = aba_fornecedores['_merge'].apply(
                        lambda x: 'Sim' if x == 'both' else 'Não'
                    )

                    aba_fornecedores.drop(columns=['_merge', 'CNPJ_Limpo',
                                                    'valor_unitario', 'quantidade_saldo', 'quantidade'], inplace=True)

                    contratos.drop(columns=['CNPJ_Limpo'], inplace=True)

                    # Exportação
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        contratos.to_excel(writer, sheet_name='Detalhado', index=False)
                        aba_contratos.to_excel(writer, sheet_name='Por Contrato', index=False)
                        aba_fornecedores.to_excel(writer, sheet_name='Por Fornecedor', index=False)

                    output.seek(0)

                st.success("Processamento concluído!")
                st.download_button("📥 Baixar Excel Consolidado", data=output,
                                   file_name="contratos_consolidado_final.xlsx",
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

                st.write("### Prévia da Aba Por Fornecedor")
                st.dataframe(aba_fornecedores.head())

            except Exception as e:
                st.error(f"Erro: {e}")

        else:
            st.warning("Envie todos os 4 arquivos.")




with aba[1]:

    st.title("🔎 Verificação de Contratos e PEPs Incorporadora")

    st.markdown("""
    **Instruções:**
    1. Envie a planilha de relatório bruto em **.XLSX**.
    2. O sistema irá tratar os dados e remover duplicatas de CPFs válidos.
    3. Em seguida, fará a consulta automática na API do Portal da Transparência.
    """)


    arquivo_bruto = st.file_uploader("Upload Relatório Bruto (.xlsx)", type=["xlsx"])


    def processar_relatorio_xlsx(file_obj):
        """
        Lê o arquivo Excel binário, percorre as linhas para identificar
        a estrutura hierárquica e remove duplicatas mantendo inválidos.
        """
        data = []
        
        try:
            df_raw = pd.read_excel(file_obj, header=None)
        except Exception as e:
            st.error(f"Erro ao ler o arquivo Excel: {e}")
            return pd.DataFrame()

        current_emp = None
        current_contract_num = None
        current_contract_date = None

        for index, row in df_raw.iterrows():
            
            col0 = str(row[0]).strip() if pd.notna(row[0]) else ""

            if "Empreendimento:" in col0:
                if len(row) > 1 and pd.notna(row[1]):
                    current_emp = str(row[1]).strip()
                continue
            
            if col0.isdigit():
                if len(row) >= 5:
                    current_contract_num = str(row[3]).strip() if pd.notna(row[3]) else ""
                    
                    raw_date = row[4]
                    if isinstance(raw_date, pd.Timestamp):
                        current_contract_date = raw_date.strftime('%Y-%m-%d')
                    else:
                        current_contract_date = str(raw_date).strip() if pd.notna(raw_date) else ""

                    client_name = str(row[1]).strip() if pd.notna(row[1]) else ""
                    
                    data.append({
                        "Numero Contrato": current_contract_num,
                        "Empreendimento": current_emp,
                        "Nome da Parte": client_name,
                        "Tipo": "CLIENTE PRINCIPAL",
                        "CPF": None,
                        "Data Cadastro": current_contract_date
                    })
                continue

            if "Cônjuge:" in col0:
                spouse_name = str(row[1]).strip() if (len(row) > 1 and pd.notna(row[1])) else ""
                data.append({
                    "Numero Contrato": current_contract_num,
                    "Empreendimento": current_emp,
                    "Nome da Parte": spouse_name,
                    "Tipo": "CONJUGE",
                    "CPF": None,
                    "Data Cadastro": current_contract_date
                })
                continue

            if "Participante:" in col0:
                part_name = str(row[1]).strip() if (len(row) > 1 and pd.notna(row[1])) else ""
                data.append({
                    "Numero Contrato": current_contract_num,
                    "Empreendimento": current_emp,
                    "Nome da Parte": part_name,
                    "Tipo": "PARTICIPANTE",
                    "CPF": None,
                    "Data Cadastro": current_contract_date
                })
                continue
            
            if col0.startswith("CPF"):
                if data:
                    cpf_value = str(row[1]).strip() if (len(row) > 1 and pd.notna(row[1])) else ""
                    data[-1]["CPF"] = cpf_value
                continue


        df = pd.DataFrame(data)

        if df.empty:
            return df

        df = df[df['Nome da Parte'] != ""]
        termos_indesejados = [
            "Clientes no Bloco", 
            "Clientes no Empreendimento", 
            "Clientes",
            "Total Clientes",
            "Total Geral"
        ]
        df = df[~df['Nome da Parte'].astype(str).str.strip().isin(termos_indesejados)]
        df['cpf_limpo_temp'] = df['CPF'].astype(str).str.replace(r'\D', '', regex=True)
        mask_invalido = (
            (df['cpf_limpo_temp'] == '') | 
            (df['cpf_limpo_temp'].isnull()) | 
            (df['cpf_limpo_temp'] == '00000000000') |
            (df['CPF'].astype(str).str.lower().str.contains('nan'))
        )

        df_invalidos = df[mask_invalido] 
        df_validos = df[~mask_invalido]    
        df_validos = df_validos.drop_duplicates(subset=['cpf_limpo_temp'], keep='first')
        df_final = pd.concat([df_validos, df_invalidos], ignore_index=True)
        df_final = df_final.drop(columns=['cpf_limpo_temp'])

        return df_final

    API_KEY = "SUA-API-KEY-AQUI"
    API_URL = "https://api.portaldatransparencia.gov.br/api-de-dados/peps"
    HEADERS = {"chave-api-dados": API_KEY}

    def consulta_pep(cpf):
        if not cpf or str(cpf).lower() == 'nan':
            return "CPF Não Informado"
        
        cpf_limpo = re.sub(r'\D', '', str(cpf))
        

        if cpf_limpo == '00000000000':
            return "CPF Zerado"

        if len(cpf_limpo) != 11:
             return "CPF Inválido/Incompleto"

        try:
            r = requests.get(API_URL, headers=HEADERS,
                             params={'cpf': cpf_limpo, 'pagina': 1}, timeout=10)

            if r.status_code == 200:
                return "Sim" if r.json() else "Não"
            if r.status_code == 404:
                return "Não" 
            if r.status_code == 429:
                return "Erro 429 (Muitas requisições)"
            if r.status_code in [401, 403]:
                return "Erro API (Auth)"
            return f"Erro {r.status_code}"

        except:
            return "Erro Conexão"


    if st.button("Processar e Consultar PEPs"):

        if arquivo_bruto is None:
            st.warning("Envie o arquivo XLSX primeiro.")
            st.stop()

        try:
            st.info("1/2 - Tratando planilha XLSX e removendo duplicatas...")
            
            df_clientes = processar_relatorio_xlsx(arquivo_bruto)
            
            if df_clientes.empty:
                st.error("Não foi possível extrair dados ou a planilha estava vazia.")
                st.stop()
            
            qtd_total = len(df_clientes)
            st.success(f"Base tratada: {qtd_total} registros para análise (Duplicatas válidas removidas).")
            st.write("Prévia dos dados:", df_clientes.head())

            st.info("2/2 - Consultando API do Governo...")
            
            lista_cpfs = df_clientes["CPF"].astype(str).tolist()

            with concurrent.futures.ThreadPoolExecutor(max_workers=20) as ex:
                resultados = list(ex.map(consulta_pep, lista_cpfs))

            df_clientes["E_PEP"] = resultados

            st.success("Consulta PEP concluída!")
            st.dataframe(df_clientes)

            output_pep = io.BytesIO()
            with pd.ExcelWriter(output_pep, engine='openpyxl') as writer:
                df_clientes.to_excel(writer, index=False, sheet_name="Resultado_PEP")
            
            output_pep.seek(0)

            st.download_button(
                "📥 Baixar Resultado (Tratado + PEP)",
                data=output_pep,
                file_name="clientes_verificacao_pep.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"Erro crítico: {e}")
