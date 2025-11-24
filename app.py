import streamlit as st
import pandas as pd
import re
import io

# Configuração da Página
st.set_page_config(page_title="Portal de Compliance", layout="wide")

st.title("📊 Controle Interno - Fornecedores sem Diligência")
st.markdown("""
Faça o upload das quatro planilhas abaixo.
""")

# --- SEÇÃO DE UPLOAD ---
col1, col2, col3, col4 = st.columns(4)
with col1:
    file_contratos = st.file_uploader("Upload Contratos (.xlsx)", type="xlsx", help="Planilha com detalhes dos contratos por item")
with col2:
    file_itens = st.file_uploader("Upload Itens (.xlsx)", type="xlsx", help="Planilha com detalhes dos itens cadastrados")
with col3:
    file_agentes = st.file_uploader("Upload Agentes (.xlsx)", type="xlsx", help="Planilha com CNPJ e Código dos fornecedores")
with col4:
    file_diligencias = st.file_uploader("Upload Diligências (.xlsx)", type="xlsx", help="Planilha com CNPJ/CPF das diligências realizadas")

# --- FUNÇÕES AUXILIARES ---
def extrair_contrato(texto):
    if isinstance(texto, str):
        resultado = re.search(r'Contrato:\s*(\d+)\s*\(', texto)
        if resultado:
            return resultado.group(1)
    return None

def extrair_empreendimento(texto):
    if isinstance(texto, str):
        resultado = re.search(r'^\d{11} - ', texto)
        if resultado:
            return texto.split(' - ')[1]
    return None

def extrair_fornecedor(texto):
    if isinstance(texto, str):
        resultado = re.search(r'Fornecedor:\s*(\d+)\s*-', texto)
        if resultado:
            return resultado.group(1)
    return None

# Limpeza de CNPJ para garantir o Merge
def limpar_documento(valor):
    if pd.isna(valor):
        return ""
    # Converte para string e remove tudo que NÃO for número (pontos, traços, barras, espaços)
    return re.sub(r'\D', '', str(valor))

# --- BOTÃO DE PROCESSAMENTO ---
if st.button("Processar Arquivos"):
    # Verifica se os 4 arquivos foram enviados (adicionei file_diligencias na checagem)
    if file_contratos and file_itens and file_agentes and file_diligencias:
        try:
            with st.spinner('Lendo, limpando CNPJs e processando...'):
                # 1. Carregamento
                contratos = pd.read_excel(file_contratos)
                itens = pd.read_excel(file_itens) # Ajuste sheet_name se necessário
                agentes = pd.read_excel(file_agentes)
                diligencias = pd.read_excel(file_diligencias, sheet_name='DDR e BCK') 

                # 2. Tratamento Inicial Contratos
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
                    na=False)]

                contratos['Contrato'] = contratos['item_codigo'].apply(extrair_contrato).ffill()
                contratos['Empreendimento'] = contratos['item_codigo'].apply(extrair_empreendimento).ffill()
                contratos['Fornecedor'] = contratos['descricao_item'].apply(extrair_fornecedor).ffill()

                cols_to_drop = ['(g = e - f)', '(h)', '(i)', '(j)', 'Unnamed: 12',
                                'Unnamed: 13', 'Unnamed: 14', 'Unnamed: 15', 'Unnamed: 16']
                contratos.drop(columns=[c for c in cols_to_drop if c in contratos.columns], inplace=True)
                contratos = contratos[~contratos['item_codigo'].astype(str).str.contains('-', na=False)]

                # 3. Tipagem e LIMPEZA DE CHAVES
                contratos['Fornecedor'] = contratos['Fornecedor'].astype(str)
                agentes['Código'] = agentes['Código'].astype(str)
                

                agentes['CNPJ_Limpo'] = agentes['CNPJ'].apply(limpar_documento)
                diligencias['CNPJ_Limpo'] = diligencias['CNPJ/CPF'].apply(limpar_documento)
                
                # Removemos duplicatas da diligencia para não duplicar linhas no merge
                # Mantemos apenas uma ocorrencia de cada CNPJ limpo para saber se existe ou não
                diligencias_unicas = diligencias[['CNPJ_Limpo']].drop_duplicates()
                # --------------------------------------

                # 4. Merges
                # Itens
                contratos = pd.merge(
                    contratos,
                    itens[['Cód.Item','Definição Item']],
                    left_on='item_codigo',
                    right_on='Cód.Item',
                    how='left'
                )

                # Agentes (Trazendo CNPJ Limpo e Nome)
                contratos = pd.merge(
                    contratos,
                    agentes[['Código', 'CNPJ', 'CNPJ_Limpo', 'Nome fantasia']], # Trazemos o CNPJ original e o Limpo
                    left_on='Fornecedor',
                    right_on='Código',
                    how='left'
                )

                # 5. Conversão Numérica
                cols_valores = ['valor_unitario', 'valor_total', 'valor_saldo', 'quantidade', 'quantidade_saldo']
                for col in cols_valores:
                    contratos[col] = pd.to_numeric(contratos[col], errors='coerce').fillna(0)

                # ---------------------------------------------------------
                # ABA 2: Agrupada por Contrato
                # ---------------------------------------------------------
                cols_group_contrato = ['Contrato', 'Empreendimento', 'Fornecedor', 'Nome fantasia', 'CNPJ']
                aba_contratos = contratos.groupby(cols_group_contrato)[cols_valores].sum().reset_index()
                # Excluindo colunas valor_unitario, quantidade_saldo e quantidade
                aba_contratos.drop(columns=['valor_unitario', 'quantidade_saldo', 'quantidade'], inplace=True)



                # ---------------------------------------------------------
                # ABA 3: Agrupada por Fornecedor
                # ---------------------------------------------------------
                # Agrupamos usando também o CNPJ_Limpo para usar no próximo merge
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
                
                # MERGE COM DILIGÊNCIAS (Usando as colunas LIMPAS)
                aba_fornecedores = pd.merge(
                    aba_fornecedores,
                    diligencias_unicas, # Usamos a base de diligencias sem duplicatas
                    on='CNPJ_Limpo',    # Chave limpa (só números)
                    how='left',
                    indicator=True
                )
                
                aba_fornecedores['Diligencias_Realizadas'] = aba_fornecedores['_merge'].apply(
                    lambda x: 'Sim' if x == 'both' else 'Não'
                )
                
                # Limpeza final (remove colunas auxiliares)
                aba_fornecedores.drop(columns=['_merge', 'CNPJ_Limpo', 'valor_unitario', 'quantidade_saldo', 'quantidade'], inplace=True)



                # Opcional: remover CNPJ_Limpo da detalhada se não quiser ver
                contratos.drop(columns=['CNPJ_Limpo'], inplace=True) 

                

                # 6. Exportação
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    contratos.to_excel(writer, sheet_name='Detalhado', index=False)
                    aba_contratos.to_excel(writer, sheet_name='Por Contrato', index=False)
                    aba_fornecedores.to_excel(writer, sheet_name='Por Fornecedor', index=False)
                
                output.seek(0)

            st.success("Processamento concluído com sucesso!")
            
            st.download_button(
                label="📥 Baixar Excel Consolidado",
                data=output,
                file_name="contratos_consolidado_final.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
            st.write("### Prévia da Aba Por Fornecedor")
            st.dataframe(aba_fornecedores.head())

        except Exception as e:
            st.error(f"Ocorreu um erro: {e}")
            st.write("Dica: Verifique se os nomes das colunas nos arquivos correspondem ao esperado no código.")
    else:
        st.warning("Por favor, faça o upload de todos os 4 arquivos.")
