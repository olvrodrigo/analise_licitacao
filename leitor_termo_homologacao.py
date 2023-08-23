import os
import re
import PyPDF2
import openpyxl

# Caminho da pasta com os arquivos PDF
pdf_folder = r"U:\#03 - CONTRATACAO\#01 - PLANEJAMENTO\PROT 2023-4086 - OUTSOURCING DE IMPRESSÃO\3.0 - COTACAO_PRECO\Analise Comprasnet"

# Termo a ser pesquisado no nome dos arquivos PDF
search_term = "homologa"

# Lista para armazenar os dados
data_list = []

# Função para extrair valor entre duas substrings
def extract_value(text, start_str, end_str):
    start_index = text.find(start_str) + len(start_str)
    end_index = text.find(end_str, start_index)
    return text[start_index:end_index].strip()

# Iterar sobre os arquivos PDF na pasta e suas subpastas
for root, dirs, files in os.walk(pdf_folder):
    for file in files:
        if file.lower().endswith('.pdf') and search_term in file.lower():
            pdf_path = os.path.join(root, file)

            with open(pdf_path, 'rb') as pdf_file:
                pdf_reader = PyPDF2.PdfReader(pdf_file)
                pdf_text = ' '.join(page.extract_text() for page in pdf_reader.pages)

                pregao_numero = extract_value(pdf_text, "Pregão Nº", "Item:")
                item_numero = extract_value(pdf_text, "Item:", "Modalidade")
                modalidade = extract_value(pdf_text, "Modalidade", "Código do CATMAT/CATSER")
                codigo_catmat = extract_value(pdf_text, "Código do CATMAT/CATSER", "Descrição:")
                descricao = extract_value(pdf_text, "Descrição:", "Unidade de fornecimento")
                unidade_fornecimento = extract_value(pdf_text, "Unidade de fornecimento", "Quantidade:")
                quantidade = extract_value(pdf_text, "Quantidade:", "Valor Máximo Aceitável:")
                valor_maximo = extract_value(pdf_text, "Valor Máximo Aceitável:", "melhor lance de")
                melhor_lance = extract_value(pdf_text, "melhor lance de", "com valor negociado a")
                valor_unitario = extract_value(pdf_text, "com valor negociado a", "Mediana")
                fornecedor = extract_value(pdf_text, "Adjudicação individual da proposta. Fornecedor:", " do dia ")
                orgao = extract_value(pdf_text, "Pregão/Concorrência Eletrônica", "Termo de Homologação do Pregão Eletrônico")
                data_compra = extract_value(pdf_text, " do dia ", "")

                data_list.append([pregao_numero, item_numero, modalidade, codigo_catmat, descricao, unidade_fornecimento, quantidade,
                                  valor_maximo, melhor_lance, valor_unitario, "", fornecedor, orgao, "", data_compra])

# Criar um arquivo Excel e escrever os dados
output_excel = "output_data.xlsx"
wb = openpyxl.Workbook()
ws = wb.active

# Escrever cabeçalhos
headers = ["Identificação da Compra", "Número do Item", "Modalidade", "Código do CATMAT/CATSER", "Item", "Unidade de Fornecimento",
           "Quantidade Ofertada", "Valor Máximo Aceitável:", "melhor lance de", "Valor Unitário", "Mediana", "Fornecedor", "Órgão",
           "UASG - Unidade Gestora", "Data da Compra"]
ws.append(headers)

# Escrever os dados
for row_data in data_list:
    ws.append(row_data)

# Salvar o arquivo Excel
wb.save(output_excel)
print(f"Dados salvos em {output_excel}")
