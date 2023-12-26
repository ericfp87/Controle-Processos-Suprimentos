# -*- coding: utf-8 -*-

import pandas as pd

# Nome do arquivo Excel
arquivo_excel = "G:\\Drives compartilhados\\Analise de Dados\\02 - Projetos\\01 - BI\Controle Processos Suprimentos\\CONTROLE PROCESSOS - SUPRIMENTOS - NF.xlsm"

# Nome do arquivo CSV de saída
arquivo_csv = "G:\\Drives compartilhados\\Analise de Dados\\02 - Projetos\\01 - BI\Controle Processos Suprimentos\\processos_suprimentos.csv"

# Nomes dos meses do ano
# nomes_meses = ["JANEIRO", "FEVEREIRO", "MARÇO", "ABRIL", "MAIO", "JUNHO", "JULHO", "AGOSTO", "SETEMBRO", "OUTUBRO", "NOVEMBRO", "DEZEMBRO"]
nomes_meses = ["JANEIRO", "FEVEREIRO"]

# Nome das colunas
nomes_colunas = [
    "Obra/Projeto", "Fornecedor", "Solicitação", "DESCRIÇÃO", "O.C", "Data O.C", "TIPO OC.", "Status processo", "Data recebimento NF",
    "Status NF", "Nota Fiscal", "Data emissão NF", "Data venc. NF", "Adiantamento Valor (R$)", "Data adiantamento",
    "Valor Total NF", "Valor devido NF (R$)", "Condição de Pgto", "Data lançamento Aviso", "Tempo lançamento",
    "Lançar Aviso Recbimento", "Data entrega material", "Status Entrega Material", "Registros", "Pago", "Colaborador_str", "Colaborador"
]

# Lista para armazenar os dados de todas as planilhas
dados_total = []

# Loop para processar cada planilha do Excel
for mes in nomes_meses:
    # Ler a planilha do Excel
    df = pd.read_excel(arquivo_excel, sheet_name=mes, skiprows=2, usecols="A:Z", header=None, names=nomes_colunas)
    # Adicionar os dados da planilha à lista total
    dados_total.append(df)



# Concatenar os dados de todas as planilhas
dados_concatenados = pd.concat(dados_total)

# Salvar os dados concatenados em um arquivo CSV
dados_concatenados.to_csv(arquivo_csv, sep=";", index=False, encoding="utf-8-sig", decimal=',')