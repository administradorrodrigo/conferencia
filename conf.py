import pandas as pd
import os
import openpyxl
import fitz  # PyMuPDF
import pdfplumber
import re


# --- CONFIGURAÇÃO ---
# Define os caminhos dos arquivos
caminho_planilha = r"C:\\Users\\89524713\\OneDrive - correiosbrasil\\Documentos\\conferencia-main\\Conferência.xlsx"
caminho_do_pdf = r"C:\\Users\\89524713\Documents\\conferencia\\holerite\\HOLERITE.pdf"

# Define os padrões de busca para cada rubrica
padroes_rubricas = {
    'funcao': r"Função[\s\S]*?Referência\s*([^\n]+)",
    'total de proventos': r"Total de Proventos[\s\S]*?([\d\.,]+)",
    'fgts do mês': r"FGTS do Mês[\s\S]*?([\d\.,]+)",
    'desconto transporte': r"Desconto Transporte[\s\S]*?([\d\.,]+)",
    'vale refeição': r"Vale Refeição[\s\S]*?([\d\.,]+)"
}
# Dicionário para mapear as rubricas para os nomes das colunas na planilha
mapeamento_colunas = {
    'funcao': 'CARGO',
    'total de proventos': 'SALARIO',
    'fgts do mês': 'FGTS',
    'desconto transporte': 'VT',
    'vale refeição': 'VA'
}

# --- EXTRAÇÃO E PROCESSAMENTO ---
# 1. Lê a planilha e prepara o DataFrame
df = pd.read_excel('Conferência.xlsx')
# Renomeia as colunas para o mapeamento
df.rename(columns={'NOME': 'nome', 'CARGO': 'cargo', 'SALARIO': 'salario', 'FGTS': 'fgts', 'VT': 'vt', 'VA': 'va'}, inplace=True)
nomes_a_procurar = df['nome'].tolist()

print("Planilha lida e colunas padronizadas.")

# 2. Extrai todo o texto do PDF
texto_completo_pdf = ""
try:
    with pdfplumber.open(caminho_do_pdf) as pdf:
        for pagina in pdf.pages:
            texto_completo_pdf += pagina.extract_text()
    print("Texto do PDF extraído com sucesso.")
except FileNotFoundError:
    print(f"Erro: Arquivo PDF não encontrado no caminho: {caminho_do_pdf}")
    exit() # Encerra o script se o arquivo não for encontrado

# 3. Itera sobre cada nome e procura as rubricas no PDF
for index, nome in enumerate(nomes_a_procurar):
    print(f"\nBuscando dados para: {nome}")
    
    # Encontra todas as ocorrências do nome no PDF
    # Isso é útil se o mesmo nome aparecer em mais de uma página
    posicoes_nome = [m.start() for m in re.finditer(re.escape(nome), texto_completo_pdf)]
    
    if posicoes_nome:
        posicao_nome = posicoes_nome[0] # Pega a primeira ocorrência
        
        # Cria uma "seção de texto" a partir da posição do nome até o próximo nome
        if index + 1 < len(nomes_a_procurar):
            proximo_nome = nomes_a_procurar[index + 1]
            try:
                # Delimita o texto entre o nome atual e o próximo nome
                limite_secao = texto_completo_pdf.find(proximo_nome, posicao_nome)
                if limite_secao != -1:
                    secao_texto = texto_completo_pdf[posicao_nome:limite_secao]
                else:
                    secao_texto = texto_completo_pdf[posicao_nome:]
            except ValueError:
                secao_texto = texto_completo_pdf[posicao_nome:]
        else:
            secao_texto = texto_completo_pdf[posicao_nome:]
        
        # Dicionário para os dados deste funcionário
        dados_funcionario = {}
        
        # Extrai cada rubrica dentro da seção de texto
        for rubrica, padrao_regex in padroes_rubricas.items():
            match = re.search(padrao_regex, secao_texto, re.IGNORECASE) # Usamos IGNORECASE para mais flexibilidade
            
            if match:
                valor_bruto = match.group(1).strip()
                if rubrica in ['total de proventos', 'fgts do mês', 'desconto transporte', 'vale refeição']:
                    valor_limpo = valor_bruto.replace('.', '').replace(',', '.')
                    dados_funcionario[rubrica] = float(valor_limpo)
                else:
                    dados_funcionario[rubrica] = valor_bruto
            else:
                dados_funcionario[rubrica] = None
                print(f"Aviso: Rubrica '{rubrica}' não encontrada para {nome}.")
        
        # Atualiza a linha do DataFrame com os dados extraídos
        for rubrica, valor in dados_funcionario.items():
            coluna_df = mapeamento_colunas.get(rubrica)
            if coluna_df:
                df.loc[df['nome'] == nome, coluna_df] = valor
    else:
        print(f"Aviso: Nome '{nome}' não encontrado no PDF.")

print("\nExtração e atualização finalizadas. Primeiras linhas do DataFrame atualizado:")
print(df.head(100))

# Opcional: Salvar o DataFrame atualizado em um novo arquivo Excel
# df.to_excel("Conferência_Atualizada.xlsx", index=False)