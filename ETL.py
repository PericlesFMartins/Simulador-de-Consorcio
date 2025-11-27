import os
import pdfplumber
import pandas as pd
import numpy as np

def extrair_dados_de_pdfs(pasta_pdfs, excel_path):
    """
    Localiza os .PDF que estão dentro da pasta PDF, extrai as tabelas de todos os arquivos e salva em uma planilha Excel.
    armazeno no .xlsx o nome do arquivo .PDF que aquela informação pertence.
    """
    print("Iniciando extração de dados dos PDFs...")
    dados_extraidos = []
    
    # Processa cada arquivo PDF na pasta
    for arquivo in os.listdir(pasta_pdfs):
        if arquivo.endswith(".pdf"):
            pdf_path = os.path.join(pasta_pdfs, arquivo)
            print(f"Processando {arquivo}...")

            with pdfplumber.open(pdf_path) as pdf:
                for page in pdf.pages:
                    tables = page.extract_tables()
                    for table in tables:
                        if table:
                            for row in table:
                                # Adiciona o nome do arquivo como referência na primeira coluna
                                dados_extraidos.append([arquivo] + row)

    if not dados_extraidos:
        print("Erro ao extrair dados dos PDF.")
        return

    # Cria um DataFrame e salva em um arquivo Excel, substituindo o anterior
    df = pd.DataFrame(dados_extraidos)
    df.to_excel(excel_path, index=False, header=False, sheet_name='banco')
    print("Extração concluída. Partindo para tratamento dos dados...", excel_path)

def tratar_dados_excel(excel_path):
    """
    Lê os dados brutos do Excel, limpa, transforma e salva o resultado final.
    """
    print("Iniciando tratamento dos dados no Excel...")
    df = pd.read_excel(excel_path, sheet_name='banco', header=None)

    # Preenche os valores de grupo ausentes usando o último valor válido
    df[1] = df[1].ffill()

    # Desloca colunas para linhas onde a coluna 3 está vazia
    filtro_col3_vazia = df[3].isnull()
    colunas_para_deslocar = [3, 4, 5, 6, 7, 8, 9]
    df.loc[filtro_col3_vazia, colunas_para_deslocar] = df.loc[filtro_col3_vazia, [4, 5, 6, 7, 8, 9, 9]].values

    # Remove a última coluna que se tornou redundante
    df = df.drop(columns=[9])

    # Define o cabeçalho a partir da primeira linha e remove quebras de linha
    novo_cabecalho = df.iloc[1].str.replace('\n', ' ', regex=False).str.strip()
    df_novo = df.iloc[1:].copy()
    df_novo.columns = novo_cabecalho
    df_novo.columns.name = None # Limpa o nome do índice das colunas

    # Limpeza e tratamento das colunas
    df_novo['Grupo'] = df_novo['Grupo'].str.replace('Caracteristicas.*|Grupo|Plano -', '', regex=True).str.lstrip('0')
    df_novo['Bem'] = df_novo['Bem'].str.replace('Bem', '', regex=False)
    df_novo['Valor do Bem'] = df_novo['Valor do Bem'].str.replace('Parcelas com seguro|Valor do Bem', '', regex=True)

    # Limpa valores indesejados em colunas de prestações
    colunas_prestacoes = df_novo.columns[4:9]
    valores_a_remover = [
        "1ª a 6ª Prestações", "Demais Prestações", 
        "Parcelas sem seguro", "Parcelas com seguro", "Situação do Grupo"
    ]
    for col in colunas_prestacoes:
        df_novo[col] = df_novo[col].replace(valores_a_remover, "")

    # Preenche a situação do grupo
    df_novo["Situação do Grupo"] = df_novo["Situação do Grupo"].ffill()

    # Move informações de "Taxa de" da coluna 'Bem' para 'Valor do Bem'
    filtro_taxa = df_novo['Bem'].str.contains("Taxa de", na=False)
    df_novo.loc[filtro_taxa, 'Valor do Bem'] = df_novo.loc[filtro_taxa, 'Bem']
    df_novo.loc[filtro_taxa, 'Bem'] = np.nan

    # Substitui strings vazias por NaN para facilitar a remoção de linhas
    df_novo['Bem'] = df_novo['Bem'].replace('', np.nan)
    df_novo['Valor do Bem'] = df_novo['Valor do Bem'].replace('', np.nan)

    # Remove linhas onde tanto 'Bem' quanto 'Valor do Bem' são nulos
    df_novo = df_novo.dropna(subset=["Bem", "Valor do Bem"], how="all")

    # Limpa a coluna 'Bem' que contém "RECIPROCIDADE"
    df_novo.loc[df_novo['Bem'].str.contains("RECIPROCIDADE", na=False), 'Bem'] = np.nan

    # Cria a coluna "Dados do grupo" e move as informações de taxa para ela
    df_novo["Dados do grupo"] = np.nan
    filtro_taxa_valor = df_novo['Valor do Bem'].str.contains("Taxa de", na=False)
    df_novo.loc[filtro_taxa_valor, 'Dados do grupo'] = df_novo.loc[filtro_taxa_valor, 'Valor do Bem']

    # Preenche para cima as informações de taxa na nova coluna
    df_novo["Dados do grupo"] = df_novo["Dados do grupo"].bfill()

    # Remove linhas que não são de dados principais (ex: linhas de taxa)
    df_novo = df_novo.dropna(subset=["Bem", "Demais Prestações"], how="all")

    # Salva o DataFrame limpo no mesmo arquivo Excel
    df_novo.to_excel(excel_path, index=False, sheet_name='banco')
    print("Tratamento de dados concluído. Arquivo final salvo.")

if __name__ == "__main__":
    try:
        # Retorna o nome do usuário para construir os caminhos
        usuario = os.getlogin()

        # Define os caminhos da pasta de PDFs e do arquivo Excel de saída
        base_path = fr"C:\Users\{usuario}\OneDrive\Documents\Simulador-de-Consorcio"
        pasta_pdfs = os.path.join(base_path, "PDF")
        excel_path = os.path.join(base_path, "PDF", "XLSX", "tabelas_banco.xlsx")

        # Garante que o diretório de saída do Excel exista
        os.makedirs(os.path.dirname(excel_path), exist_ok=True)

        # Etapa 1: Extrair dados dos PDFs para o Excel
        extrair_dados_de_pdfs(pasta_pdfs, excel_path)

        # Etapa 2: Tratar e limpar os dados no arquivo Excel
        tratar_dados_excel(excel_path)

        print("\nProcesso finalizado com sucesso!")
    except FileNotFoundError:
        print(f"Erro: O diretório '{pasta_pdfs}' não foi encontrado. Verifique o caminho.")
    except Exception as e:
        print(f"Ocorreu um erro inesperado durante a execução: {e}")
