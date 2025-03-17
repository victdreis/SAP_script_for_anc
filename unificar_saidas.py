import os
import pandas as pd

# Diretórios
pasta_input = "output"
pasta_output = "deliver"

# Criar pasta de saída se não existir
os.makedirs(pasta_output, exist_ok=True)

# Listar arquivos Excel na pasta
arquivos = [os.path.join(pasta_input, f) for f in os.listdir(pasta_input) if f.endswith(".xlsx")]

# Lista para armazenar os DataFrames
dataframes = []

# Iterar sobre cada arquivo e carregar no Pandas
for arquivo in arquivos:
    try:
        df = pd.read_excel(arquivo, engine="openpyxl")
        dataframes.append(df)  # Adicionar ao conjunto de dados
    except Exception as e:
        print(f"Erro ao processar {arquivo}: {e}")

# Se encontrou arquivos válidos, concatena e salva
if dataframes:
    df_final = pd.concat(dataframes, ignore_index=True)  # Concatenar todos os DataFrames
    caminho_saida = os.path.join(pasta_output, "historico_de_pedidos_CONSOLIDADO.csv")
    df_final.to_csv(caminho_saida, index=False, sep=";", encoding="utf-8")  # Salvar como CSV
    print(f"Arquivo único salvo com sucesso: {caminho_saida}")
else:
    print("Nenhum arquivo válido encontrado para processar.")
