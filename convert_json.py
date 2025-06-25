import pandas as pd
import json
import os

def converter_todos_json_para_excel(pasta_json, pasta_saida):
    # Cria pasta de saída, se não existir
    os.makedirs(pasta_saida, exist_ok=True)

    # Percorre todos os arquivos .json na pasta
    for arquivo in os.listdir(pasta_json):
        if arquivo.lower().endswith('.json'):
            caminho_json = os.path.join(pasta_json, arquivo)
            nome_base = os.path.splitext(arquivo)[0]
            caminho_excel = os.path.join(pasta_saida, f"{nome_base}.xlsx")

            try:
                with open(caminho_json, 'r', encoding='utf-8') as f:
                    dados = json.load(f)

                # Tenta detectar uma lista dentro do JSON (estrutura comum)
                if isinstance(dados, dict):
                    for valor in dados.values():
                        if isinstance(valor, list):
                            dados = valor
                            break

                # Converte para DataFrame e salva
                df = pd.DataFrame(dados)
                df.to_excel(caminho_excel, index=False)
                print(f"[✓] Convertido: {arquivo} → {nome_base}.xlsx")

            except Exception as e:
                print(f"[x] Erro no arquivo {arquivo}: {e}")

# Exemplo de uso
pasta_entrada = 'C:/Users/wagner.felipen.ext/Downloads/Transactions/Set_2024/JSON'       # ex: r'C:\MeusArquivos\Jsons'
pasta_saida = 'C:/Users/wagner.felipen.ext/Downloads/Transactions/Set_2024/Excel'          # ex: r'C:\MeusArquivos\Excels'

converter_todos_json_para_excel(pasta_entrada, pasta_saida)
