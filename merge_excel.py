import pandas as pd
import os

def merge_excels_em_csv(pasta_excel, arquivo_saida_csv):
    arquivos = [f for f in os.listdir(pasta_excel) if f.lower().endswith('.xlsx')]
    todos_df = []

    for arquivo in arquivos:
        caminho = os.path.join(pasta_excel, arquivo)
        try:
            df = pd.read_excel(caminho)
            df['Arquivo_Origem'] = arquivo  # opcional
            todos_df.append(df)
            print(f"[✓] Lido: {arquivo}")
        except Exception as e:
            print(f"[x] Erro ao ler {arquivo}: {e}")

    if todos_df:
        resultado = pd.concat(todos_df, ignore_index=True)
        resultado.to_csv(arquivo_saida_csv, index=False, encoding='utf-8-sig')
        print(f"\n[✓] Arquivo final salvo em: {arquivo_saida_csv}")
    else:
        print("Nenhum arquivo Excel válido encontrado para mesclar.")

# Exemplo de uso
pasta_entrada = r'C:\Users\wagner.felipen.ext\Downloads\Transactions\Set_2024\Excel\16a31'
arquivo_saida = r'C:\Users\wagner.felipen.ext\Downloads\Transactions\Set_2024\Excel\16a31\merged.csv'

merge_excels_em_csv(pasta_entrada, arquivo_saida)
