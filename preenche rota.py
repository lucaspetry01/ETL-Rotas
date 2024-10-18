import pandas as pd
import os
from unidecode import unidecode

def remover_acentos(texto):
    return unidecode(str(texto).upper())

caminho_arquivo_rota = r"C:\Users\Petry\Documents\Py\rota\rota_atualizada.xlsx"
caminho_arquivo_mdfe = r"C:\Users\Petry\Documents\Py\Relatorio MDFe (5).xlsx"

try:
    df_rota = pd.read_excel(caminho_arquivo_rota)
    df_mdfe = pd.read_excel(caminho_arquivo_mdfe, header=1, skipfooter=1, engine='openpyxl')

    if "Cidade de destino" not in df_mdfe.columns:
        raise ValueError("Coluna 'Cidade de destino' não encontrada no arquivo MDFe.")
    
    if "Rota" not in df_mdfe.columns:
        df_mdfe.insert(df_mdfe.columns.get_loc("Cidade de destino") + 1, "Rota", "")

    colunas_necessarias = ["Rota", "Cidade de destino"]
    for df, nome_arquivo in [(df_rota, "rota_atualizada.xlsx"), (df_mdfe, "Relatorio MDFe (5).xlsx")]:
        colunas_faltantes = [col for col in colunas_necessarias if col not in df.columns]
        if colunas_faltantes:
            raise ValueError(f"Colunas não encontradas no arquivo {nome_arquivo}: {', '.join(colunas_faltantes)}")

    for df in [df_rota, df_mdfe]:
        for coluna in colunas_necessarias:
            df[coluna] = df[coluna].apply(remover_acentos)

    mapeamento_rotas = dict(zip(df_rota["Cidade de destino"], df_rota["Rota"]))
    df_mdfe["Rota"] = df_mdfe["Cidade de destino"].map(mapeamento_rotas)

    # Corrigir os valores NaN na coluna Rota
    df_mdfe["Rota"] = df_mdfe["Rota"].fillna(df_mdfe["Cidade de destino"])

    # Validação para verificar se o campo Rota foi preenchido
    rotas_nao_preenchidas = df_mdfe[df_mdfe["Rota"] == ""]
    if not rotas_nao_preenchidas.empty:
        print("Atenção: Algumas rotas não foram preenchidas:")
        print(rotas_nao_preenchidas[["Cidade de destino", "Rota"]])
        if input("Deseja continuar mesmo com rotas não preenchidas? (S/N): ").strip().upper() != 'S':
            raise ValueError("Processamento interrompido devido a rotas não preenchidas.")

    print("Primeiras linhas após o preenchimento da Rota:")
    print(df_mdfe[colunas_necessarias].head())

    if input("Deseja gerar o Relatorio MDFe (51)? (S/N): ").strip().upper() == 'S':
        novo_caminho = r"C:\Users\Petry\Documents\Py\rota\Relatorio MDFe (51).xlsx"
        df_mdfe.to_excel(novo_caminho, index=False)
        print(f"Novo relatório gerado: {novo_caminho}")
    elif input("Deseja sobrescrever o arquivo original? (S/N): ").strip().upper() == 'S':
        df_mdfe.to_excel(caminho_arquivo_mdfe, index=False)
        print(f"Arquivo MDFe original atualizado: {caminho_arquivo_mdfe}")
    else:
        novo_caminho = os.path.splitext(caminho_arquivo_mdfe)[0] + "_atualizado.xlsx"
        df_mdfe.to_excel(novo_caminho, index=False)
        print(f"Novo arquivo criado: {novo_caminho}")

    print("Processamento concluído com sucesso!")

except FileNotFoundError as e:
    print(f"Arquivo não encontrado: {e}")
except ValueError as ve:
    print(f"Erro de valor: {ve}")
except Exception as e:
    print(f"Ocorreu um erro ao processar os arquivos: {e}")