########### explicação do código ###########
# Esse código é uma função que lê um arquivo Excel, adiciona uma coluna "Rota" e remove os acentos das cidades de destino. 
# Ele pergunta ao usuário se deseja sobrescrever o arquivo original ou criar um novo com sufixo "_atualizado". 
# Se o arquivo não for encontrado, uma mensagem de erro é exibida. 
# Se houver um erro de valor, uma mensagem de erro específica é exibida. 
# Se ocorrer qualquer outro erro, uma mensagem de erro genérica é exibida.


import pandas as pd
import os

# Tenta importar unidecode, se não estiver disponível, usa uma função alternativa
try:
    from unidecode import unidecode
except ImportError:
    print("A biblioteca unidecode não está instalada. Usando uma função alternativa para remover acentos.")
    import unicodedata

    def unidecode(texto):
        return ''.join(c for c in unicodedata.normalize('NFD', texto)
                       if unicodedata.category(c) != 'Mn')

# Caminho do arquivo
caminho_arquivo = r"C:\Users\Petry\Documents\Py\rota\Relatorio MDFe (5).xlsx"

try:
    # Lê o arquivo Excel, ignorando a primeira e a última linha
    df = pd.read_excel(caminho_arquivo, header=1, skipfooter=1, engine='openpyxl')

    # Adiciona a coluna "Rota" ao lado de "Cidade de destino"
    if "Cidade de destino" in df.columns:
        df.insert(df.columns.get_loc("Cidade de destino") + 1, "Rota", "")
    else:
        raise ValueError("Coluna 'Cidade de destino' não encontrada no arquivo.")

    # Converte a coluna "Cidade de destino" para maiúsculas e remove acentos
    df["Cidade de destino"] = df["Cidade de destino"].apply(lambda x: unidecode(str(x).upper()))

    # Pergunta ao usuário se deseja sobrescrever o arquivo original ou criar um novo
    sobrescrever = input("Deseja sobrescrever o arquivo original? (S/N): ").strip().lower()

    if sobrescrever == 's':
        # Sobrescreve o arquivo original
        df.to_excel(caminho_arquivo, index=False)
        print(f"Arquivo original atualizado: {caminho_arquivo}")
    else:
        # Cria um novo arquivo com sufixo "_atualizado"
        novo_caminho = os.path.splitext(caminho_arquivo)[0] + "_atualizado.xlsx"
        df.to_excel(novo_caminho, index=False)
        print(f"Novo arquivo criado: {novo_caminho}")

    print("Arquivo processado e salvo com sucesso!")

except FileNotFoundError:
    print("Arquivo não encontrado. Verifique se o caminho está correto.")
except ValueError as ve:
    print(f"Erro de valor: {ve}")
except Exception as e:
    print(f"Ocorreu um erro ao processar o arquivo: {e}")

########### explicação do código ###########
# Esse código é uma função que lê um arquivo Excel, adiciona uma coluna "Rota" e remove os acentos das cidades de destino. 
# Ele pergunta ao usuário se deseja sobrescrever o arquivo original ou criar um novo com sufixo "_atualizado". 
# Se o arquivo não for encontrado, uma mensagem de erro é exibida. 
# Se houver um erro de valor, uma mensagem de erro específica é exibida. 
# Se ocorrer qualquer outro erro, uma mensagem de erro genérica é exibida.

