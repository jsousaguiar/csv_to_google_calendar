import pandas as pd
import os

try:
    df = pd.read_excel("modelo_calendario.xlsx")
except FileNotFoundError:
    print("Arquivo modelo não encontrado.")
    exit()
if os.path.exists("importar.csv"):
    os.remove("importar.csv")

df.to_csv("importar.csv", sep=",", encoding="latin-1", index=False)
if os.path.exists("importar.csv"):
    print("Arquivo 'importar.csv' gerado com sucesso!")
else:
    print("Não foi possível gerar o arquivo 'importar.csv'.")
