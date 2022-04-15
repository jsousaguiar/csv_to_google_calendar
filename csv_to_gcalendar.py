import pandas as pd
import os
from criar_modelo import criar_modelo


def csv_to_gcalendar():
    try:
        df = pd.read_excel("modelo_calendario.xlsx")
    except FileNotFoundError:
        print(
            "Arquivo modelo não encontrado. O arquivo 'modelo_calendario.xlsx' foi criado."
            + "Preenha-o e rode o programa novamente."
        )
        criar_modelo()
        exit()

    if os.path.exists("importar.csv"):
        try:
            os.remove("importar.csv")
        except PermissionError:
            print(
                "O arquivo 'importar.csv' está aberto. Feche-o e rode o programa novamente."
            )
            exit()

    try:
        df.to_csv("importar.csv", sep=",", encoding="latin-1", index=False)
    except PermissionError:
        print(
            "O arquivo 'importar.csv' está aberto. Feche-o e rode o programa novamente."
        )
        exit()
    if os.path.exists("importar.csv"):
        print("Arquivo 'importar.csv' gerado com sucesso!")
    else:
        print("Não foi possível gerar o arquivo 'importar.csv'.")
    return None


if __name__ == "__main__":
    csv_to_gcalendar()
