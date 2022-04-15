import pandas as pd
import pandas.io.formats.excel


def criar_modelo():
    """
    Cria modelo para ser preenchido pelo usuário

    """
    # Formata o cabeçalho do arquivo .xlsx
    pandas.io.formats.excel.ExcelFormatter.header_style = None

    modelo = pd.DataFrame(
        columns=[
            "Subject",
            "Start Date",
            "Start Time",
            "End Date",
            "End Time",
            "All Day Event",
            "Description",
            "Location",
            "Private",
        ]
    )

    lines = [
        [
            "Reunião Teste 1",
            "04/15/2022",
            "15:10:00",
            "04/15/2022",
            "16:00:00",
            "False",
            "Testando modificação da descrição",
            "Online",
            "True",
        ],
        [
            "Reunião Teste 2",
            "04/15/2022",
            "16:10:00",
            "04/15/2022",
            "17:00:00",
            "False",
            "Testando a importação do CSV",
            "",
            "True",
        ],
        [
            "Reunião Teste 3",
            "04/15/2022",
            "",
            "04/15/2022",
            "",
            "True",
            "Testando a importação do CSV",
            "",
            "True",
        ],
    ]

    for line in lines:
        modelo.loc[len(modelo)] = line
    modelo.to_excel("modelo_calendario.xlsx", header=True, index=False)
    return None
