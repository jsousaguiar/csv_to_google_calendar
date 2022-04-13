import pandas as pd

df = pd.read_excel("modelo_calendario.xlsx")

print(df)
df.to_csv("importar.csv", sep=",", encoding="latin-1", index=False)
