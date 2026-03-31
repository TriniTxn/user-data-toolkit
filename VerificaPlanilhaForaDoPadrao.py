import pandas as pd

df = pd.read_excel(".xlsx")

df["NroMatricula"] = df["NroMatricula"].astype(str).str.strip()

df_filtrado = df[df["NroMatricula"].str.len() != 6]

df_filtrado.to_excel("MatriculasForaDoPadrao.xlsx", index=False)