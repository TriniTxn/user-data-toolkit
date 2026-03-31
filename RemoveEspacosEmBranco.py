import pandas as pd

df = pd.read_excel(".xlsx")

df["Codigo_Centro_de_Custo"] = (
    df["Codigo_Centro_de_Custo"]
        .astype(str)              # garante string
        .str.replace("\u00A0", "", regex=False)  # remove espaço não-quebrável
        .str.strip()               # remove espaços antes e depois
)

df.to_excel("espacoTratado.xlsx", index=False)
