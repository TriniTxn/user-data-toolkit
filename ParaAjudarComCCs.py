import pandas as pd

df = pd.read_excel(".xlsx")

df['Email'] = "'" + df['Email'].astype(str) + "'"

df.to_excel(".xlsx", index=False)
