import pandas as pd

# carregar o arquivo
df = pd.read_excel(".xlsx")

# pegar a primeira coluna automaticamente
col = df.columns[0]

# garantir string limpa
df[col] = df[col].astype(str).str.strip()

# extrair RED, classificação e fila
df[['RED', 'CLASSIFICACAO', 'FILA']] = df[col].str.extract(
    r'(_?RED\d+)\s+\.C\s*(\d+)\.*\s*(\d+)',
    expand=True
)

# remover "_" do RED
df['RED'] = df['RED'].str.replace('_', '', regex=False)

# salvar novo arquivo
df.to_excel("planilha_tratada.xlsx", index=False)
