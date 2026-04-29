import pandas as pd
df = pd.read_excel('c:\Volume\Produtos.xlsx')
print("Colunas:", df.columns.tolist())
print("\nForma:", df.shape)
print("\nPrimeiras 3 linhas:")
print(df.head(3).to_string())
