#!/usr/bin/env python
import sys
sys.path.insert(0, 'c:/Volume/shipping_ai')

import pandas as pd

template_path = "c:/Volume/Produtos.xlsx"
df = pd.read_excel(template_path)
print(f"Total de linhas: {len(df)}")
print(f"Colunas: {df.columns.tolist()}")
print("\nPrimeiras 3 linhas:")
print(df[['Descrição do produto', 'Comprimento', 'Largura', 'altura', 'Peso']].head(3))

print("\nVerificando valores:")
for idx, row in df.head(3).iterrows():
    name = str(row.get('Descrição do produto', '')).strip()
    length = float(row.get('Comprimento', 0))
    width = float(row.get('Largura', 0))
    height = float(row.get('altura', 0))
    weight = float(row.get('Peso', 0))
    print(f"Linha {idx}: '{name}' - {length}x{width}x{height}, peso={weight}")
    print(f"  Valores validos? name={bool(name)}, dims={length>0 and width>0 and height>0}, peso={weight>0}")
