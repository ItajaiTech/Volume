#!/usr/bin/env python
import sys
sys.path.insert(0, 'c:/Volume/shipping_ai')

from database import import_products_from_excel, list_products, init_db

# Initialize DB
db_path = "c:/Volume/shipping_ai/database.db"
init_db(db_path)

# Test with the template file including SKU
template_path = "c:/Volume/Produtos.xlsx"
column_mapping = {
    'sku': 'SKU',
    'name': 'Descrição do produto',
    'length': 'Comprimento',
    'width': 'Largura',
    'height': 'altura',
    'weight': 'Peso',
}

imported, errors, msgs = import_products_from_excel(db_path, template_path, column_mapping)
print(f"Importados: {imported}")
print(f"Erros: {errors}")
if msgs:
    print(f"Primeiras mensagens: {msgs[:2]}")

# Check products
produtos = list_products(db_path)
print(f"\nTotal de produtos no banco: {len(produtos)}")
if len(produtos) > 0:
    print("\nPrimeiros 3 produtos:")
    for p in produtos[:3]:
        print(f"  - {p['name']}: {p['length_cm']}x{p['width_cm']}x{p['height_cm']} cm, Peso: {p['weight']}")
