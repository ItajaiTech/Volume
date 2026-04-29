#!/usr/bin/env python
import os
import sys
sys.path.insert(0, 'c:/Volume/shipping_ai')

from database import import_products_from_excel

# Test with the template file
template_path = "c:/Volume/Produtos.xlsx"
db_path = "c:/Volume/shipping_ai/database.db"
column_mapping = {
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
    print(f"Primeiras mensagens: {msgs[:3]}")
print("Teste concluído com sucesso!")
