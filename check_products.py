#!/usr/bin/env python
import sys
sys.path.insert(0, 'c:/Volume/shipping_ai')

from database import list_products

db_path = "c:/Volume/shipping_ai/database.db"
produtos = list_products(db_path)
print(f"Total de produtos: {len(produtos)}")
if len(produtos) > 0:
    print("Primeiros 3 produtos:")
    for p in produtos[:3]:
        print(f"  - {p['name']}: {p['length_cm']}x{p['width_cm']}x{p['height_cm']} cm, Peso: {p['weight']}")
