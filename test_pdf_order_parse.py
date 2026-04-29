import sys

sys.path.insert(0, r"c:\Volume\shipping_ai")

from app_volum import map_uploaded_items_to_catalog, parse_items_from_pdf

pdf_path = r"c:\Volume\pedido_venda_948664170.pdf"

parsed = parse_items_from_pdf(pdf_path)
print("Itens extraidos do PDF:")
for item in parsed:
    print(item)

order_items, unknown = map_uploaded_items_to_catalog(parsed)
print("\nItens mapeados no catalogo:", order_items)
print("Itens nao encontrados:", unknown)
