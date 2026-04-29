from pypdf import PdfReader

pdf_path = r"c:\Volume\pedido_venda_948664170.pdf"
reader = PdfReader(pdf_path)

all_lines = []
for page_idx, page in enumerate(reader.pages, start=1):
    text = page.extract_text() or ""
    lines = [" ".join(line.split()) for line in text.splitlines() if line.strip()]
    print(f"--- PAGINA {page_idx} ({len(lines)} linhas) ---")
    for i, line in enumerate(lines[:120], start=1):
        print(f"{i:03d}: {line}")
    all_lines.extend(lines)

print("\nTOTAL LINHAS:", len(all_lines))
