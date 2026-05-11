import io
import json
import os
import platform
import re
import textwrap
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from pypdf import PdfReader, PdfWriter
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas

BASE_DIR = Path(__file__).resolve().parent
CONFIG_PATH = BASE_DIR / "layout_config.json"
DEFAULT_OUTPUT = BASE_DIR / "etiqueta_10x15.pdf"
MODELS_DIR = BASE_DIR / "modelos"
LAYOUTS_DIR = BASE_DIR / "layouts"
PDF_GLOB = "*.pdf"


def list_pdf_models() -> list[Path]:
    if not MODELS_DIR.exists():
        return []
    all_pdfs = sorted(MODELS_DIR.glob(PDF_GLOB), key=lambda p: p.name.lower())
    # Only consider PDFs that have an explicit per-model layout.
    # This prevents generating labels over unrelated PDFs (e.g. order documents).
    valid = [p for p in all_pdfs if (LAYOUTS_DIR / f"{p.stem}.json").exists()]
    return valid


def mm_to_pt(value_mm: float) -> float:
    return float(value_mm) * mm


def top_to_pdf_y(page_height_pt: float, y_from_top_mm: float) -> float:
    return page_height_pt - mm_to_pt(y_from_top_mm)


def load_layout_config() -> dict:
    with CONFIG_PATH.open("r", encoding="utf-8") as f:
        return json.load(f)


def load_layout_for_model(template_path: Path | None) -> tuple[dict, Path]:
    if template_path and template_path.exists():
        model_layout_path = LAYOUTS_DIR / f"{template_path.stem}.json"
        if model_layout_path.exists():
            with model_layout_path.open("r", encoding="utf-8") as f:
                return json.load(f), model_layout_path

    with CONFIG_PATH.open("r", encoding="utf-8") as f:
        return json.load(f), CONFIG_PATH


def save_layout_config(config_path: Path, layout: dict) -> None:
    config_path.parent.mkdir(parents=True, exist_ok=True)
    with config_path.open("w", encoding="utf-8") as f:
        json.dump(layout, f, ensure_ascii=True, indent=2)


def normalize_multiline(value: str) -> str:
    if not value:
        return ""
    lines = [line.strip() for line in value.replace("\r", "").split("\n")]
    return "\n".join([line for line in lines if line])


def wrap_text(value: str, max_chars: int) -> list[str]:
    if not value:
        return []
    lines = []
    for line in normalize_multiline(value).split("\n"):
        wrapped = textwrap.wrap(line, width=max_chars) or [""]
        lines.extend(wrapped)
    return lines


def draw_field(c: canvas.Canvas, page_height_pt: float, field_cfg: dict, text_value: str, base_font: dict) -> None:
    if not text_value:
        return

    x_pt = mm_to_pt(field_cfg.get("x_mm", 0))
    y_top_mm = field_cfg.get("y_mm", 0)
    y_pt = top_to_pdf_y(page_height_pt, y_top_mm)

    font_name = field_cfg.get("font", base_font.get("name", "Helvetica"))
    font_size = field_cfg.get("size", base_font.get("size", 11))
    c.setFont(font_name, font_size)

    max_width_mm = field_cfg.get("max_width_mm", 84)
    max_chars = max(6, int(max_width_mm * 1.9))

    multiline = field_cfg.get("multiline", False)
    line_gap_mm = base_font.get("line_gap_mm", 1.6)
    line_height_pt = font_size + mm_to_pt(line_gap_mm)

    if multiline:
        lines = wrap_text(text_value, max_chars)
    else:
        flat = " ".join(text_value.split())
        lines = wrap_text(flat, max_chars)[:1]

    for idx, line in enumerate(lines):
        c.drawString(x_pt, y_pt - (idx * line_height_pt), line)


def build_overlay_pdf(
    layout: dict,
    values: dict[str, str],
    override_width_pt: float | None = None,
    override_height_pt: float | None = None,
) -> bytes:
    page_width_pt = override_width_pt or mm_to_pt(layout["page"]["width_mm"])
    page_height_pt = override_height_pt or mm_to_pt(layout["page"]["height_mm"])

    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=(page_width_pt, page_height_pt))

    base_font = layout.get("font", {})
    fields = layout.get("fields", {})

    for field_name, field_cfg in fields.items():
        draw_field(c, page_height_pt, field_cfg, values.get(field_name, ""), base_font)

    c.showPage()
    c.save()
    return buffer.getvalue()


def merge_overlay_with_template(overlay_pdf: bytes, template_path: Path | None) -> PdfWriter:
    overlay_reader = PdfReader(io.BytesIO(overlay_pdf))
    overlay_page = overlay_reader.pages[0]

    writer = PdfWriter()
    if template_path and template_path.exists():
        template_reader = PdfReader(str(template_path))
        if not template_reader.pages:
            raise ValueError("O PDF modelo nao possui paginas.")
        base_page = template_reader.pages[0]
        base_page.merge_page(overlay_page)
        writer.add_page(base_page)
    else:
        writer.add_page(overlay_page)

    return writer


def build_and_merge(layout: dict, values: dict[str, str], template_path: Path | None) -> PdfWriter:
    """Build overlay at template dimensions (if available) and merge."""
    w_pt = h_pt = None
    if template_path and template_path.exists():
        tmpl = PdfReader(str(template_path))
        if tmpl.pages:
            page = tmpl.pages[0]
            w_pt = float(page.mediabox.width)
            h_pt = float(page.mediabox.height)
    overlay = build_overlay_pdf(layout, values, override_width_pt=w_pt, override_height_pt=h_pt)
    return merge_overlay_with_template(overlay, template_path)



def print_pdf_windows(pdf_path: Path) -> None:
    if platform.system().lower() != "windows":
        raise RuntimeError("Impressao automatica implementada apenas para Windows.")
    os.startfile(str(pdf_path), "print")


# ---------------------------------------------------------------------------
# NF-e / DANFE PDF parser
# ---------------------------------------------------------------------------

def _format_cpf(digits: str) -> str:
    return f"{digits[:3]}.{digits[3:6]}.{digits[6:9]}-{digits[9:]}"


def _format_cnpj(digits: str) -> str:
    return f"{digits[:2]}.{digits[2:5]}.{digits[5:8]}/{digits[8:12]}-{digits[12:]}"


def _format_tax_id(raw: str) -> str:
    digits = re.sub(r"\D", "", raw)
    if len(digits) == 14:
        return _format_cnpj(digits)
    if len(digits) == 11:
        return _format_cpf(digits)
    return raw.strip()


def _clean(value: str) -> str:
    return re.sub(r"\s+", " ", value).strip()


def _extract_pdf_text(pdf_path: Path) -> str:
    """Extract text from PDF using pdfminer (primary) or pypdf (fallback)."""
    # Primary: pdfminer.six — handles most NF-e / DANFE formats
    try:
        from io import StringIO
        from pdfminer.high_level import extract_text as pm_extract
        text = pm_extract(str(pdf_path))
        if text and text.strip():
            return text
    except Exception:
        pass

    # Fallback: pypdf
    reader = PdfReader(str(pdf_path))
    return "\n".join(page.extract_text() or "" for page in reader.pages)


def parse_nf_pdf(pdf_path: Path) -> dict:
    """Parse a Brazilian NF-e DANFE PDF and extract label fields.

    Strategy: pdfminer extracts DANFE text in "label line / value line" pairs.
    We split into lines, find each section and pick values positionally.
    """
    full_text = _extract_pdf_text(pdf_path)
    lines = [l.strip() for l in full_text.splitlines() if l.strip()]

    result: dict[str, str] = {
        "cliente": "",
        "cpf_cnpj": "",
        "endereco": "",
        "cidade_uf": "",
        "cep": "",
        "pedido": "",
        "rastreio": "",
        "ns": "",
        "obs": "",
    }

    # Labels that indicate a new field (not a data value)
    _LABEL_RE = re.compile(
        r'^(?:NOME|RAZ[AÃ]O|CNPJ|CPF|ENDERE|BAIRRO|MUNIC|CEP\b|'
        r'\bUF\b|FONE|DATA|HORA|INSCRI|PROTOCOLO|DESTINAT|NATUREZA|'
        r'TRANSPORT|FATURA|CALCULO|IMPOSTO|PRODUTO|SERIE\b|CFOP|NCM|'
        r'VALOR|ALIQ|QUANT)',
        re.IGNORECASE,
    )

    def _next_val(keyword: str, src: list) -> str:
        """Return the next non-label line after the line containing `keyword`."""
        kw_up = keyword.upper()
        for i, line in enumerate(src):
            if kw_up in line.upper():
                for j in range(i + 1, min(i + 5, len(src))):
                    cand = src[j].strip()
                    if cand and not _LABEL_RE.match(cand):
                        return cand
                break
        return ""

    # --- NF number → pedido ---------------------------------------------------
    for line in lines:
        m = re.search(r'N[°º]?\s*(\d{3,6})\b', line)
        if m and int(m.group(1)) > 100:
            result["pedido"] = "NF " + m.group(1)
            break

    # --- Narrow to DESTINATÁRIO section ---------------------------------------
    dest_start = next(
        (i for i, l in enumerate(lines) if "DESTINAT" in l.upper()), 0
    )
    dest_end = next(
        (i for i, l in enumerate(lines)
         if i > dest_start and re.search(r'FATURA|TRANSPORTAD|CALCULO\s+DO\s+IMP', l, re.IGNORECASE)),
        len(lines),
    )
    dest = lines[dest_start:dest_end]

    # --- cliente --------------------------------------------------------------
    result["cliente"] = (
        _next_val("NOME / RAZ", dest)
        or _next_val("RAZAO SOCIAL", dest)
        or _next_val("NOME", dest)
    )

    # --- cpf_cnpj: second CNPJ occurrence is the destinatário ----------------
    # The emitente CNPJ appears before DESTINAT; dest_lines only has dest.
    raw_id = (
        _next_val("CNPJ / CPF", dest)
        or _next_val("CPF / CNPJ", dest)
        or _next_val("CNPJ", dest)
    )
    if raw_id:
        result["cpf_cnpj"] = _format_tax_id(raw_id)

    # --- endereco + bairro ----------------------------------------------------
    end_val = _next_val("ENDERE", dest)
    bairro = _next_val("BAIRRO", dest)
    if end_val:
        if bairro and bairro.upper() not in end_val.upper():
            end_val = f"{end_val} - {bairro}"
        result["endereco"] = end_val

    # --- CEP ------------------------------------------------------------------
    cep_raw = _next_val("CEP", dest)
    if not cep_raw:
        m = re.search(r'\b(\d{5})[\-\.]?(\d{3})\b', " ".join(dest))
        cep_raw = f"{m.group(1)}-{m.group(2)}" if m else ""
    if cep_raw:
        digits = re.sub(r"\D", "", cep_raw)
        result["cep"] = f"{digits[:5]}-{digits[5:]}" if len(digits) == 8 else cep_raw.strip()

    # --- cidade_uf ------------------------------------------------------------
    city = _next_val("MUNIC", dest)
    uf_raw = _next_val("UF", dest)
    uf = re.sub(r"[^A-Z]", "", uf_raw.upper())[:2]
    if city:
        result["cidade_uf"] = f"{_clean(city)}/{uf}" if uf else _clean(city)
    else:
        m = re.search(r'([A-ZÀ-Ú][A-ZÀ-Úa-zà-ú ]{2,20})\s*/\s*([A-Z]{2})\b', " ".join(dest))
        if m:
            result["cidade_uf"] = f"{_clean(m.group(1))}/{m.group(2)}"

    # --- Produtos → ns (description lines from DADOS DOS PRODUTOS section) ----
    prod_start = next(
        (i for i, l in enumerate(lines)
         if re.search(r'DADOS\s+DO\s+PROD|DESCRI[ÇC][AÃ]O\s+DOS\s+PROD', l, re.IGNORECASE)),
        -1,
    )
    _PROD_SKIP = re.compile(
        r'^(?:COD|DESCRI|NCM|CFOP|UNID|QUANT|VALOR|ALIQ|CST|B\.?\s*CAL|'
        r'NUMER|PESO|LIQUID|BRUTO|UNIT|IPI\b|\d{1,3}[,\.]\d{2}$)',
        re.IGNORECASE,
    )
    if prod_start != -1:
        prods: list[str] = []
        for pl in lines[prod_start + 1: prod_start + 40]:
            if re.search(r'CALCULO|ISSQN|DADOS ADIC|INFORMA', pl, re.IGNORECASE):
                break
            if _LABEL_RE.match(pl) or _PROD_SKIP.match(pl):
                continue
            # Skip pure-number lines (prices, qty)
            if re.match(r'^[\d\s,.]+$', pl):
                continue
            if len(pl) > 5:
                prods.append(pl)
        if prods:
            result["ns"] = "\n".join(prods)

    return result


class EtiquetaApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("ETIQ - Impressao de Etiquetas 10x15 cm")
        self.root.geometry("850x690")

        self.template_path = tk.StringVar(value="")
        self.model_choice = tk.StringVar(value="")
        self.output_path = tk.StringVar(value=str(DEFAULT_OUTPUT))
        self.last_generated = None
        self.model_map: dict[str, Path] = {}

        self.entries: dict[str, tk.Text | ttk.Entry] = {}

        self._build_ui()

    def _build_ui(self) -> None:
        top = ttk.Frame(self.root, padding=12)
        top.pack(fill=tk.BOTH, expand=True)

        ttk.Label(top, text="Etiqueta fisica: 10 cm x 15 cm", font=("Segoe UI", 11, "bold")).pack(anchor="w")

        # --- NF import --------------------------------------------------------
        nf_frame = ttk.LabelFrame(top, text="Importar Nota Fiscal (PDF)", padding=10)
        nf_frame.pack(fill=tk.X, pady=(10, 4))

        self.nf_path_var = tk.StringVar(value="")
        ttk.Entry(nf_frame, textvariable=self.nf_path_var, state="readonly", width=50).pack(
            side=tk.LEFT, fill=tk.X, expand=True,
        )
        ttk.Button(nf_frame, text="Selecionar NF (PDF)", command=self.import_nf).pack(side=tk.LEFT, padx=8)
        ttk.Button(nf_frame, text="Gerar AMBAS etiquetas", command=self.generate_both_labels).pack(
            side=tk.LEFT, padx=(0, 4),
        )

        # --- Model selector ---------------------------------------------------
        model_frame = ttk.LabelFrame(top, text="Modelo PDF (opcional)", padding=10)
        model_frame.pack(fill=tk.X, pady=(10, 8))

        self.model_combo = ttk.Combobox(model_frame, textvariable=self.model_choice, state="readonly", width=32)
        self.model_combo.pack(side=tk.LEFT, padx=(0, 8))
        self.model_combo.bind("<<ComboboxSelected>>", self.on_model_selected)
        ttk.Button(model_frame, text="Atualizar lista", command=self.refresh_models).pack(side=tk.LEFT, padx=(0, 8))

        ttk.Entry(model_frame, textvariable=self.template_path).pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Button(model_frame, text="Selecionar PDF", command=self.select_template).pack(side=tk.LEFT, padx=8)

        data_frame = ttk.LabelFrame(top, text="Dados da Etiqueta", padding=10)
        data_frame.pack(fill=tk.BOTH, expand=False, pady=6)

        self._add_entry(data_frame, "cliente", "Cliente", 0)
        self._add_entry(data_frame, "cpf_cnpj", "CPF/CNPJ", 1)
        self._add_entry(data_frame, "endereco", "Endereco", 2)
        self._add_entry(data_frame, "cidade_uf", "Cidade/UF", 3)
        self._add_entry(data_frame, "cep", "CEP", 4)
        self._add_entry(data_frame, "pedido", "Pedido", 5)
        self._add_entry(data_frame, "rastreio", "Rastreio", 6)
        self._add_text(data_frame, "ns", "N. Serie (um por linha)", 7, height=4)
        self._add_text(data_frame, "obs", "Observacoes", 8, height=3)

        output_frame = ttk.LabelFrame(top, text="Saida", padding=10)
        output_frame.pack(fill=tk.X, pady=8)

        ttk.Entry(output_frame, textvariable=self.output_path).pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Button(output_frame, text="Salvar como...", command=self.select_output).pack(side=tk.LEFT, padx=8)

        action = ttk.Frame(top)
        action.pack(fill=tk.X, pady=(8, 0))

        ttk.Button(action, text="Gerar PDF", command=self.generate_pdf).pack(side=tk.LEFT)
        ttk.Button(action, text="Gerar teste", command=self.generate_test_pdf).pack(side=tk.LEFT, padx=8)
        ttk.Button(action, text="Imprimir", command=self.print_last).pack(side=tk.LEFT, padx=8)
        ttk.Button(action, text="Abrir PDF", command=self.open_pdf).pack(side=tk.LEFT)
        ttk.Button(action, text="Abrir layout atual", command=self.open_current_layout).pack(side=tk.LEFT, padx=8)
        ttk.Button(action, text="Ajuste rapido", command=self.open_quick_adjust).pack(side=tk.LEFT)

        self.refresh_models()

    def _add_entry(self, parent: ttk.Frame, key: str, label: str, row: int) -> None:
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", pady=4)
        entry = ttk.Entry(parent)
        entry.grid(row=row, column=1, sticky="ew", pady=4, padx=(8, 0))
        parent.grid_columnconfigure(1, weight=1)
        self.entries[key] = entry

    def _add_text(self, parent: ttk.Frame, key: str, label: str, row: int, height: int = 3) -> None:
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="nw", pady=4)
        txt = tk.Text(parent, height=height, wrap=tk.WORD)
        txt.grid(row=row, column=1, sticky="ew", pady=4, padx=(8, 0))
        parent.grid_columnconfigure(1, weight=1)
        self.entries[key] = txt

    def refresh_models(self) -> None:
        MODELS_DIR.mkdir(parents=True, exist_ok=True)
        models = list_pdf_models()
        self.model_map = {path.name: path for path in models}

        names = list(self.model_map.keys())
        if not names:
            names = ["(sem modelos em ETIQ/modelos)"]

        self.model_combo["values"] = names
        self.model_combo.current(0)

        if names[0].startswith("("):
            self.model_choice.set(names[0])
        else:
            first = self.model_map[names[0]]
            self.model_choice.set(names[0])
            self.template_path.set(str(first))

    def on_model_selected(self, _event=None) -> None:
        selected = self.model_choice.get().strip()
        if selected in self.model_map:
            self.template_path.set(str(self.model_map[selected]))

    def import_nf(self) -> None:
        path = filedialog.askopenfilename(
            title="Selecionar Nota Fiscal (PDF)",
            filetypes=[("PDF", "*.pdf")],
        )
        if not path:
            return
        self.nf_path_var.set(path)
        try:
            data = parse_nf_pdf(Path(path))
        except Exception as exc:
            messagebox.showerror("Erro ao ler NF", str(exc))
            return

        # Fill all form fields with parsed values
        for key, value in data.items():
            widget = self.entries.get(key)
            if widget is None:
                continue
            if isinstance(widget, tk.Text):
                widget.delete("1.0", tk.END)
                widget.insert("1.0", value)
            else:
                widget.delete(0, tk.END)
                widget.insert(0, value)

        messagebox.showinfo(
            "NF importada",
            "Dados extraídos da nota fiscal.\n\nRevise os campos e clique em "
            "\"Gerar AMBAS etiquetas\" para gerar os dois modelos de uma vez,\n"
            "ou escolha um modelo e use \"Gerar PDF\" para gerar individualmente.",
        )

    def generate_both_labels(self) -> None:
        """Generate one PDF for EACH model in MODELS_DIR using the current form values."""
        models = list_pdf_models()
        if not models:
            messagebox.showwarning("Sem modelos", "Nenhum PDF encontrado em ETIQ/modelos.")
            return

        try:
            values = self.collect_values()
        except Exception as exc:
            messagebox.showerror("Erro", str(exc))
            return

        saida_dir = BASE_DIR / "saida"
        saida_dir.mkdir(parents=True, exist_ok=True)
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")

        generated: list[Path] = []
        errors: list[str] = []

        for model_path in models:
            try:
                layout, _ = load_layout_for_model(model_path)
                writer = build_and_merge(layout, values, model_path)

                stem = re.sub(r"[^\w\-]", "_", model_path.stem)
                out_path = saida_dir / f"{ts}_{stem}.pdf"
                with out_path.open("wb") as f:
                    writer.write(f)

                generated.append(out_path)
            except Exception as exc:
                errors.append(f"{model_path.name}: {exc}")

        if generated:
            self.last_generated = generated[-1]
            msg = "Etiquetas geradas:\n" + "\n".join(str(p) for p in generated)
            if errors:
                msg += "\n\nErros:\n" + "\n".join(errors)
            messagebox.showinfo("Concluído", msg)
            for out_path in generated:
                os.startfile(str(out_path))
        else:
            messagebox.showerror("Erro", "Nenhuma etiqueta gerada.\n\n" + "\n".join(errors))

    def select_template(self) -> None:
        path = filedialog.askopenfilename(
            title="Selecionar PDF modelo",
            filetypes=[("PDF", PDF_GLOB)],
        )
        if path:
            self.template_path.set(path)
            self.model_choice.set("arquivo externo")

    def select_output(self) -> None:
        path = filedialog.asksaveasfilename(
            title="Salvar etiqueta",
            defaultextension=".pdf",
            filetypes=[("PDF", "*.pdf")],
            initialfile="etiqueta_10x15.pdf",
        )
        if path:
            self.output_path.set(path)

    def collect_values(self) -> dict[str, str]:
        values: dict[str, str] = {}
        for key, widget in self.entries.items():
            if isinstance(widget, tk.Text):
                raw = widget.get("1.0", tk.END)
            else:
                raw = widget.get()
            values[key] = raw.strip()
        return values

    def get_current_template(self) -> Path | None:
        raw = self.template_path.get().strip()
        return Path(raw) if raw else None

    def get_current_layout(self) -> tuple[dict, Path]:
        return load_layout_for_model(self.get_current_template())

    def generate_with_values(self, values: dict[str, str]) -> None:
        layout, layout_path = self.get_current_layout()
        template = self.get_current_template()
        writer = build_and_merge(layout, values, template)

        out_path = Path(self.output_path.get().strip() or str(DEFAULT_OUTPUT))
        out_path.parent.mkdir(parents=True, exist_ok=True)
        with out_path.open("wb") as f:
            writer.write(f)

        self.last_generated = out_path
        messagebox.showinfo("OK", f"Etiqueta gerada com sucesso:\n{out_path}\n\nLayout usado:\n{layout_path}")

    def generate_pdf(self) -> None:
        try:
            values = self.collect_values()
            self.generate_with_values(values)
        except Exception as exc:
            messagebox.showerror("Erro", str(exc))

    def generate_test_pdf(self) -> None:
        test_values = {
            "cliente": "CLIENTE TESTE ALINHAMENTO",
            "cpf_cnpj": "00.000.000/0001-00",
            "endereco": "Rua Exemplo, 123",
            "cidade_uf": "Sao Paulo/SP",
            "cep": "00000-000",
            "pedido": "PED-123456",
            "rastreio": "BR123456789BR",
            "ns": "PRD00001 - NS000001\nPRD00002 - NS000002",
            "obs": "Teste de impressao 10x15",
        }
        try:
            self.generate_with_values(test_values)
        except Exception as exc:
            messagebox.showerror("Erro", str(exc))

    def open_current_layout(self) -> None:
        try:
            _layout, layout_path = self.get_current_layout()
            os.startfile(str(layout_path))
        except Exception as exc:
            messagebox.showerror("Erro", str(exc))

    def open_quick_adjust(self) -> None:
        try:
            layout, layout_path = self.get_current_layout()
        except Exception as exc:
            messagebox.showerror("Erro", str(exc))
            return

        win = tk.Toplevel(self.root)
        win.title("Ajuste rapido de layout (mm)")
        win.geometry("460x300")
        win.transient(self.root)
        win.grab_set()

        ttk.Label(win, text=f"Layout: {layout_path.name}", font=("Segoe UI", 10, "bold")).pack(anchor="w", padx=12, pady=(10, 8))

        field_names = ["__TODOS__"] + list(layout.get("fields", {}).keys())
        field_var = tk.StringVar(value=field_names[0])
        step_var = tk.StringVar(value="1.0")

        row1 = ttk.Frame(win)
        row1.pack(fill=tk.X, padx=12, pady=6)
        ttk.Label(row1, text="Campo").pack(side=tk.LEFT)
        field_combo = ttk.Combobox(row1, textvariable=field_var, values=field_names, state="readonly", width=28)
        field_combo.pack(side=tk.LEFT, padx=8)

        row2 = ttk.Frame(win)
        row2.pack(fill=tk.X, padx=12, pady=6)
        ttk.Label(row2, text="Passo (mm)").pack(side=tk.LEFT)
        ttk.Entry(row2, textvariable=step_var, width=10).pack(side=tk.LEFT, padx=8)

        info = ttk.Label(
            win,
            text="Use os botoes para mover o campo selecionado.\nSe escolher __TODOS__, move todos os campos juntos.",
        )
        info.pack(anchor="w", padx=12, pady=6)

        def apply_delta(dx_mm: float, dy_mm: float) -> None:
            try:
                step = float(step_var.get().replace(",", ".").strip())
            except ValueError:
                messagebox.showerror("Erro", "Passo invalido. Exemplo: 1.0")
                return

            sel = field_var.get().strip()
            fields = layout.get("fields", {})
            targets = fields.keys() if sel == "__TODOS__" else [sel]

            for key in targets:
                if key not in fields:
                    continue
                current_x = float(fields[key].get("x_mm", 0))
                current_y = float(fields[key].get("y_mm", 0))
                fields[key]["x_mm"] = round(current_x + (dx_mm * step), 2)
                fields[key]["y_mm"] = round(current_y + (dy_mm * step), 2)

            save_layout_config(layout_path, layout)
            messagebox.showinfo("OK", "Layout salvo. Gere um teste para validar.")

        arrows = ttk.Frame(win)
        arrows.pack(pady=10)

        ttk.Button(arrows, text="↑", width=8, command=lambda: apply_delta(0, -1)).grid(row=0, column=1, padx=6, pady=4)
        ttk.Button(arrows, text="←", width=8, command=lambda: apply_delta(-1, 0)).grid(row=1, column=0, padx=6, pady=4)
        ttk.Button(arrows, text="→", width=8, command=lambda: apply_delta(1, 0)).grid(row=1, column=2, padx=6, pady=4)
        ttk.Button(arrows, text="↓", width=8, command=lambda: apply_delta(0, 1)).grid(row=2, column=1, padx=6, pady=4)

        bottom = ttk.Frame(win)
        bottom.pack(fill=tk.X, padx=12, pady=8)
        ttk.Button(bottom, text="Gerar teste agora", command=self.generate_test_pdf).pack(side=tk.LEFT)
        ttk.Button(bottom, text="Fechar", command=win.destroy).pack(side=tk.RIGHT)

    def print_last(self) -> None:
        if not self.last_generated or not Path(self.last_generated).exists():
            messagebox.showwarning("Atencao", "Gere a etiqueta antes de imprimir.")
            return
        try:
            print_pdf_windows(Path(self.last_generated))
            messagebox.showinfo("Impressao", "Comando de impressao enviado para o Windows.")
        except Exception as exc:
            messagebox.showerror("Erro", str(exc))

    def open_pdf(self) -> None:
        if not self.last_generated or not Path(self.last_generated).exists():
            messagebox.showwarning("Atencao", "Gere a etiqueta antes de abrir.")
            return
        os.startfile(str(self.last_generated))


if __name__ == "__main__":
    root = tk.Tk()
    app = EtiquetaApp(root)
    root.mainloop()
