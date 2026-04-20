# =============================================================================
# ENGINES - CALCULATIONS & PDF GENERATION
# Pure business logic. No Streamlit imports here.
# =============================================================================
import re
import unicodedata
from fpdf import FPDF

# --- SECTION: FORMATTING HELPERS ---
def format_real(valor):
    """Formats a float to Brazilian currency string: R$ 1.234,56"""
    return f"R$ {valor:,.2f}".replace(",", "x").replace(".", ",").replace("x", ".")

def format_id_or_phone(text, type_format):
    """Applies input masks. type_format: 'phone' | 'tax_id' (CPF or CNPJ)."""
    nums = re.sub(r'\D', '', text)
    if type_format == "phone":
        if len(nums) <= 10:
            return re.sub(r'(\d{2})(\d{4})(\d{4})', r'(\1) \2-\3', nums)
        return re.sub(r'(\d{2})(\d{5})(\d{4})', r'(\1) \2-\3', nums)
    elif type_format == "tax_id":
        if len(nums) <= 11:
            return re.sub(r'(\d{3})(\d{3})(\d{3})(\d{2})', r'\1.\2.\3-\4', nums)
        return re.sub(r'(\d{2})(\d{3})(\d{3})(\d{4})(\d{2})', r'\1.\2.\3/\4-\5', nums)
    return nums

def _safe_latin1(text):
    """Normalizes a string to latin-1 safe characters for fpdf1 compatibility.
    Decomposes accented chars (e.g. ã→a) so the PDF never crashes on
    Brazilian names. Remove this function once fully migrated to fpdf2."""
    return unicodedata.normalize('NFKD', str(text)).encode('latin-1', 'ignore').decode('latin-1')

# --- SECTION: CALCULATION ENGINES ---
def _calc_base(largura, altura, preco, tipo, max_dim=3000):
    """Base calculator for all area-priced products.
    Returns a result dict with formatted values and a validity flag."""
    area  = (largura / 1000) * (altura / 1000)
    valid = largura <= max_dim and altura <= max_dim
    prob  = ""
    if not valid:
        if largura > max_dim and altura > max_dim:
            prob = "Largura e Altura excedem o máximo"
        else:
            prob = "Largura excede o máximo" if largura > max_dim else "Altura excede o máximo"
    return {
        "tipo":     tipo,
        "area_m2":  area,
        "br_price": format_real(area * preco),
        "br_m2":    format_real(preco),
        "is_valid": valid,
        "max_dim":  max_dim,
        "problema": prob,
    }

def calcular_flexdoor(largura, altura):
    """Calculates Flexdoor pricing. Dupla if largura > 1200mm, else Simples."""
    preco = 1790.0 if largura > 1200 else 1590.0
    tipo  = "Dupla" if largura > 1200 else "Simples"
    return _calc_base(largura, altura, preco, tipo)

def calcular_pvc(largura, altura):
    """Calculates PVC Curtain pricing. No dimension limit."""
    res = _calc_base(largura, altura, 260.0, "Cortina PVC", max_dim=99999)
    res["espessura_especial"] = altura > 4000
    return res

def calcular_peca_pvc(largura, altura):
    """Calculates PVC Piece pricing."""
    return _calc_base(largura, altura, 790.0, "Peça PVC")

def calcular_deslocamento(km):
    """Calculates travel fee: km * 2 (round trip) * R$3/km."""
    return {"br_valor": format_real(km * 2 * 3)}

def calcular_acomodacao(dias, funcionarios):
    """Calculates accommodation fee with 10% overhead.
    Formula: ((R$350/day) + (R$100/day * staff)) * 1.1"""
    return {"br_valor": format_real(((350 * dias) + (100 * dias * funcionarios)) * 1.1)}

# --- SECTION: PDF GENERATION ---
def gerar_pdf_orcamento(dados, servicos, produtos, obs):
    """Generates a quote PDF in memory and returns raw bytes.

    Args:
        dados:    dict with keys: nome, empresa, email, fone, cnpj, dimensoes, area
        servicos: list of dicts with keys: item, preco, qtd
        produtos: list of dicts with keys: item, preco, qtd
        obs:      str, additional observations

    Returns:
        bytes: raw PDF content ready for st.download_button
    """
    pdf = FPDF()
    pdf.add_page()

    # Header
    pdf.set_font("Arial", "B", 16)
    pdf.cell(190, 10, _safe_latin1("Base de Orçamento - NC Portas"), ln=True, align='C')
    pdf.ln(10)

    # Client Info Block
    pdf.set_font("Arial", "", 10)
    client_block = (
        f"Cliente: {dados['nome']} | Empresa: {dados['empresa']}\n"
        f"Email: {dados['email']} | Fone: {dados['fone']}\n"
        f"CNPJ: {dados['cnpj']}\n"
        f"Vão Informado: {dados['dimensoes']} | Área: {dados['area']}"
    )
    pdf.multi_cell(190, 7, _safe_latin1(client_block), border=0)

    # Item Tables (Services and Products)
    for label, items in [("Serviços", servicos), ("Produtos", produtos)]:
        if not items:
            continue
        pdf.ln(8)

        # Table Section Label
        pdf.set_font("Arial", "B", 12)
        pdf.cell(190, 10, _safe_latin1(label), ln=True, align='L')

        # Table Header Row
        pdf.set_fill_color(240, 240, 240)
        pdf.set_font("Arial", "B", 10)
        pdf.cell(140, 7, _safe_latin1(" Descrição"),   border=1, fill=True)
        pdf.cell(30,  7, "Preco Unit.",                border=1, fill=True, align='L')
        pdf.cell(20,  7, "Qtd",                        border=1, fill=True, align='C', ln=True)

        # Table Data Rows
        pdf.set_font("Arial", "", 10)
        for item in items:
            pdf.cell(140, 7, _safe_latin1(f" {item['item']}"), border=1)
            pdf.cell(30,  7, format_real(item['preco']),        border=1, align='L')
            pdf.cell(20,  7, str(item['qtd']),                  border=1, align='C', ln=True)

    # Observations Block
    if obs:
        pdf.ln(10)
        pdf.set_font("Arial", "B", 12)
        pdf.cell(190, 10, _safe_latin1("Observações"), ln=True)
        pdf.set_font("Arial", "", 10)
        pdf.multi_cell(190, 7, _safe_latin1(obs), border=1)

    # Output to memory buffer
    pdf_out = pdf.output(dest='S')
    return bytes(pdf_out) if isinstance(pdf_out, bytearray) else pdf_out.encode('latin-1')
