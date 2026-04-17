# =============================================================================
# ENGINES - CALCULATIONS & PDF
# =============================================================================
import re
from fpdf import FPDF

def format_real(valor):
    return f"R$ {valor:,.2f}".replace(",", "x").replace(".", ",").replace("x", ".")

def format_id_or_phone(text, type_format):
    nums = re.sub(r'\D', '', text)
    if type_format == "phone":
        if len(nums) <= 10: return re.sub(r'(\d{2})(\d{4})(\d{4})', r'(\1) \2-\3', nums)
        return re.sub(r'(\d{2})(\d{5})(\d{4})', r'(\1) \2-\3', nums)
    elif type_format == "tax_id":
        if len(nums) <= 11: return re.sub(r'(\d{3})(\d{3})(\d{3})(\d{2})', r'\1.\2.\3-\4', nums)
        return re.sub(r'(\d{2})(\d{3})(\d{3})(\d{4})(\d{2})', r'\1.\2.\3/\4-\5', nums)
    return nums

def _calc_base(largura, altura, preco, tipo, max_dim=3000):
    area = (largura / 1000) * (altura / 1000)
    valid = largura <= max_dim and altura <= max_dim
    prob = ""
    if not valid: prob = "Largura/Altura excedem" if largura > max_dim and altura > max_dim else f"Dimensão excede"
    return {"tipo": tipo, "area_m2": area, "br_price": format_real(area * preco), "br_m2": format_real(preco), "is_valid": valid, "max_dim": max_dim, "problema": prob}

def calcular_flexdoor(l, a): return _calc_base(l, a, 1790.0 if l > 1200 else 1590.0, "Dupla" if l > 1200 else "Simples")
def calcular_pvc(l, a): 
    res = _calc_base(l, a, 260.0, "Cortina PVC", 99999)
    res["espessura_especial"] = a > 4000
    return res
def calcular_peca_pvc(l, a): return _calc_base(l, a, 790.0, "Peça PVC")
def calcular_deslocamento(km): return {"br_valor": format_real(km * 2 * 3)}
def calcular_acomodacao(d, f): return {"br_valor": format_real(((350*d) + (100*d*f)) * 1.1)}

def gerar_pdf_orcamento(dados, servicos, produtos, obs):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", "B", 16)
    pdf.cell(190, 10, "Base de Orçamento - NC Portas", ln=True, align='C')
    pdf.set_font("Arial", "", 10)
    pdf.ln(10)
    pdf.multi_cell(190, 7, f"Cliente: {dados['nome']} | Empresa: {dados['empresa']}\nFone: {dados['fone']} | CNPJ: {dados['cnpj']}\nVão: {dados['dimensoes']} | Área: {dados['area']}")
    
    for label, items in [("Serviços", servicos), ("Produtos", produtos)]:
        if items:
            pdf.ln(5); pdf.set_font("Arial", "B", 12)
            pdf.cell(190, 10, label, ln=True)
            pdf.set_font("Arial", "", 10)
            for i in items:
                pdf.cell(150, 7, f"- {i['item']}", border=1)
                pdf.cell(40, 7, f"Qtd: {i['qtd']}", border=1, ln=True)
    if obs:
        pdf.ln(5); pdf.set_font("Arial", "B", 12); pdf.cell(190, 10, "Observações", ln=True)
        pdf.set_font("Arial", "", 10); pdf.multi_cell(190, 7, obs)
    name = f"Orcamento_{dados['empresa'].replace(' ', '_')}.pdf"
    pdf_out = pdf.output(dest='S')
    return bytes(pdf_out) if isinstance(pdf_out, bytearray) else pdf_out.encode('latin-1')