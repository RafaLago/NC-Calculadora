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

# --- SECTION: EXCEL REPORT GENERATION ---
# Generates monthly invoice reports by writing raw sheet XML and copying all
# styles, theme, and assets directly from the model files. This guarantees
# pixel-perfect output — no openpyxl style mapping, no index mismatch.
#
# DEPENDENCY: The two model files must sit in the same folder as this module:
#   NC_Portas_-_NFe_-_Relatório_Março.xlsx
#   NC_Portas_-_NFSe_-_Relatório_Março.xlsx

import io as _io, os as _os, zipfile as _zipfile, re as _re, xml.sax.saxutils as _saxutils

_BASE          = _os.path.dirname(__file__)
_NFE_TEMPLATE  = _os.path.join(_BASE, "NC_Portas_-_NFe_-_Relatório_Março.xlsx")
_NFSE_TEMPLATE = _os.path.join(_BASE, "NC_Portas_-_NFSe_-_Relatório_Março.xlsx")


def _esc(val):
    """XML-escape a cell value string."""
    return _saxutils.escape(str(val)) if val is not None else ""


def _build_xlsx(template_path, sheet_title, sheet_xml_bytes, table_xml_bytes):
    """Assembles the final .xlsx by copying all assets from the template
    and replacing only the sheet and table XML with new content.

    Returns bytes.
    """
    buf = _io.BytesIO()
    with _zipfile.ZipFile(template_path, 'r') as src, \
         _zipfile.ZipFile(buf, 'w', compression=_zipfile.ZIP_DEFLATED) as dst:

        for item in src.infolist():
            data = src.read(item.filename)

            # Skip calcChain — it caches formula order from the template's row
            # coordinates. Copying it causes Excel repair errors when row count
            # differs from the model. Excel rebuilds it automatically on open.
            if item.filename == 'xl/calcChain.xml':
                continue

            if item.filename == 'xl/worksheets/sheet1.xml':
                data = sheet_xml_bytes
            elif 'xl/tables/' in item.filename:
                data = table_xml_bytes
            elif item.filename == 'xl/workbook.xml':
                # Update sheet title
                data = _re.sub(
                    rb'name="[^"]*"',
                    f'name="{sheet_title}"'.encode(),
                    data, count=1
                )

            dst.writestr(item, data)

    return buf.getvalue()


# ---------------------------------------------------------------------------
# NF-e NC
# ---------------------------------------------------------------------------
# Exact style indices from model NF-e styles.xml:
#   Row stripe A (white fill=0): cols A-H → s= 1  2  3  4  3  5  6  7
#   Row stripe B (grey  fill=2): cols A-H → s= 8  9 10 11 10 12 13 14
#   Header row:                  cols A-H → s=37 38 39 40 39 41 42 43
#   Footer label col A:  s=34   Footer value col B: qty=35, total=36

_NFE_HDR_S  = [37, 38, 39, 40, 39, 41, 42, 43]   # header row, 8 cols
_NFE_ROW_A  = [ 1,  2,  3,  4,  3,  5,  6,  7]   # odd  data rows (white)
_NFE_ROW_B  = [ 8,  9, 10, 11, 10, 12, 13, 14]   # even data rows (grey)

_NFE_COLS_XML = (
    '<cols>'
    '<col min="1" max="1" width="52.625" bestFit="1" customWidth="1"/>'
    '<col min="2" max="2" width="16.125" bestFit="1" customWidth="1"/>'
    '<col min="3" max="3" width="19.625" bestFit="1" customWidth="1"/>'
    '<col min="4" max="4" width="10.875" bestFit="1" customWidth="1"/>'
    '<col min="5" max="5" width="58.875" bestFit="1" customWidth="1"/>'
    '<col min="6" max="6" width="52.5"   bestFit="1" customWidth="1"/>'
    '<col min="7" max="7" width="14.125" bestFit="1" customWidth="1"/>'
    '<col min="8" max="8" width="13.5"   bestFit="1" customWidth="1"/>'
    '</cols>'
)

_NFE_SHEET_VIEW = (
    '<sheetViews>'
    '<sheetView showGridLines="0" tabSelected="1" zoomScale="90" zoomScaleNormal="90" workbookViewId="0">'
    '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>'
    '<selection pane="bottomLeft"/>'
    '</sheetView>'
    '</sheetViews>'
)

_NFE_COLS = ["CLIENTE", "NF-e", "VALOR NF-e", "CFOP", "NATUREZA", "CHAVE", "EMISSÃO", "STATUS"]
_NFE_DFCOLS = ["CLIENTE", "NF-e", "VALOR NF-e", "CFOP", "NATUREZA", "CHAVE", "EMISSÃO", "STATUS"]


def _nfe_cell(col_idx, row, val, s, is_numeric=False, is_date=False):
    """Render a single <c> element. col_idx is 0-based."""
    col_letters = ["A","B","C","D","E","F","G","H"]
    ref = f"{col_letters[col_idx]}{row}"
    if is_numeric and val != "" and val is not None:
        return f'<c r="{ref}" s="{s}"><v>{val}</v></c>'
    elif is_date and val != "" and val is not None:
        return f'<c r="{ref}" s="{s}"><v>{val}</v></c>'
    else:
        v = _esc(val)
        return f'<c r="{ref}" s="{s}" t="inlineStr"><is><t xml:space="preserve">{v}</t></is></c>'


def gerar_relatorio_nfe(df, mes_ano):
    """Generates the NF-e NC monthly report Excel file with exact model formatting.

    Args:
        df:      DataFrame filtered for the target month.
        mes_ano: str label e.g. 'Março 2026'

    Returns:
        bytes: raw .xlsx content ready for st.download_button
    """
    import pandas as pd

    last_data_row = len(df) + 1   # row 1 = header, data starts row 2
    n_cols        = 8
    col_letters   = "ABCDEFGH"

    # --- Header row ---
    hdr_cells = ""
    for ci, (hdr, s) in enumerate(zip(_NFE_COLS, _NFE_HDR_S)):
        ref = f"{col_letters[ci]}1"
        hdr_cells += f'<c r="{ref}" s="{s}" t="inlineStr"><is><t>{_esc(hdr)}</t></is></c>'
    header_row = (
        f'<row r="1" spans="1:{n_cols}" ht="17.25" thickBot="1" x14ac:dyDescent="0.35">'
        f'{hdr_cells}</row>'
    )

    # --- Data rows ---
    data_rows = ""
    for ri, (_, row) in enumerate(df.iterrows()):
        er  = ri + 2
        idx = _NFE_ROW_A if ri % 2 == 0 else _NFE_ROW_B
        cells = ""

        for ci, col in enumerate(_NFE_DFCOLS):
            s   = idx[ci]
            val = row.get(col, "")

            if col == "EMISSÃO":
                # Convert to Excel serial date number
                try:
                    dt = pd.to_datetime(val)
                    serial = (dt - pd.Timestamp("1899-12-30")).days
                    cells += f'<c r="{col_letters[ci]}{er}" s="{s}"><v>{serial}</v></c>'
                except Exception:
                    cells += f'<c r="{col_letters[ci]}{er}" s="{s}" t="inlineStr"><is><t>{_esc(val)}</t></is></c>'
            elif col in ("NF-e", "VALOR NF-e", "CFOP"):
                try:
                    cells += f'<c r="{col_letters[ci]}{er}" s="{s}"><v>{float(val)}</v></c>'
                except Exception:
                    cells += f'<c r="{col_letters[ci]}{er}" s="{s}" t="inlineStr"><is><t>{_esc(val)}</t></is></c>'
            else:
                v = _esc(str(val).strip()) if val == val and val is not None else ""
                cells += f'<c r="{col_letters[ci]}{er}" s="{s}" t="inlineStr"><is><t xml:space="preserve">{v}</t></is></c>'

        data_rows += (
            f'<row r="{er}" spans="1:{n_cols}" x14ac:dyDescent="0.3">'
            f'{cells}</row>'
        )

    # --- Blank separator row and footer rows ---
    blank_row   = last_data_row + 1
    footer_row1 = last_data_row + 2
    footer_row2 = last_data_row + 3

    footer = (
        f'<row r="{blank_row}" spans="1:{n_cols}" ht="17.25" thickBot="1" x14ac:dyDescent="0.35"/>'
        f'<row r="{footer_row1}" spans="1:{n_cols}" ht="17.25" thickBot="1" x14ac:dyDescent="0.35">'
        f'<c r="A{footer_row1}" s="34" t="inlineStr"><is><t>Quantidade Notas</t></is></c>'
        f'<c r="B{footer_row1}" s="35"><f>COUNT(NFe[NF-e])</f><v>0</v></c>'
        f'</row>'
        f'<row r="{footer_row2}" spans="1:{n_cols}" ht="17.25" thickBot="1" x14ac:dyDescent="0.35">'
        f'<c r="A{footer_row2}" s="34" t="inlineStr"><is><t>Valor Total</t></is></c>'
        f'<c r="B{footer_row2}" s="36"><f>SUM(NFe[VALOR NF-e])</f><v>0</v></c>'
        f'</row>'
    )

    # --- Conditional formatting (sqref covers data rows only) ---
    cf_xml = (
        f'<conditionalFormatting sqref="B2:H{last_data_row}">'
        f'<cfRule type="expression" dxfId="2" priority="1"><formula>$H2="Em Andamento"</formula></cfRule>'
        f'<cfRule type="expression" dxfId="1" priority="2"><formula>$H2="Emitida"</formula></cfRule>'
        f'<cfRule type="expression" dxfId="0" priority="3"><formula>$H2="Cancelada"</formula></cfRule>'
        f'</conditionalFormatting>'
    )

    # --- Assemble full sheet XML ---
    sheet_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
        ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
        ' xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">'
        f'{_NFE_SHEET_VIEW}'
        '<sheetFormatPr defaultRowHeight="16.5" x14ac:dyDescent="0.3"/>'
        f'{_NFE_COLS_XML}'
        '<sheetData>'
        f'{header_row}{data_rows}'
        '</sheetData>'
        '<phoneticPr fontId="4" type="noConversion"/>'
        f'{cf_xml}'
        '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>'
        '</worksheet>'
    )

    # --- Table XML ---
    table_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
        ' id="1" name="NFe" displayName="NFe"'
        f' ref="A1:H{last_data_row}" totalsRowShown="0">'
        f'<autoFilter ref="A1:H{last_data_row}"/>'
        '<tableColumns count="8">'
        '<tableColumn id="1" name="CLIENTE"/>'
        '<tableColumn id="3" name="NF-e" dataCellStyle="Currency"/>'
        '<tableColumn id="4" name="VALOR NF-e" dataCellStyle="Currency"/>'
        '<tableColumn id="5" name="CFOP" dataCellStyle="Currency"/>'
        '<tableColumn id="6" name="NATUREZA" dataCellStyle="Currency"/>'
        '<tableColumn id="7" name="CHAVE" dataCellStyle="Currency"/>'
        '<tableColumn id="8" name="EMISSÃO"/>'
        '<tableColumn id="9" name="STATUS"/>'
        '</tableColumns>'
        '<tableStyleInfo name="TableStyleMedium15" showFirstColumn="0"'
        ' showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>'
        '</table>'
    )

    return _build_xlsx(
        _NFE_TEMPLATE,
        "NF-e NC",
        sheet_xml.encode("utf-8"),
        table_xml.encode("utf-8"),
    )


# ---------------------------------------------------------------------------
# NFS-e NC
# ---------------------------------------------------------------------------
# Exact style indices from model NFS-e styles.xml:
#   Row A (white): A=1  B=3  C=4  D=5  E=6  F=7  G=8  H=9  I=10 J=11 K=12 L=13 M=2
#   Row B (grey):  A=13 B=16 C=17 D=18 E=19 F=20 G=21 H=22 I=23 J=24 K=25 L=26 M=15
#   Header:        A=64 B=65 C=66 D=65 E=67 F=68 G=69 H=70 I=71 J=72 K=72 L=72 M=72
#   Footer label s=60, qty s=61, valor s=62, liquido s=63

_NFSE_HDR_S = [64, 65, 66, 65, 67, 68, 69, 70, 71, 72, 72, 72, 72]
_NFSE_ROW_A = [ 1,  3,  4,  5,  6,  7,  8,  9, 10, 11, 12, 13,  2]
_NFSE_ROW_B = [13, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 15]

_NFSE_COLS_XML = (
    '<cols>'
    '<col min="1"  max="1"  width="49.625" bestFit="1" customWidth="1"/>'
    '<col min="2"  max="2"  width="15"     bestFit="1" customWidth="1"/>'
    '<col min="3"  max="3"  width="20.625" bestFit="1" customWidth="1"/>'
    '<col min="4"  max="4"  width="9.125"  bestFit="1" customWidth="1"/>'
    '<col min="5"  max="5"  width="58.875" bestFit="1" customWidth="1"/>'
    '<col min="6"  max="6"  width="11.375" bestFit="1" customWidth="1"/>'
    '<col min="7"  max="7"  width="13.25"  bestFit="1" customWidth="1"/>'
    '<col min="8"  max="8"  width="19.125" bestFit="1" customWidth="1"/>'
    '<col min="9"  max="9"  width="15.25"  bestFit="1" customWidth="1"/>'
    '<col min="10" max="10" width="12.125" bestFit="1" customWidth="1"/>'
    '<col min="11" max="11" width="18.875" bestFit="1" customWidth="1"/>'
    '<col min="12" max="12" width="13.875" bestFit="1" customWidth="1"/>'
    '<col min="13" max="13" width="23.625" bestFit="1" customWidth="1"/>'
    '</cols>'
)

_NFSE_SHEET_VIEW = (
    '<sheetViews>'
    '<sheetView showGridLines="0" tabSelected="1" zoomScale="90" zoomScaleNormal="90" workbookViewId="0">'
    '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>'
    '<selection pane="bottomLeft"/>'
    '</sheetView>'
    '</sheetViews>'
)

_NFSE_COLS    = ["CLIENTE","NFS-e","VALOR NFS-e","DPS","CHAVE","DATA","STATUS",
                 "OBSERVAÇÕES","RETER ISS","ISS","RETER INSS","INSS","VALOR LIQUIDO"]
_NFSE_NUMERIC = {"NFS-e","VALOR NFS-e","DPS","ISS","INSS","VALOR LIQUIDO"}
_NFSE_LETTERS = list("ABCDEFGHIJKLM")


def gerar_relatorio_nfse(df, mes_ano):
    """Generates the NFS-e NC monthly report Excel file with exact model formatting.

    Args:
        df:      DataFrame filtered for the target month.
        mes_ano: str label e.g. 'Março 2026'

    Returns:
        bytes: raw .xlsx content ready for st.download_button
    """
    import pandas as pd

    last_data_row = len(df) + 1
    n_cols        = 13

    # --- Header row ---
    hdr_cells = ""
    for ci, (hdr, s) in enumerate(zip(_NFSE_COLS, _NFSE_HDR_S)):
        ref = f"{_NFSE_LETTERS[ci]}1"
        hdr_cells += f'<c r="{ref}" s="{s}" t="inlineStr"><is><t>{_esc(hdr)}</t></is></c>'
    header_row = (
        f'<row r="1" spans="1:{n_cols}" ht="17.25" thickBot="1" x14ac:dyDescent="0.35">'
        f'{hdr_cells}</row>'
    )

    # --- Data rows ---
    data_rows = ""
    for ri, (_, row) in enumerate(df.iterrows()):
        er  = ri + 2
        idx = _NFSE_ROW_A if ri % 2 == 0 else _NFSE_ROW_B
        cells = ""

        for ci, col in enumerate(_NFSE_COLS):
            s   = idx[ci]
            val = row.get(col, "")

            if col == "DATA":
                try:
                    dt     = pd.to_datetime(val)
                    serial = (dt - pd.Timestamp("1899-12-30")).days
                    cells += f'<c r="{_NFSE_LETTERS[ci]}{er}" s="{s}"><v>{serial}</v></c>'
                except Exception:
                    cells += f'<c r="{_NFSE_LETTERS[ci]}{er}" s="{s}" t="inlineStr"><is><t>{_esc(val)}</t></is></c>'
            elif col in _NFSE_NUMERIC:
                try:
                    cells += f'<c r="{_NFSE_LETTERS[ci]}{er}" s="{s}"><v>{float(val)}</v></c>'
                except Exception:
                    cells += f'<c r="{_NFSE_LETTERS[ci]}{er}" s="{s}" t="inlineStr"><is><t>{_esc(val)}</t></is></c>'
            else:
                v = _esc(str(val).strip()) if val == val and val is not None else ""
                cells += f'<c r="{_NFSE_LETTERS[ci]}{er}" s="{s}" t="inlineStr"><is><t xml:space="preserve">{v}</t></is></c>'

        data_rows += (
            f'<row r="{er}" spans="1:{n_cols}" x14ac:dyDescent="0.3">'
            f'{cells}</row>'
        )

    # --- Footer ---
    blank_row   = last_data_row + 1
    fr1         = last_data_row + 2
    fr2         = last_data_row + 3
    fr3         = last_data_row + 4

    footer = (
        f'<row r="{blank_row}" spans="1:{n_cols}" ht="17.25" thickBot="1" x14ac:dyDescent="0.35"/>'
        f'<row r="{fr1}" spans="1:{n_cols}" ht="17.25" thickBot="1" x14ac:dyDescent="0.35">'
        f'<c r="A{fr1}" s="60" t="inlineStr"><is><t>Quantidade Notas</t></is></c>'
        f'<c r="B{fr1}" s="61"><f>COUNT(NFSe[NFS-e])</f><v>0</v></c>'
        f'</row>'
        f'<row r="{fr2}" spans="1:{n_cols}" ht="17.25" thickBot="1" x14ac:dyDescent="0.35">'
        f'<c r="A{fr2}" s="60" t="inlineStr"><is><t>Valor Total</t></is></c>'
        f'<c r="B{fr2}" s="62"><f>SUM(NFSe[VALOR NFS-e])</f><v>0</v></c>'
        f'</row>'
        f'<row r="{fr3}" spans="1:{n_cols}" ht="17.25" thickBot="1" x14ac:dyDescent="0.35">'
        f'<c r="A{fr3}" s="60" t="inlineStr"><is><t>Valor Liquido</t></is></c>'
        f'<c r="B{fr3}" s="63"><f>SUM(NFSe[VALOR LIQUIDO])</f><v>0</v></c>'
        f'</row>'
    )

    # --- Conditional formatting ---
    cf_xml = (
        f'<conditionalFormatting sqref="B2:H{last_data_row}">'
        f'<cfRule type="expression" dxfId="9" priority="8"><formula>$G2="Cancelada"</formula></cfRule>'
        f'<cfRule type="expression" dxfId="8" priority="9"><formula>$G2="Substituida"</formula></cfRule>'
        f'<cfRule type="expression" dxfId="7" priority="10"><formula>$G2="Emitida"</formula></cfRule>'
        f'</conditionalFormatting>'
        f'<conditionalFormatting sqref="I2:J{last_data_row}">'
        f'<cfRule type="expression" dxfId="6" priority="6"><formula>$I2="NÃO"</formula></cfRule>'
        f'<cfRule type="expression" dxfId="5" priority="7"><formula>$I2="SIM"</formula></cfRule>'
        f'</conditionalFormatting>'
        f'<conditionalFormatting sqref="K2:L{last_data_row}">'
        f'<cfRule type="expression" dxfId="4" priority="4"><formula>$K2="NÃO"</formula></cfRule>'
        f'<cfRule type="expression" dxfId="3" priority="5"><formula>$K2="SIM"</formula></cfRule>'
        f'</conditionalFormatting>'
        f'<conditionalFormatting sqref="M2:M{last_data_row}">'
        f'<cfRule type="expression" dxfId="2" priority="1"><formula>$G2="Cancelada"</formula></cfRule>'
        f'<cfRule type="expression" dxfId="1" priority="2"><formula>$G2="Substituida"</formula></cfRule>'
        f'<cfRule type="expression" dxfId="0" priority="3"><formula>$G2="Emitida"</formula></cfRule>'
        f'</conditionalFormatting>'
    )

    # --- Assemble sheet XML ---
    sheet_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
        ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
        ' xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">'
        f'{_NFSE_SHEET_VIEW}'
        '<sheetFormatPr defaultRowHeight="16.5" x14ac:dyDescent="0.3"/>'
        f'{_NFSE_COLS_XML}'
        '<sheetData>'
        f'{header_row}{data_rows}{footer}'
        '</sheetData>'
        '<phoneticPr fontId="4" type="noConversion"/>'
        f'{cf_xml}'
        '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>'
        '</worksheet>'
    )

    # --- Table XML ---
    table_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
        ' id="1" name="NFSe" displayName="NFSe"'
        f' ref="A1:M{last_data_row}" totalsRowShown="0">'
        f'<autoFilter ref="A1:M{last_data_row}"/>'
        '<tableColumns count="13">'
        '<tableColumn id="1"  name="CLIENTE"/>'
        '<tableColumn id="2"  name="NFS-e"/>'
        '<tableColumn id="3"  name="VALOR NFS-e"/>'
        '<tableColumn id="4"  name="DPS"/>'
        '<tableColumn id="5"  name="CHAVE"/>'
        '<tableColumn id="6"  name="DATA"/>'
        '<tableColumn id="7"  name="STATUS"/>'
        '<tableColumn id="8"  name="OBSERVAÇÕES"/>'
        '<tableColumn id="9"  name="RETER ISS"/>'
        '<tableColumn id="10" name="ISS"/>'
        '<tableColumn id="11" name="RETER INSS"/>'
        '<tableColumn id="12" name="INSS"/>'
        '<tableColumn id="13" name="VALOR LIQUIDO"/>'
        '</tableColumns>'
        '<tableStyleInfo name="TableStyleMedium15" showFirstColumn="0"'
        ' showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>'
        '</table>'
    )

    return _build_xlsx(
        _NFSE_TEMPLATE,
        "NFS-e NC",
        sheet_xml.encode("utf-8"),
        table_xml.encode("utf-8"),
    )
