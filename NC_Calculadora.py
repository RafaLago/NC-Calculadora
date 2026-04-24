# =============================================================================
# NC CALCULADORA - MAIN UI
# Entry point. Streamlit layout, session state, and tab routing only.
# All business logic lives in engines.py. All DB reads in database.py.
# =============================================================================
import streamlit as st
import os
import engines as eng
import database as db

# --- SECTION: PAGE CONFIG & STYLES ---
st.set_page_config(page_title="Calculadora NC Portas", layout="wide")

st.markdown("""
    <style>
        .block-container {
            padding-top: 2.5rem;
            padding-bottom: 0rem;
            margin-top: 0rem;
        }
        header[data-testid="stHeader"] {
            height: 3.5rem;
            background: rgba(0,0,0,0);
        }
        @import url('https://fonts.cdnfonts.com/css/gill-sans-nova');
        * { font-family: 'Gill Sans Nova', sans-serif; }
        .stMetric { border: 1px solid #ddd; padding: 10px; border-radius: 10px; }
    </style>
""", unsafe_allow_html=True)

# --- SECTION: LOGO & TITLE ---
logo_path = os.path.join(os.path.dirname(__file__), "logo.png")
if os.path.exists(logo_path):
    st.image(logo_path, width=300)
else:
    st.warning("Logo não encontrada.")

st.title("Calculadora NC Portas")

# --- SECTION: SESSION STATE INITIALIZATION ---
_state_defaults = {
    "lista_servicos":      [],
    "lista_produtos":      [],
    "lista_s_manusa":      [],
    "lista_p_manusa":      [],
    "lista_portas_manusa": [],
    "pdf_nc":              None,   # bytes | None — generated PDF for Tab 3
    "pdf_nc_filename":     "",
    "pdf_manusa":          None,   # bytes | None — generated PDF for Tab 4
    "pdf_manusa_filename": "",
    "rel_xlsx":             None,   # bytes | None — generated report
    "rel_xlsx_name":        "",
}
for _key, _val in _state_defaults.items():
    if _key not in st.session_state:
        st.session_state[_key] = _val

# --- SECTION: DB LOAD (once per rerun) ---
s_list, p_list = db.get_unique_items()

# --- SECTION: SHARED CALLBACKS ---

# Price sync callbacks — one per selectbox key
def sync_price_s_nc():
    item = st.session_state.sel_s_orc
    if item:
        st.session_state.p_s_orc = db.get_item_price(item, "Serviços")

def sync_price_p_nc():
    item = st.session_state.sel_p_orc
    if item:
        st.session_state.p_p_orc = db.get_item_price(item, "Produtos")

def sync_price_s_manusa():
    item = st.session_state.sel_s_man
    if item:
        st.session_state.p_s_man = db.get_item_price(item, "Serviços")

def sync_price_p_manusa():
    item = st.session_state.sel_p_man
    if item:
        st.session_state.p_p_man = db.get_item_price(item, "Produtos")

# Mask callbacks — one per field key to avoid key detection ambiguity
def fmt_phone_nc():
    st.session_state.fone_in = eng.format_id_or_phone(st.session_state.fone_in, "phone")

def fmt_taxid_nc():
    st.session_state.id_in = eng.format_id_or_phone(st.session_state.id_in, "tax_id")

def fmt_phone_manusa():
    st.session_state.fone_m = eng.format_id_or_phone(st.session_state.fone_m, "phone")

def fmt_taxid_manusa():
    st.session_state.id_m = eng.format_id_or_phone(st.session_state.id_m, "tax_id")

# Clear callbacks — explicit type reset per key (no string-matching heuristics)
def clear_nc_callback():
    st.session_state.lista_servicos    = []
    st.session_state.lista_produtos    = []
    st.session_state.pdf_nc            = None
    st.session_state.pdf_nc_filename   = ""
    str_keys   = ["nome_in", "empresa_in", "fone_in", "email_in", "id_in", "obs_in",
                  "sel_s_orc", "sel_p_orc"]
    float_keys = ["p_s_orc", "p_p_orc"]
    for k in str_keys:
        if k in st.session_state:
            st.session_state[k] = ""
    for k in float_keys:
        if k in st.session_state:
            st.session_state[k] = 0.0

def clear_manusa_callback():
    st.session_state.lista_s_manusa        = []
    st.session_state.lista_p_manusa        = []
    st.session_state.lista_portas_manusa   = []
    st.session_state.pdf_manusa            = None
    st.session_state.pdf_manusa_filename   = ""
    str_keys   = ["nome_m", "emp_m", "fone_m", "email_m", "id_m", "obs_m",
                  "sel_s_man", "sel_p_man"]
    float_keys = ["p_s_man", "p_p_man"]
    for k in str_keys:
        if k in st.session_state:
            st.session_state[k] = ""
    for k in float_keys:
        if k in st.session_state:
            st.session_state[k] = 0.0

# List reorder/delete helper
def update_list(lista_name, index, action):
    lista = st.session_state[lista_name]
    if action == "up"   and index > 0:
        lista[index], lista[index - 1] = lista[index - 1], lista[index]
    elif action == "down" and index < len(lista) - 1:
        lista[index], lista[index + 1] = lista[index + 1], lista[index]
    elif action == "delete":
        lista.pop(index)
    st.rerun()

# --- SECTION: SIDEBAR ---
with st.sidebar:
    st.header("🚚 Deslocamento")
    km = st.number_input("Distância KM (Ida)", value=0)
    if km > 0:
        st.metric("Taxa Deslocamento", eng.calcular_deslocamento(km)['br_valor'])

    st.write("---")
    st.header("🛏️ Acomodação")
    dias = st.number_input("Dias",         min_value=0)
    func = st.number_input("Funcionários", min_value=0)
    if dias > 0 and func > 0:
        st.metric("Taxa Acomodação", eng.calcular_acomodacao(dias, func)['br_valor'])

    st.write("---")
    st.header("🔍 Consulta de Preços")

    sel_s = st.selectbox("Serviços", [""] + s_list)
    if sel_s:
        st.metric("Valor", eng.format_real(db.get_item_price(sel_s, "Serviços")))
        hist_s = db.get_last_10_entries(sel_s, "Serviços")
        if hist_s is not None:
            st.dataframe(hist_s, hide_index=True, use_container_width=True)

    sel_p = st.selectbox("Produtos", [""] + p_list)
    if sel_p:
        st.metric("Valor", eng.format_real(db.get_item_price(sel_p, "Produtos")))
        hist_p = db.get_last_10_entries(sel_p, "Produtos")
        if hist_p is not None:
            st.dataframe(hist_p, hide_index=True, use_container_width=True)

# --- SECTION: MAIN TABS ---
tabs = st.tabs(["Flexdoor", "Cortina de PVC", "Peça de PVC", "Orçamento NC Portas", "Orçamento Manusa", "Notas Fiscais"])

# --- TAB 0: FLEXDOOR ---
with tabs[0]:
    st.subheader("Informe as Dimensões do Vão (Flexdoor)")
    col_in1, col_in2 = st.columns(2)
    largura = col_in1.number_input("Largura (mm)", value=2000, key="w_flex")
    altura  = col_in2.number_input("Altura (mm)",  value=2100, key="h_flex")
    res = eng.calcular_flexdoor(largura, altura)
    if res["is_valid"]:
        c1, c2 = st.columns(2)
        c1.metric("Tipo",        res["tipo"])
        c1.metric("Área do Vão", f"{res['area_m2']:.2f} m²")
        c2.metric("Valor Total", res['br_price'])
        c2.metric("Preço do m²", res['br_m2'])
        st.success(f"Porta Flexdoor {res['tipo']} nas Dimensões {largura}mm x {altura}mm.")
    else:
        st.error(f"Atenção: {res['problema']} (Máx: {res['max_dim']}mm)")

# --- TAB 1: CORTINA DE PVC ---
with tabs[1]:
    st.subheader("Informe as Dimensões do Vão (Cortina)")
    col_in1, col_in2 = st.columns(2)
    largura = col_in1.number_input("Largura (mm)", value=2000, key="w_pvc")
    altura  = col_in2.number_input("Altura (mm)",  value=2100, key="h_pvc")
    res = eng.calcular_pvc(largura, altura)
    c1, c2 = st.columns(2)
    c1.metric("Área do Vão", f"{res['area_m2']:.2f} m²")
    c2.metric("Valor Total", res['br_price'])
    st.success(f"{res['tipo']} nas Dimensões {largura}mm x {altura}mm.")
    if res["espessura_especial"]:
        st.warning("Atenção: Necessária espessura especial para altura > 4m")

# --- TAB 2: PEÇA DE PVC ---
with tabs[2]:
    st.subheader("Informe as Dimensões do Vão (Peça)")
    col_in1, col_in2 = st.columns(2)
    largura = col_in1.number_input("Largura (mm)", value=2000, key="w_peca")
    altura  = col_in2.number_input("Altura (mm)",  value=2100, key="h_peca")
    res = eng.calcular_peca_pvc(largura, altura)
    if res["is_valid"]:
        c1, c2 = st.columns(2)
        c1.metric("Tipo",        res["tipo"])
        c1.metric("Área do Vão", f"{res['area_m2']:.2f} m²")
        c2.metric("Valor Total", res['br_price'])
        c2.metric("Preço do m²", res['br_m2'])
        st.success(f"{res['tipo']} nas Dimensões {largura}mm x {altura}mm.")
    else:
        st.error(f"Erro: {res['problema']}")

# --- TAB 3: ORÇAMENTO NC PORTAS ---
with tabs[3]:
    st.header("📝 Orçamento Base NC")

    # Client Data
    st.subheader("👤 Dados do Cliente")
    with st.container(border=True):
        c1, c2 = st.columns(2)
        nome    = c1.text_input("Nome",     key="nome_in")
        empresa = c2.text_input("Empresa",  key="empresa_in")
        fone    = c1.text_input("Telefone", placeholder="Ex: 41999999999",
                                key="fone_in", on_change=fmt_phone_nc)
        email   = c2.text_input("Email",    key="email_in")
        id_doc  = st.text_input("CPF/CNPJ", placeholder="Apenas números",
                                key="id_in",  on_change=fmt_taxid_nc)

    # Door Configuration (optional — dimensions only, no door list)
    with st.container(border=True):
        st.subheader("📏 Configuração da Porta")
        col_in1, col_in2 = st.columns(2)
        w_orc = col_in1.number_input("Largura (mm)", value=2000, key="w_orc")
        h_orc = col_in2.number_input("Altura (mm)",  value=2100, key="h_orc")
        st.success(f"Porta Automática nas Dimensões {w_orc}mm x {h_orc}mm.")

    # Services
    st.subheader("🛠️ Serviços")
    with st.container(border=True):
        col1, col2, col3 = st.columns([3, 1, 1])
        item_s  = col1.selectbox("Buscar Serviço", [""] + s_list,
                                  key="sel_s_orc", on_change=sync_price_s_nc)
        preco_s = col2.number_input("Preço (R$)", key="p_s_orc", step=10.0)
        qtd_s   = col3.number_input("Quant", min_value=1, step=1, key="q_s_orc")
        if st.button("Adicionar Serviço", key="btn_s_nc", use_container_width=True):
            if item_s:
                st.session_state.lista_servicos.append(
                    {"item": item_s, "preco": preco_s, "qtd": qtd_s,
                     "total": preco_s * qtd_s}
                )
                st.session_state.pdf_nc = None   # invalidate previous PDF on list change
                st.rerun()

    for i, s in enumerate(st.session_state.lista_servicos):
        r = st.columns([0.4, 0.4, 3, 1, 0.5, 0.8])
        if r[0].button("▲", key=f"us_{i}"):      update_list("lista_servicos", i, "up")
        if r[1].button("▼", key=f"ds_{i}"):      update_list("lista_servicos", i, "down")
        r[2].write(s['item'])
        r[3].write(eng.format_real(s['preco']))
        r[4].write(f"x{s['qtd']}")
        if r[5].button("🗑️", key=f"del_s_{i}"): update_list("lista_servicos", i, "delete")

    # Products
    st.subheader("📦 Produtos")
    with st.container(border=True):
        col1, col2, col3 = st.columns([3, 1, 1])
        item_p  = col1.selectbox("Buscar Produto", [""] + p_list,
                                  key="sel_p_orc", on_change=sync_price_p_nc)
        preco_p = col2.number_input("Preço (R$)", key="p_p_orc", step=10.0)
        qtd_p   = col3.number_input("Quant", min_value=1, step=1, key="q_p_orc")
        if st.button("Adicionar Produto", key="btn_p_nc", use_container_width=True):
            if item_p:
                st.session_state.lista_produtos.append(
                    {"item": item_p, "preco": preco_p, "qtd": qtd_p,
                     "total": preco_p * qtd_p}
                )
                st.session_state.pdf_nc = None   # invalidate previous PDF on list change
                st.rerun()

    for i, p in enumerate(st.session_state.lista_produtos):
        r = st.columns([0.4, 0.4, 3, 1, 0.5, 0.8])
        if r[0].button("▲", key=f"up_{i}"):      update_list("lista_produtos", i, "up")
        if r[1].button("▼", key=f"dp_{i}"):      update_list("lista_produtos", i, "down")
        r[2].write(p['item'])
        r[3].write(eng.format_real(p['preco']))
        r[4].write(f"x{p['qtd']}")
        if r[5].button("🗑️", key=f"del_p_{i}"): update_list("lista_produtos", i, "delete")

    # Observations
    st.subheader("📝 Observações")
    obs = st.text_area("Notas Adicionais", key="obs_in")

    # PDF Generation & Clear
    st.write("---")
    c_pdf, c_clear = st.columns(2)
    with c_pdf:
        if st.session_state.lista_servicos or st.session_state.lista_produtos:
            if st.button("🚀 Gerar e Baixar Orçamento PDF", key="btn_pdf_nc",
                         use_container_width=True):
                dados_pdf = {
                    "nome":      nome,
                    "empresa":   empresa,
                    "email":     email,
                    "fone":      fone,
                    "cnpj":      id_doc,
                    "dimensoes": f"{w_orc}mm x {h_orc}mm",
                    "area":      f"{(w_orc * h_orc) / 1_000_000:.2f} m²",
                }
                st.session_state.pdf_nc          = eng.gerar_pdf_orcamento(
                    dados_pdf,
                    st.session_state.lista_servicos,
                    st.session_state.lista_produtos,
                    obs,
                )
                st.session_state.pdf_nc_filename = f"Orcamento_{empresa}.pdf"

            if st.session_state.pdf_nc:
                st.download_button(
                    label="📥 Baixar PDF",
                    data=st.session_state.pdf_nc,
                    file_name=st.session_state.pdf_nc_filename,
                    mime="application/pdf",
                    use_container_width=True,
                )
        else:
            st.info("Adicione um Serviço ou Produto para Gerar o Orçamento.")

    with c_clear:
        st.button("🗑️ Limpar Tudo", key="btn_clr_nc",
                  on_click=clear_nc_callback, use_container_width=True)

# --- TAB 4: ORÇAMENTO PORTA MANUSA ---
with tabs[4]:
    st.header("🚪 Porta Automática Manusa")

    # Client Data
    st.subheader("👤 Dados do Cliente")
    with st.container(border=True):
        c1, c2 = st.columns(2)
        nome_m  = c1.text_input("Nome",     key="nome_m")
        emp_m   = c2.text_input("Empresa",  key="emp_m")
        fone_m  = c1.text_input("Telefone", placeholder="Ex: 41999999999",
                                key="fone_m", on_change=fmt_phone_manusa)
        email_m = c2.text_input("Email",    key="email_m")
        id_m    = st.text_input("CPF/CNPJ", placeholder="Apenas números",
                                key="id_m",  on_change=fmt_taxid_manusa)

    # Door Configuration
    with st.container(border=True):
        st.subheader("📏 Configuração da Porta")
        qtd_m = st.number_input("Quantidade Total de Portas", min_value=1, value=1,
                                key="q_man_total")

        if qtd_m == 1:
            col_m1, col_m2 = st.columns(2)
            w_m = col_m1.number_input("Largura (mm)", value=2000, key="w_man_single")
            h_m = col_m2.number_input("Altura (mm)",  value=2100, key="h_man_single")
            final_doors_summary = [f"01 Porta Automática nas Dimensões {w_m}mm x {h_m}mm"]
            st.success(final_doors_summary[0])
        else:
            col_add1, col_add2, col_add3 = st.columns([2, 2, 1.2])
            w_add = col_add1.number_input("Largura (mm)", value=2000, key="w_man_add")
            h_add = col_add2.number_input("Altura (mm)",  value=2100, key="h_man_add")
            with col_add3:
                st.write(" ")
                st.markdown("<div style='height: 4px;'></div>", unsafe_allow_html=True)
                if st.button("Adicionar Porta", key="btn_add_manusa", use_container_width=True):
                    current_total = sum(p['qtd'] for p in st.session_state.lista_portas_manusa)
                    if current_total < qtd_m:
                        dim_str = f"{w_add}mm x {h_add}mm"
                        found   = False
                        for p in st.session_state.lista_portas_manusa:
                            if p['dim'] == dim_str:
                                p['qtd'] += 1
                                found = True
                                break
                        if not found:
                            st.session_state.lista_portas_manusa.append(
                                {"dim": dim_str, "qtd": 1}
                            )
                        st.session_state.pdf_manusa = None   # invalidate on door list change
                        st.rerun()
                    else:
                        st.error(f"Limite de {qtd_m} portas atingido.")

            final_doors_summary = []
            for i, p in enumerate(st.session_state.lista_portas_manusa):
                label        = "Porta Automática" if p['qtd'] == 1 else "Portas Automáticas"
                summary_line = f"{p['qtd']:02d} {label} nas Dimensões {p['dim']}"
                final_doors_summary.append(summary_line)
                r = st.columns([4, 1])
                r[0].write(summary_line)
                if r[1].button("🗑️", key=f"del_door_{i}", use_container_width=True):
                    update_list("lista_portas_manusa", i, "delete")

    # Services
    st.subheader("🛠️ Serviços")
    with st.container(border=True):
        col1, col2, col3 = st.columns([3, 1, 1])
        item_s_m  = col1.selectbox("Buscar Serviço", [""] + s_list,
                                    key="sel_s_man", on_change=sync_price_s_manusa)
        preco_s_m = col2.number_input("Preço (R$)", key="p_s_man", step=10.0)
        qtd_s_m   = col3.number_input("Quant", min_value=1, value=1, key="q_s_man")
        if st.button("Adicionar Serviço", key="btn_s_man", use_container_width=True):
            if item_s_m:
                st.session_state.lista_s_manusa.append(
                    {"item": item_s_m, "preco": preco_s_m, "qtd": qtd_s_m}
                )
                st.session_state.pdf_manusa = None   # invalidate on list change
                st.rerun()

    for i, s in enumerate(st.session_state.lista_s_manusa):
        r = st.columns([0.4, 0.4, 3, 1, 0.5, 0.8])
        r[0].button("▲", key=f"usm_{i}", use_container_width=True,
                    on_click=update_list, args=("lista_s_manusa", i, "up"))
        r[1].button("▼", key=f"dsm_{i}", use_container_width=True,
                    on_click=update_list, args=("lista_s_manusa", i, "down"))
        r[2].write(s['item'])
        r[3].write(eng.format_real(s['preco']))
        r[4].write(f"x{s['qtd']}")
        r[5].button("🗑️", key=f"del_sm_{i}", use_container_width=True,
                    on_click=update_list, args=("lista_s_manusa", i, "delete"))

    # Products
    st.subheader("📦 Produtos")
    with st.container(border=True):
        col1, col2, col3 = st.columns([3, 1, 1])
        item_p_m  = col1.selectbox("Buscar Produto", [""] + p_list,
                                    key="sel_p_man", on_change=sync_price_p_manusa)
        preco_p_m = col2.number_input("Preço (R$)", key="p_p_man", step=10.0)
        qtd_p_m   = col3.number_input("Quant", min_value=1, value=1, key="q_p_man")
        if st.button("Adicionar Produto", key="btn_p_man", use_container_width=True):
            if item_p_m:
                st.session_state.lista_p_manusa.append(
                    {"item": item_p_m, "preco": preco_p_m, "qtd": qtd_p_m}
                )
                st.session_state.pdf_manusa = None   # invalidate on list change
                st.rerun()

    for i, p in enumerate(st.session_state.lista_p_manusa):
        r = st.columns([0.4, 0.4, 3, 1, 0.5, 0.8])
        r[0].button("▲", key=f"upm_{i}", use_container_width=True,
                    on_click=update_list, args=("lista_p_manusa", i, "up"))
        r[1].button("▼", key=f"dpm_{i}", use_container_width=True,
                    on_click=update_list, args=("lista_p_manusa", i, "down"))
        r[2].write(p['item'])
        r[3].write(eng.format_real(p['preco']))
        r[4].write(f"x{p['qtd']}")
        r[5].button("🗑️", key=f"del_pm_{i}", use_container_width=True,
                    on_click=update_list, args=("lista_p_manusa", i, "delete"))

    # Observations
    st.subheader("📝 Observações")
    obs_m = st.text_area("Notas Adicionais", key="obs_m")

    # PDF Generation & Clear
    st.write("---")
    c_pdf_m, c_clr_m = st.columns(2)
    with c_pdf_m:
        if st.session_state.lista_s_manusa or st.session_state.lista_p_manusa:
            if st.button("🚀 Gerar e Baixar Orçamento Manusa", key="btn_pdf_man",
                         use_container_width=True):
                doors_text = "ITENS DO PROJETO:\n" + "\n".join(final_doors_summary)
                obs_final  = f"{doors_text}\n\nOBSERVAÇÕES:\n{obs_m}"
                main_dim   = (st.session_state.lista_portas_manusa[0]['dim']
                              if st.session_state.lista_portas_manusa
                              else f"{w_m}mm x {h_m}mm")
                dados_m = {
                    "nome":      nome_m,
                    "empresa":   emp_m,
                    "email":     email_m,
                    "fone":      fone_m,
                    "cnpj":      id_m,
                    "dimensoes": main_dim,
                    "area":      "N/A",
                }
                st.session_state.pdf_manusa          = eng.gerar_pdf_orcamento(
                    dados_m,
                    st.session_state.lista_s_manusa,
                    st.session_state.lista_p_manusa,
                    obs_final,
                )
                st.session_state.pdf_manusa_filename = f"Manusa_{emp_m}.pdf"

            if st.session_state.pdf_manusa:
                st.download_button(
                    label="📥 Baixar PDF Manusa",
                    data=st.session_state.pdf_manusa,
                    file_name=st.session_state.pdf_manusa_filename,
                    mime="application/pdf",
                    use_container_width=True,
                )
        else:
            st.info("Adicione um Serviço ou Produto para Gerar o Orçamento.")

    with c_clr_m:
        st.button("🗑️ Limpar Tudo", key="btn_clr_man",
                  on_click=clear_manusa_callback, use_container_width=True)

# --- TAB 5: NOTAS FISCAIS ---
with tabs[5]:
    st.header("🧾 Notas Fiscais - NC Portas")

    # --- View mode selector: single sheet or combined search ---
    nf_sheet_options = list(db._NF_SHEETS.keys())   # ["NF-e NC", "NFS-e NC"]
    nf_view_options  = nf_sheet_options + ["🔎 Busca Combinada (NF-e + NFS-e)"]
    nf_view_key      = st.radio(
        "Selecionar Tabela",
        options=nf_view_options,
        horizontal=True,
        key="nf_view_sel",
    )
    is_combined = nf_view_key == nf_view_options[-1]

    st.write("---")

    # Search bar + CHAVE toggle (CHAVE toggle only relevant for single-sheet views)
    col_search, col_toggle = st.columns([4, 1])
    nf_query   = col_search.text_input(
        "🔍 Buscar por OS, Cliente ou Número da Nota",
        placeholder="Ex: 427  |  Mondelez  |  5011",
        key="nf_query",
    )
    show_chave = col_toggle.toggle(
        "Mostrar CHAVE", value=False, key="nf_show_chave",
        disabled=is_combined,   # CHAVE not in combined view
    )

    # CSS: center OS, TIPO, NOTA, NF-e, NFS-e, DPS columns.
    # Red text rows for STATUS = Cancelada rendered via st.dataframe styling below.
    st.markdown("""
        <style>
        [data-testid="stDataFrame"] td:first-child,
        [data-testid="stDataFrame"] th:first-child { text-align: center !important; }
        </style>
    """, unsafe_allow_html=True)

    # --- COMBINED VIEW ---
    if is_combined:
        if not nf_query.strip():
            st.info("Digite um termo para buscar em NF-e NC e NFS-e NC simultaneamente.")
        else:
            df_combined = db.search_nf_combined(nf_query)
            if df_combined.empty:
                st.warning("Nenhum resultado encontrado.")
            else:
                df_display = df_combined.copy()

                # Format currency columns
                for col in ["TOTAL OS", "VALOR"]:
                    if col in df_display.columns:
                        df_display[col] = df_display[col].apply(
                            lambda v: eng.format_real(float(v)) if v == v and v is not None else ""
                        )

                # Red text for Cancelada rows via Pandas Styler
                def _style_cancelada(row):
                    color = "color: red;" if str(row.get("STATUS", "")).strip() == "Cancelada" else ""
                    return [color] * len(row)

                styled = df_display.style.apply(_style_cancelada, axis=1)

                col_cfg_comb = {
                    "OS":   st.column_config.TextColumn("OS",   width="small"),
                    "TIPO": st.column_config.TextColumn("TIPO", width="small"),
                    "NOTA": st.column_config.TextColumn("NOTA", width="small"),
                }

                row_height = 35; header_h = 38; max_height = 900
                tbl_height = min(len(df_display) * row_height + header_h, max_height)
                st.caption(f"{len(df_display)} registro(s) encontrado(s)")
                st.dataframe(
                    styled,
                    hide_index=True,
                    use_container_width=True,
                    height=tbl_height,
                    column_config=col_cfg_comb,
                )

    # --- SINGLE SHEET VIEW ---
    else:
        nf_sheet_key  = nf_view_key
        df_nf         = db.search_nf_sheet(nf_sheet_key, nf_query)

        if df_nf.empty:
            st.warning("Nenhum resultado encontrado." if nf_query else "Nenhum dado disponível.")
        else:
            # Hide CHAVE unless toggled on
            all_cols     = df_nf.columns.tolist()
            hide_cols    = ["CHAVE"] if not show_chave else []
            display_cols = [c for c in all_cols if c not in hide_cols]
            df_display   = df_nf[display_cols].copy()

            # Currency formatting
            currency_cols = db._NF_SHEETS[nf_sheet_key]["currency_cols"]
            for col in currency_cols:
                if col in df_display.columns:
                    df_display[col] = df_display[col].apply(
                        lambda v: eng.format_real(float(v)) if v == v and v is not None else ""
                    )

            # Red text for Cancelada rows
            def _style_cancelada_single(row):
                color = "color: red;" if str(row.get("STATUS", "")).strip() == "Cancelada" else ""
                return [color] * len(row)

            styled = df_display.style.apply(_style_cancelada_single, axis=1)

            # column_config: center-aligned columns, currency as text
            _center_cols = ["OS", "NF-e", "NFS-e", "DPS"]
            col_cfg = {}
            for col in _center_cols:
                if col in display_cols:
                    col_cfg[col] = st.column_config.TextColumn(col, width="small")
            for col in currency_cols:
                if col in display_cols:
                    col_cfg[col] = st.column_config.TextColumn(col)

            row_height = 35; header_h = 38; max_height = 900
            tbl_height = min(len(df_display) * row_height + header_h, max_height)
            st.caption(f"{len(df_display)} registro(s) encontrado(s)")
            st.dataframe(
                styled,
                hide_index=True,
                use_container_width=True,
                height=tbl_height,
                column_config=col_cfg,
            )
    # --- REPORT GENERATOR ---
    st.write("---")
    st.subheader("📊 Gerar Relatório Mensal")

    with st.container(border=True):
        # Month and year selectors
        MESES = {
            "Janeiro": 1, "Fevereiro": 2, "Março": 3, "Abril": 4,
            "Maio": 5, "Junho": 6, "Julho": 7, "Agosto": 8,
            "Setembro": 9, "Outubro": 10, "Novembro": 11, "Dezembro": 12,
        }
        col_r1, col_r2, col_r3 = st.columns([2, 1, 2])

        rel_mes_nome = col_r1.selectbox(
            "Mês", options=list(MESES.keys()), key="rel_mes"
        )
        rel_ano = col_r2.number_input(
            "Ano", min_value=2020, max_value=2099,
            value=2026, step=1, key="rel_ano"
        )
        rel_sheet = col_r3.selectbox(
            "Relatório", options=list(db._NF_SHEETS.keys()), key="rel_sheet"
        )

        if st.button("🚀 Gerar Relatório Excel", key="btn_rel_excel",
                     use_container_width=True):
            rel_month = MESES[rel_mes_nome]
            df_rel    = db.get_nf_by_month(rel_sheet, int(rel_ano), rel_month)

            if df_rel.empty:
                st.warning(f"Nenhuma nota encontrada em {rel_mes_nome}/{rel_ano} para {rel_sheet}.")
                st.session_state.rel_xlsx      = None
                st.session_state.rel_xlsx_name = ""
            else:
                mes_ano_label = f"{rel_mes_nome} {rel_ano}"
                if rel_sheet == "NF-e NC":
                    xlsx_bytes = eng.gerar_relatorio_nfe(df_rel, mes_ano_label)
                else:
                    xlsx_bytes = eng.gerar_relatorio_nfse(df_rel, mes_ano_label)

                safe_mes = rel_mes_nome.replace(" ", "_")
                st.session_state.rel_xlsx      = xlsx_bytes
                st.session_state.rel_xlsx_name = (
                    f"NC_Portas_-_{rel_sheet.replace(' ', '_').replace('-', '')}"
                    f"_-_Relatório_{safe_mes}.xlsx"
                )

        if st.session_state.get("rel_xlsx"):
            st.download_button(
                label="📥 Baixar Relatório Excel",
                data=st.session_state.rel_xlsx,
                file_name=st.session_state.rel_xlsx_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
