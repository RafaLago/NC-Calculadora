# python -m streamlit run "NC Calculadora\NC_Calculadora.py"

# =============================================================================
# NC CALCULADORA - MAIN UI
# =============================================================================
import streamlit as st
import os
import re
import engines as eng
import database as db

# --- GLOBAL STYLES & CONFIG ---
def apply_custom_styles():
    st.markdown("""
        <style>
        @import url('https://fonts.cdnfonts.com/css/gill-sans-nova');
        * { font-family: 'Gill Sans Nova', sans-serif; }
        .stMetric { border: 1px solid #ddd; padding: 10px; border-radius: 10px; }
        </style>
    """, unsafe_allow_html=True)

st.set_page_config(page_title="Calculadora NC Portas", layout="wide")
apply_custom_styles()

# --- LOGO & TITLE ---
logo_path = os.path.join(os.path.dirname(__file__), "logo.png")
if os.path.exists(logo_path):
    st.image(logo_path, width=300)
else:
    st.warning("Logo não encontrada.")

st.title("Calculadora NC Portas")
st.write("---")

# --- SIDEBAR: CONSULTA DE PREÇOS & TAXAS ---
with st.sidebar:
    st.header("🚚 Deslocamento")
    km = st.number_input("Distância KM (Ida)", value=0)
    if km > 0:
        st.metric("Taxa Deslocamento", eng.calcular_deslocamento(km)['br_valor'])
    
    st.write("---")
    st.header("🛏️ Acomodação")
    dias = st.number_input("Dias", min_value=0)
    func = st.number_input("Funcionários", min_value=0)
    if dias > 0 and func > 0:
        st.metric("Taxa Acomodação", eng.calcular_acomodacao(dias, func)['br_valor'])

    st.write("---")
    st.header("🔍 Consulta de Preços")
    s_list, p_list = db.get_unique_items()
    
    # Serviços
    sel_s = st.selectbox("Serviços", [""] + s_list)
    if sel_s:
        price_s = db.get_item_price(sel_s, "Serviços")
        st.metric("Valor", eng.format_real(price_s))
        hist_s = db.get_last_10_entries(sel_s, "Serviços")
        if hist_s is not None:
            st.dataframe(hist_s, hide_index=True, use_container_width=True)

    # Produtos
    sel_p = st.selectbox("Produtos", [""] + p_list)
    if sel_p:
        price_p = db.get_item_price(sel_p, "Produtos")
        st.metric("Valor", eng.format_real(price_p))
        hist_p = db.get_last_10_entries(sel_p, "Produtos")
        if hist_p is not None:
            st.dataframe(hist_p, hide_index=True, use_container_width=True)

# --- MAIN SCREEN: CALCULATOR & PDF GENERATOR ---

# Define the tabs with your original names
tabs = st.tabs(["Flexdoor", "Cortina de PVC", "Peça de PVC", "Orçamento NC Portas"])

# --- TAB 1: FLEXDOOR ---
with tabs[0]:
    st.subheader("Informe as Dimensões do Vão (Flexdoor)")
    col_in1, col_in2 = st.columns(2)
    largura = col_in1.number_input("Largura (mm)", value=2000, key="w_flex")
    altura = col_in2.number_input("Altura (mm)", value=2100, key="h_flex")

    res = eng.calcular_flexdoor(largura, altura)
    if res["is_valid"]:
        c1, c2 = st.columns(2)
        c1.metric("Tipo", res["tipo"])
        c1.metric("Área do Vão", f"{res['area_m2']:.2f} m²")
        c2.metric("Valor Total", res['br_price'])
        c2.metric("Preço do m²", res['br_m2'])
        st.success(f"Porta Flexdoor {res['tipo']} nas Dimensões {largura}mm x {altura}mm.")
    else:
        st.error(f"Atenção: {res['problema']} (Máx: {res['max_dim']}mm)")

# --- TAB 2: CORTINA PVC ---
with tabs[1]:
    st.subheader("Informe as Dimensões do Vão (Cortina)")
    col_in1, col_in2 = st.columns(2)
    largura = col_in1.number_input("Largura (mm)", value=2000, key="w_pvc")
    altura = col_in2.number_input("Altura (mm)", value=2100, key="h_pvc")

    res = eng.calcular_pvc(largura, altura)
    c1, c2 = st.columns(2)
    c1.metric("Área do Vão", f"{res['area_m2']:.2f} m²")
    c2.metric("Valor Total", res['br_price'])
    if res["espessura_especial"]:
        st.warning("Atenção: Necessária espessura especial para altura > 4m")

# --- TAB 3: PEÇA PVC ---
with tabs[2]:
    st.subheader("Informe as Dimensões do Vão (Peça)")
    col_in1, col_in2 = st.columns(2)
    largura = col_in1.number_input("Largura (mm)", value=2000, key="w_peca")
    altura = col_in2.number_input("Altura (mm)", value=2100, key="h_peca")

    res = eng.calcular_peca_pvc(largura, altura)
    if res["is_valid"]:
        c1, c2 = st.columns(2)
        c1.metric("Tipo", res["tipo"])
        c1.metric("Área do Vão", f"{res['area_m2']:.2f} m²")
        c2.metric("Valor Total", res['br_price'])
        c2.metric("Preço do m²", res['br_m2'])
    else:
        st.error(f"Erro: {res['problema']}")

# --- TAB 4: ORÇAMENTO NC PORTAS ---
with tabs[3]:
    st.header("📝 Gerador de Base de Orçamento")
    
    # --- SESSION STATE INITIALIZATION ---
    if 'lista_servicos' not in st.session_state: st.session_state.lista_servicos = []
    if 'lista_produtos' not in st.session_state: st.session_state.lista_produtos = []

    # --- CALLBACKS FOR AUTOMATIC PRICE SYNC ---
    def sync_price_s():
        item = st.session_state.sel_s_orc
        if item:
            st.session_state.p_s_orc = db.get_item_price(item, "Serviços")

    def sync_price_p():
        item = st.session_state.sel_p_orc
        if item:
            st.session_state.p_p_orc = db.get_item_price(item, "Produtos")

    # --- HELPER FUNCTIONS ---
    def update_list(lista_name, index, action):
        lista = st.session_state[lista_name]
        if action == "up" and index > 0: 
            lista[index], lista[index-1] = lista[index-1], lista[index]
        elif action == "down" and index < len(lista)-1: 
            lista[index], lista[index+1] = lista[index+1], lista[index]
        elif action == "delete": 
            lista.pop(index)
        st.rerun()

    def clear_form_callback():
        st.session_state.lista_servicos = []
        st.session_state.lista_produtos = []
        keys_to_reset = [
            "nome_in", "empresa_in", "fone_in", "email_in", "id_in", "obs_in",
            "sel_s_orc", "sel_p_orc", "p_s_orc", "p_p_orc", "q_s_orc", "q_p_orc"
        ]
        for key in keys_to_reset:
            if key in st.session_state:
                if "sel_" in key: st.session_state[key] = ""
                elif "q_" in key: st.session_state[key] = 1
                elif "p_" in key: st.session_state[key] = 0.0
                else: st.session_state[key] = ""

    # --- DIMENSIONS FOR PDF ---
    st.subheader("Dimensões para o PDF")
    col_in1, col_in2 = st.columns(2)
    w_orc = col_in1.number_input("Largura (mm)", value=2000, key="w_orc")
    h_orc = col_in2.number_input("Altura (mm)", value=2100, key="h_orc")

    # --- CLIENT DATA ---
    with st.expander("Dados do Cliente", expanded=True):
        c1, c2 = st.columns(2)
        nome = c1.text_input("Nome", key="nome_in")
        empresa = c2.text_input("Empresa", key="empresa_in")
        fone = c1.text_input("Telefone", key="fone_in")
        email = c2.text_input("Email", key="email_in")
        id_doc = st.text_input("CPF/CNPJ", key="id_in")
        obs = st.text_area("Observações", key="obs_in")

    # --- SECTION: SERVIÇOS ---
    st.subheader("🛠️ Serviços")
    with st.container(border=True):
        col1, col2, col3 = st.columns([3, 1, 1])
        item_s = col1.selectbox("Buscar Serviço", [""] + s_list, key="sel_s_orc", on_change=sync_price_s)
        preco_s = col2.number_input("Preço (R$)", key="p_s_orc", step=10.0)
        qtd_s = col3.number_input("Quant", min_value=1, step=1, key="q_s_orc")
        
        if st.button("Adicionar Serviço", use_container_width=True):
            if item_s:
                st.session_state.lista_servicos.append({
                    "item": item_s, "preco": preco_s, "qtd": qtd_s, "total": preco_s * qtd_s
                })
                st.rerun()

    for i, s in enumerate(st.session_state.lista_servicos):
        r = st.columns([0.4, 0.4, 3, 1, 0.5, 0.8])
        if r[0].button("▲", key=f"us_{i}"): update_list("lista_servicos", i, "up")
        if r[1].button("▼", key=f"ds_{i}"): update_list("lista_servicos", i, "down")
        r[2].write(s['item'])
        r[3].write(eng.format_real(s['preco']))
        r[4].write(f"x{s['qtd']}")
        if r[5].button("🗑️", key=f"del_s_{i}"): update_list("lista_servicos", i, "delete")

    # --- SECTION: PRODUTOS ---
    st.subheader("📦 Produtos")
    with st.container(border=True):
        col1, col2, col3 = st.columns([3, 1, 1])
        item_p = col1.selectbox("Buscar Produto", [""] + p_list, key="sel_p_orc") # on_change=sync_price_p added below
        
        # Adding on_change manually to ensure price syncs
        if st.session_state.sel_p_orc: sync_price_p()
            
        preco_p = col2.number_input("Preço (R$)", key="p_p_orc", step=10.0)
        qtd_p = col3.number_input("Quant", min_value=1, step=1, key="q_p_orc")
        
        if st.button("Adicionar Produto", use_container_width=True):
            if item_p:
                st.session_state.lista_produtos.append({
                    "item": item_p, "preco": preco_p, "qtd": qtd_p, "total": preco_p * qtd_p
                })
                st.rerun()

    for i, p in enumerate(st.session_state.lista_produtos):
        r = st.columns([0.4, 0.4, 3, 1, 0.5, 0.8])
        if r[0].button("▲", key=f"up_{i}"): update_list("lista_produtos", i, "up")
        if r[1].button("▼", key=f"dp_{i}"): update_list("lista_produtos", i, "down")
        r[2].write(p['item'])
        r[3].write(eng.format_real(p['preco']))
        r[4].write(f"x{p['qtd']}")
        if r[5].button("🗑️", key=f"del_p_{i}"): update_list("lista_produtos", i, "delete")

    # --- SECTION: ACTIONS ---
    st.write("---")
    c_pdf, c_clear = st.columns(2)
    
    with c_pdf:
        if st.session_state.lista_servicos or st.session_state.lista_produtos:
            # Step 1: User clicks to generate
            if st.button("🚀 Gerar Orçamento PDF", use_container_width=True):
                try:
                    dados_pdf = {
                        "nome": nome, "empresa": empresa, "email": email,
                        "fone": eng.format_id_or_phone(fone, "phone"),
                        "cnpj": eng.format_id_or_phone(id_doc, "tax_id"),
                        "dimensoes": f"{w_orc}mm x {h_orc}mm",
                        "area": f"{(w_orc*h_orc)/1000000:.2f} m²"
                    }

                    # Step 2: Generate PDF directly into memory
                    pdf_data = eng.gerar_pdf_orcamento(dados_pdf, st.session_state.lista_servicos, st.session_state.lista_produtos, obs)
                    
                    # Step 3: Offer the download immediately
                    st.success("Orçamento gerado com sucesso!")
                    st.download_button(
                        label="📥 Clique aqui para Baixar", 
                        data=pdf_data, 
                        file_name=f"Orcamento_{empresa if empresa else 'NC'}.pdf",
                        mime="application/pdf",
                        use_container_width=True
                    )
                except Exception as e:
                    st.error(f"Erro ao gerar PDF: {e}")
        else:
            st.info("Adicione itens para habilitar a geração do PDF.")

    with c_clear:
        st.button("🗑️ Limpar Tudo", on_click=clear_form_callback, use_container_width=True)