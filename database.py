# =============================================================================
# DATABASE - EXCEL HANDLER
# =============================================================================
import pandas as pd
import os

BASE = os.path.dirname(__file__)
FILE = os.path.join(BASE, "Serviços & Produtos NC.xlsx")

def get_unique_items():
    try:
        df_s = pd.read_excel(FILE, sheet_name="Serviços")
        df_p = pd.read_excel(FILE, sheet_name="Produtos")
        return sorted(df_s["SERVIÇO"].dropna().unique().tolist()), sorted(df_p["PRODUTO"].dropna().unique().tolist())
    except: return [], []

def get_item_price(item, cat):
    try:
        sheet = "Serviços" if cat == "Serviços" else "Produtos"
        col = "SERVIÇO" if cat == "Serviços" else "PRODUTO"
        df = pd.read_excel(FILE, sheet_name=sheet)
        filt = df[df[col] == item].copy().sort_values(by="OS", ascending=False)
        modes = filt["V. UNIT"].mode()
        return float(modes[0]) if not modes.empty else float(filt["V. UNIT"].iloc[0])
    except: return 0.0

def get_last_10_entries(item, cat):
    try:
        from engines import format_real
        sheet = "Serviços" if cat == "Serviços" else "Produtos"
        col = "SERVIÇO" if cat == "Serviços" else "PRODUTO"
        df = pd.read_excel(FILE, sheet_name=sheet)
        filt = df[df[col] == item].copy().sort_values(by="OS", ascending=False).head(10)
        filt["V. UNIT"] = filt["V. UNIT"].map(format_real)
        filt["OS"] = filt["OS"].astype(str)
        return filt[["CLIENTE", "V. UNIT", "OS"]]
    except: return None