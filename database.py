# =============================================================================
# DATABASE - EXCEL HANDLER
# Reads and queries data from the NC Portas master spreadsheet.
# All public functions return safe defaults on failure and log the real error.
# =============================================================================
import os
import logging
import pandas as pd

# --- SECTION: CONFIG ---
logger  = logging.getLogger(__name__)
BASE    = os.path.dirname(__file__)
FILE    = os.path.join(BASE, "NC Portas 2026.xlsx")

# --- SECTION: DATA LOADERS ---
def get_unique_items():
    """Returns (servicos_list, produtos_list) sorted unique names from the DB.
    Returns ([], []) on any read failure."""
    try:
        df_s = pd.read_excel(FILE, sheet_name="Serviços")
        df_p = pd.read_excel(FILE, sheet_name="Produtos")
        servicos = sorted(df_s["SERVIÇO"].dropna().unique().tolist())
        produtos  = sorted(df_p["PRODUTO"].dropna().unique().tolist())
        return servicos, produtos
    except (FileNotFoundError, ValueError, KeyError) as e:
        logger.error("get_unique_items failed: %s", e)
        return [], []

def get_item_price(item, cat):
    """Returns the modal unit price for a given item and category.
    Falls back to the most recent entry price if no mode exists.
    Returns 0.0 on any failure."""
    try:
        sheet = "Serviços" if cat == "Serviços" else "Produtos"
        col   = "SERVIÇO"  if cat == "Serviços" else "PRODUTO"
        df    = pd.read_excel(FILE, sheet_name=sheet)
        filt  = df[df[col] == item].copy().sort_values(by="OS", ascending=False)
        if filt.empty:
            return 0.0
        modes = filt["V. UNIT"].mode()
        return float(modes.iloc[0]) if not modes.empty else float(filt["V. UNIT"].iloc[0])
    except (FileNotFoundError, ValueError, KeyError, IndexError) as e:
        logger.error("get_item_price failed for '%s' (%s): %s", item, cat, e)
        return 0.0

def get_last_10_entries(item, cat):
    """Returns a formatted DataFrame of the last 10 OS entries for an item.
    Columns: CLIENTE, V. UNIT (formatted), OS.
    Returns None on any failure."""
    try:
        from engines import format_real
        sheet = "Serviços" if cat == "Serviços" else "Produtos"
        col   = "SERVIÇO"  if cat == "Serviços" else "PRODUTO"
        df    = pd.read_excel(FILE, sheet_name=sheet)
        filt  = df[df[col] == item].copy().sort_values(by="OS", ascending=False).head(10)
        filt["V. UNIT"] = filt["V. UNIT"].map(format_real)
        filt["OS"]      = filt["OS"].astype(str)
        return filt[["CLIENTE", "V. UNIT", "OS"]]
    except (FileNotFoundError, ValueError, KeyError) as e:
        logger.error("get_last_10_entries failed for '%s' (%s): %s", item, cat, e)
        return None
