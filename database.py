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

# --- SECTION: NOTAS FISCAIS LOADERS ---

# Sheet metadata: name, invoice column, date column, currency columns
_NF_SHEETS = {
    "NF-e NC": {
        "sheet":         "NF-e NC",
        "inv_col":       "NF-e",
        "date_col":      "EMISSÃO",
        "currency_cols": ["TOTAL OS", "VALOR NF-e"],
    },
    "NFS-e NC": {
        "sheet":         "NFS-e NC",
        "inv_col":       "NFS-e",
        "date_col":      "DATA",
        "currency_cols": ["TOTAL OS", "VALOR NFS-e", "ISS", "INSS", "VALOR LIQUIDO"],
    },
}

def get_nf_sheet(sheet_key):
    """Returns a cleaned DataFrame for the given invoice sheet key.
    Date column is formatted as DD/MM/YYYY string.
    Returns empty DataFrame on any failure."""
    meta = _NF_SHEETS.get(sheet_key)
    if not meta:
        logger.error("get_nf_sheet: unknown sheet_key '%s'", sheet_key)
        return pd.DataFrame()
    try:
        df = pd.read_excel(FILE, sheet_name=meta["sheet"])
        # Format date column to Brazilian standard
        date_col = meta["date_col"]
        if date_col in df.columns:
            df[date_col] = pd.to_datetime(df[date_col], errors="coerce").dt.strftime("%d/%m/%Y")
        # Ensure OS is string for consistent search matching
        df["OS"] = df["OS"].astype(str).str.strip()
        return df
    except (FileNotFoundError, ValueError, KeyError) as e:
        logger.error("get_nf_sheet failed for '%s': %s", sheet_key, e)
        return pd.DataFrame()

def search_nf_sheet(sheet_key, query):
    """Filters the invoice sheet by query string.
    Matches against OS, CLIENTE, and the invoice number column (NF-e or NFS-e).
    Search is case-insensitive and partial.
    Returns filtered DataFrame, or full DataFrame if query is empty."""
    df = get_nf_sheet(sheet_key)
    if df.empty or not query.strip():
        return df
    q           = query.strip().lower()
    inv_col     = _NF_SHEETS[sheet_key]["inv_col"]
    mask_os     = df["OS"].str.lower().str.contains(q, na=False)
    mask_client = df["CLIENTE"].astype(str).str.lower().str.contains(q, na=False)
    mask_inv    = df[inv_col].astype(str).str.lower().str.contains(q, na=False)
    return df[mask_os | mask_client | mask_inv]

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

def get_nf_by_month(sheet_key, year, month):
    """Returns a DataFrame filtered to a specific year/month for report generation.
    Date column is kept as datetime for Excel writing (not formatted to string).
    Returns empty DataFrame on any failure."""
    meta = _NF_SHEETS.get(sheet_key)
    if not meta:
        logger.error("get_nf_by_month: unknown sheet_key '%s'", sheet_key)
        return pd.DataFrame()
    try:
        df       = pd.read_excel(FILE, sheet_name=meta["sheet"])
        date_col = meta["date_col"]
        df[date_col] = pd.to_datetime(df[date_col], errors="coerce")
        df["OS"]     = df["OS"].astype(str).str.strip()
        mask = (df[date_col].dt.year == year) & (df[date_col].dt.month == month)
        return df[mask].copy()
    except (FileNotFoundError, ValueError, KeyError) as e:
        logger.error("get_nf_by_month failed for '%s' %d/%d: %s", sheet_key, month, year, e)
        return pd.DataFrame()
