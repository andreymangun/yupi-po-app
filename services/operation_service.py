import io
import zipfile
import inspect
from datetime import datetime
from io import StringIO

import pandas as pd
import requests
import streamlit as st

from utils.formatters import safe_float, clean_val
from pdf_engine.po_pdf import generate_po_pdf
from pdf_engine.dn_pdf import generate_dn_pdf


DEFAULT_DISPLAY_COLUMNS = [
    "PO YUPI",
    "PO SEMENTARA",
    "Vendor Name",
    "Item Yupi",
    "Item name",
    "Spec",
    "Ord. Q'ty",
    "Unit",
    "PURCHASE PRICE",
]

POSSIBLE_QTY_COLS = ["Ord. Q'ty", "Ord. Q'ty__1", "ORD QTY", "Qty", "Quantity"]
POSSIBLE_PRICE_COLS = ["PURCHASE PRICE", "PURCHASE PRICE__1", "Price", "Unit Price"]


# =========================================================
# DATA PREP
# =========================================================
def _normalize_header(value) -> str:
    return str(value).replace("\n", " ").replace("\r", " ").strip()


def _make_unique_headers(headers: list[str]) -> list[str]:
    seen = {}
    result = []

    for h in headers:
        if h in seen:
            seen[h] += 1
            result.append(f"{h}__{seen[h]}")
        else:
            seen[h] = 0
            result.append(h)

    return result


def _find_header_row(df_full: pd.DataFrame) -> int:
    for idx in range(min(15, len(df_full))):
        row = [_normalize_header(x) for x in df_full.iloc[idx].tolist()]
        if "PO YUPI" in row and "Vendor Name" in row:
            return idx
    raise ValueError("Header CSV tidak ditemukan. Pastikan format file sesuai.")


def _first_existing_column(df: pd.DataFrame, candidates: list[str]) -> str | None:
    for col in candidates:
        if col in df.columns:
            return col
    return None


def _prepare_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    qty_col = _first_existing_column(df, POSSIBLE_QTY_COLS)
    price_col = _first_existing_column(df, POSSIBLE_PRICE_COLS)

    if qty_col and qty_col != "Ord. Q'ty":
        df["Ord. Q'ty"] = df[qty_col]
    if price_col and price_col != "PURCHASE PRICE":
        df["PURCHASE PRICE"] = df[price_col]

    if "Ord. Q'ty" in df.columns:
        df["Ord. Q'ty"] = df["Ord. Q'ty"].apply(safe_float)

    if "PURCHASE PRICE" in df.columns:
        df["PURCHASE PRICE"] = df["PURCHASE PRICE"].apply(safe_float)

    for col in ["PO YUPI", "PO SEMENTARA", "Vendor Name", "Item Yupi", "Item name", "Spec", "Unit"]:
        if col in df.columns:
            df[col] = df[col].apply(clean_val)

    if "PO YUPI" in df.columns:
        df = df[df["PO YUPI"].astype(str).str.strip() != ""]
    if "Vendor Name" in df.columns:
        df = df[df["Vendor Name"].astype(str).str.strip() != ""]

    return df.reset_index(drop=True)


@st.cache_data(show_spinner=False, ttl=300)
def fetch_operation_dataframe(csv_url: str) -> pd.DataFrame:
    response = requests.get(csv_url, timeout=20)
    response.raise_for_status()

    df_full = pd.read_csv(StringIO(response.text), header=None, low_memory=False)
    header_row = _find_header_row(df_full)

    raw_headers = [_normalize_header(x) for x in df_full.iloc[header_row].tolist()]
    final_headers = _make_unique_headers(raw_headers)

    df = df_full.iloc[header_row + 1:].copy()
    df.columns = final_headers

    required_cols = ["PO YUPI", "Vendor Name"]
    for col in required_cols:
        if col not in df.columns:
            raise ValueError(f"Kolom wajib '{col}' tidak ditemukan di CSV.")

    return _prepare_dataframe(df)


# =========================================================
# SESSION STATE
# =========================================================
def init_operation_state():
    defaults = {
        "operation_df": None,
        "operation_csv_url": "",
        "operation_loaded": False,
        "operation_step": 1,
        "operation_context_key": "",
        "dn_counter": 1,
        "dn_last_date": "",
        "dn_vendor": "",
        "dn_no_pol": "",
        "dn_remarks": "",
        "generated_po_bytes": None,
        "generated_po_filename": "",
        "generated_dn_bytes": None,
        "generated_dn_filename": "",
        "generated_bulk_po_zip": None,
        "generated_bulk_po_filename": "",
        "generated_bulk_dn_zip": None,
        "generated_bulk_dn_filename": "",
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value


def reset_generated_files():
    st.session_state.generated_po_bytes = None
    st.session_state.generated_po_filename = ""
    st.session_state.generated_dn_bytes = None
    st.session_state.generated_dn_filename = ""
    st.session_state.generated_bulk_po_zip = None
    st.session_state.generated_bulk_po_filename = ""
    st.session_state.generated_bulk_dn_zip = None
    st.session_state.generated_bulk_dn_filename = ""


def load_operation_data(csv_url: str):
    df = fetch_operation_dataframe(csv_url)
    st.session_state.operation_df = df
    st.session_state.operation_csv_url = csv_url
    st.session_state.operation_loaded = True
    st.session_state.operation_step = 1
    reset_generated_files()
    return df


def clear_operation_data():
    st.session_state.operation_df = None
    st.session_state.operation_loaded = False
    st.session_state.operation_step = 1
    st.session_state.operation_context_key = ""
    reset_generated_files()


def update_dn_counter_daily():
    today = datetime.now().strftime("%Y-%m-%d")
    if st.session_state.get("dn_last_date") != today:
        st.session_state.dn_last_date = today
        st.session_state.dn_counter = 1


def get_next_dn_number() -> str:
    update_dn_counter_daily()
    counter = st.session_state.get("dn_counter", 1)
    today = datetime.now().strftime("%d%m%Y")
    return f"DN-{today}-{counter:03d}"


def bump_dn_counter():
    st.session_state.dn_counter = st.session_state.get("dn_counter", 1) + 1


# =========================================================
# FILTER HELPERS
# =========================================================
def get_available_po_numbers(df: pd.DataFrame) -> list[str]:
    if df is None or df.empty or "PO YUPI" not in df.columns:
        return []
    po_list = df["PO YUPI"].dropna().astype(str).str.strip()
    po_list = po_list[po_list != ""].unique().tolist()
    return sorted(po_list)


def filter_po_dataframe(df: pd.DataFrame, po_number: str) -> pd.DataFrame:
    if df is None or df.empty or not po_number:
        return pd.DataFrame()
    return df[df["PO YUPI"].astype(str).str.strip() == str(po_number).strip()].copy()


def get_available_vendors(po_df: pd.DataFrame) -> list[str]:
    if po_df is None or po_df.empty or "Vendor Name" not in po_df.columns:
        return []
    vendors = po_df["Vendor Name"].dropna().astype(str).str.strip()
    vendors = vendors[vendors != ""].unique().tolist()
    return sorted(vendors)


def filter_vendor_dataframe(po_df: pd.DataFrame, vendor_name: str) -> pd.DataFrame:
    if po_df is None or po_df.empty or not vendor_name:
        return pd.DataFrame()
    return po_df[po_df["Vendor Name"].astype(str).str.strip() == str(vendor_name).strip()].copy()


def get_display_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame()
    existing_cols = [c for c in DEFAULT_DISPLAY_COLUMNS if c in df.columns]
    return df[existing_cols].copy() if existing_cols else df.copy()


# =========================================================
# PDF ENGINE COMPAT HELPERS
# =========================================================
def _extract_pdf_bytes(result):
    """
    Mendukung beberapa bentuk return dari engine lama:
    - bytes
    - bytearray
    - io.BytesIO
    - tuple(bytes, ...)
    - tuple(BytesIO, ...)
    """
    if isinstance(result, bytes):
        return result

    if isinstance(result, bytearray):
        return bytes(result)

    if hasattr(result, "getvalue"):
        return result.getvalue()

    if isinstance(result, tuple):
        first = result[0]
        if isinstance(first, bytes):
            return first
        if isinstance(first, bytearray):
            return bytes(first)
        if hasattr(first, "getvalue"):
            return first.getvalue()

    raise ValueError(f"Format output PDF tidak dikenali: {type(result)}")


def _call_generate_po_pdf(vendor_df: pd.DataFrame, po_number: str):
    # Sesuai engine asli Anda: generate_po_pdf(po_data, po_number)
    return generate_po_pdf(vendor_df, po_number)


def _call_generate_dn_pdf(
    selected_df: pd.DataFrame,
    po_number: str,
    checklist_notes: dict,
    dn_vendor: str,
    no_pol: str,
    remarks: str,
    dn_number: str,
):
    # Masukkan checklist notes ke kolom yang memang dibaca oleh dn_pdf.py
    selected_df = selected_df.copy()

    if checklist_notes:
        selected_df["Catatan Item"] = [
            checklist_notes.get(i, "") for i in range(len(selected_df))
        ]
    else:
        if "Catatan Item" not in selected_df.columns:
            selected_df["Catatan Item"] = ""

    # Sesuai signature asli dn_pdf.py:
    # generate_dn_pdf(po_data, po_number, dn_vendor, no_pol, dn_remarks, dn_serveone_number)
    return generate_dn_pdf(
        selected_df,
        po_number,
        dn_vendor,
        no_pol,
        remarks,
        dn_number,
    )


# =========================================================
# PDF HELPERS
# =========================================================
def build_po_pdf_bytes(df_vendor: pd.DataFrame) -> bytes:
    if df_vendor is None or df_vendor.empty:
        raise ValueError("Data vendor kosong. Tidak bisa membuat PO PDF.")

    vendor_df = df_vendor.reset_index(drop=True).copy()
    po_number = clean_val(vendor_df.iloc[0].get("PO YUPI", ""))

    result = _call_generate_po_pdf(vendor_df, po_number)
    return _extract_pdf_bytes(result)


def build_dn_pdf_bytes(
    df_vendor: pd.DataFrame,
    selected_indexes: list[int],
    checklist_notes: dict | None = None,
    dn_vendor: str = "",
    no_pol: str = "",
    remarks: str = "",
    dn_number: str = "",
) -> bytes:
    if df_vendor is None or df_vendor.empty:
        raise ValueError("Data vendor kosong. Tidak bisa membuat DN PDF.")

    if not selected_indexes:
        raise ValueError("Pilih minimal 1 item untuk Delivery Note.")

    vendor_df = df_vendor.reset_index(drop=True).copy()

    max_index = len(vendor_df) - 1
    invalid_indexes = [i for i in selected_indexes if i < 0 or i > max_index]
    if invalid_indexes:
        raise ValueError(f"Index item tidak valid: {invalid_indexes}")

    selected_df = vendor_df.loc[selected_indexes].copy().reset_index(drop=True)
    po_number = clean_val(vendor_df.iloc[0].get("PO YUPI", ""))

    # notes dari editor sebelumnya berbentuk {index_asli_editor: note}
    # setelah selected_df di-reset index, kita remap sesuai urutan item terpilih
    remapped_notes = {}
    for new_idx, old_idx in enumerate(selected_indexes):
        if checklist_notes and old_idx in checklist_notes:
            remapped_notes[new_idx] = checklist_notes[old_idx]

    result = _call_generate_dn_pdf(
        selected_df=selected_df,
        po_number=po_number,
        checklist_notes=remapped_notes,
        dn_vendor=dn_vendor,
        no_pol=no_pol,
        remarks=remarks,
        dn_number=dn_number,
    )
    return _extract_pdf_bytes(result)


def build_bulk_po_zip(po_df: pd.DataFrame) -> bytes:
    if po_df is None or po_df.empty:
        raise ValueError("Data PO kosong.")

    output = io.BytesIO()

    with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as zipf:
        for vendor in get_available_vendors(po_df):
            vendor_df = filter_vendor_dataframe(po_df, vendor)
            if vendor_df.empty:
                continue

            pdf_bytes = build_po_pdf_bytes(vendor_df)
            po_number = clean_val(vendor_df.reset_index(drop=True).iloc[0].get("PO YUPI", "NO_PO"))
            filename = f"PO_{po_number}_{vendor}.pdf".replace("/", "-")
            zipf.writestr(filename, pdf_bytes)

    output.seek(0)
    return output.getvalue()


def build_bulk_dn_zip(
    po_df: pd.DataFrame,
    selected_map: dict,
    notes_map: dict,
    dn_vendor: str,
    no_pol: str,
    remarks: str,
) -> bytes:
    if po_df is None or po_df.empty:
        raise ValueError("Data PO kosong.")

    output = io.BytesIO()

    with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as zipf:
        local_counter = st.session_state.get("dn_counter", 1)
        today = datetime.now().strftime("%d%m%Y")

        for vendor in get_available_vendors(po_df):
            vendor_df = filter_vendor_dataframe(po_df, vendor)
            if vendor_df.empty:
                continue

            selected_indexes = selected_map.get(vendor, [])
            if not selected_indexes:
                continue

            dn_number = f"DN-{today}-{local_counter:03d}"
            notes = notes_map.get(vendor, {})

            pdf_bytes = build_dn_pdf_bytes(
                df_vendor=vendor_df,
                selected_indexes=selected_indexes,
                checklist_notes=notes,
                dn_vendor=dn_vendor,
                no_pol=no_pol,
                remarks=remarks,
                dn_number=dn_number,
            )

            po_number = clean_val(vendor_df.reset_index(drop=True).iloc[0].get("PO YUPI", "NO_PO"))
            filename = f"DN_{dn_number}_{po_number}_{vendor}.pdf".replace("/", "-")
            zipf.writestr(filename, pdf_bytes)
            local_counter += 1

    output.seek(0)
    return output.getvalue()