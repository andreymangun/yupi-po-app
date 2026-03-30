import pandas as pd


def _first_scalar(val):
    if isinstance(val, pd.Series):
        return val.iloc[0] if not val.empty else None
    if isinstance(val, (list, tuple)):
        return val[0] if len(val) > 0 else None
    return val


def clean_val(val):
    val = _first_scalar(val)
    if val is None:
        return ""
    if pd.isna(val) or str(val).lower() == "nan":
        return ""
    return str(val).encode("latin-1", "replace").decode("latin-1").strip()


def safe_float(val) -> float:
    val = _first_scalar(val)
    if val is None:
        return 0.0
    if pd.isna(val) or str(val).lower() == "nan":
        return 0.0
    try:
        if isinstance(val, str):
            val = val.replace(",", "").replace('"', "").strip()
        return float(val) if val != "" else 0.0
    except Exception:
        return 0.0


def format_currency(val) -> str:
    val = _first_scalar(val)
    try:
        return f"{float(val):,.2f}"
    except Exception:
        return "0.00"


def format_qty(val) -> str:
    val = _first_scalar(val)
    try:
        value = float(val)
        return f"{value:,.0f}" if value % 1 == 0 else f"{value:,.2f}"
    except Exception:
        return "0"