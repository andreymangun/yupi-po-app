import streamlit as st

def is_superuser():
    """Mengecek apakah user saat ini memiliki role superuser."""
    user = st.session_state.get("current_user", {})
    return user.get("role") == "superuser"

def require_superuser():
    """Menghentikan eksekusi halaman jika user bukan superuser."""
    if not is_superuser():
        st.error("🔒 Akses Ditolak: Fitur ini hanya untuk level Manager/Superuser.")
        st.stop()