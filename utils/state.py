import streamlit as st

def init_session():
    """Fungsi ini WAJIB dipanggil di baris paling atas setiap halaman."""
    if "logged_in" not in st.session_state:
        st.session_state["logged_in"] = False
    if "current_user" not in st.session_state:
        st.session_state["current_user"] = {"name": "User", "role": "admin"}
    if "global_chat_log" not in st.session_state:
        st.session_state["global_chat_log"] = [{"role": "assistant", "content": "Halo! Ada yang bisa saya bantu?"}]
    if "kanban_tasks" not in st.session_state:
        # Data awal jika kosong
        st.session_state["kanban_tasks"] = []
    if "operation_df" not in st.session_state:
        st.session_state["operation_df"] = None