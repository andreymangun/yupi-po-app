import streamlit as st
from utils.auth import logout

@st.dialog("✏️ Edit Profile")
def modal_edit_profile():
    user = st.session_state.get("current_user", {})
    with st.form("form_edit_profile"):
        new_name = st.text_input("Nama Lengkap", value=user.get("name", ""))
        if st.form_submit_button("Simpan Perubahan", type="primary"):
            st.session_state.current_user["name"] = new_name; st.success("Tersimpan!"); st.rerun()

@st.dialog("🔑 Ganti Password")
def modal_ganti_password():
    with st.form("form_ganti_pw"):
        st.text_input("Password Lama", type="password")
        p1 = st.text_input("Password Baru", type="password")
        p2 = st.text_input("Konfirmasi Password", type="password")
        if st.form_submit_button("Update Password", type="primary"): st.success("Diubah!")

def render_topbar():
    st.markdown("""<style>[data-testid="collapsedControl"] { display: none !important; }</style>""", unsafe_allow_html=True)
    
    # 3 Kolom: Tombol Home (Kiri), Kosong (Tengah), Akun (Kanan)
    c_home, c_empty, c_menu = st.columns([1.5, 7, 1.5])
    
    with c_home:
        if st.button("🏠 Dashboard", use_container_width=True):
            st.switch_page("app.py")
            
    with c_menu:
        user = st.session_state.get("current_user", {})
        user_name = user.get("name", "Akun")
        with st.popover(f"👤 {user_name}", use_container_width=True):
            st.markdown(f"**{user_name}**\n\n<small>{user.get('role', 'Pegawai').upper()}</small>", unsafe_allow_html=True)
            st.divider()
            if st.button("✏️ Edit Profile", use_container_width=True): modal_edit_profile()
            if st.button("🔑 Ganti Password", use_container_width=True): modal_ganti_password()
            st.divider()
            if st.button("🚪 Logout", use_container_width=True, type="primary"): logout(); st.rerun()