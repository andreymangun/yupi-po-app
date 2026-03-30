import streamlit as st
import requests

def get_supabase_headers():
    """Menghasilkan headers yang dibutuhkan untuk API Supabase"""
    key = st.secrets["supabase"]["key"]
    return {
        "apikey": key,
        "Authorization": f"Bearer {key}",
        "Content-Type": "application/json"
    }

def login(email, password):
    url = st.secrets["supabase"]["url"]
    
    # 1. Autentikasi User (Gotrue REST API)
    auth_url = f"{url}/auth/v1/token?grant_type=password"
    auth_payload = {"email": email, "password": password}
    
    try:
        response = requests.post(auth_url, headers=get_supabase_headers(), json=auth_payload)
        
        if response.status_code != 200:
            # Gagal login
            st.error("Login Gagal: Pastikan Email dan Password sudah benar.")
            return False
            
        auth_data = response.json()
        user_id = auth_data["user"]["id"]
        
        # 2. Ambil Profil User dari Database (PostgREST API)
        # Pastikan Anda sudah membuat tabel 'profiles' di Supabase
        db_url = f"{url}/rest/v1/profiles?id=eq.{user_id}&select=*"
        db_response = requests.get(db_url, headers=get_supabase_headers())
        
        if db_response.status_code == 200 and len(db_response.json()) > 0:
            profile = db_response.json()[0]
            
            # 3. Simpan state login ke Session Streamlit
            st.session_state["logged_in"] = True
            st.session_state["current_user"] = {
                "id": user_id,
                "email": email,
                "name": profile.get("full_name", "Karyawan"),
                "role": profile.get("role", "user"), # Peran untuk RBAC (user/superuser)
                "lokasi": profile.get("lokasi", "HQ Jakarta")
            }
            # Token ini berguna jika nanti kita butuh write-back data ke Supabase
            st.session_state["access_token"] = auth_data.get("access_token")
            return True
        else:
            st.error("Profil tidak ditemukan di database. Pastikan Trigger SQL sudah berjalan.")
            return False
            
    except requests.exceptions.RequestException as e:
        st.error(f"Koneksi ke server gagal: {e}")
        return False

def logout():
    st.session_state.clear()
    st.rerun()

def require_login():
    """Fungsi penjaga halaman operasional."""
    if not st.session_state.get("logged_in", False):
        st.warning("Anda harus login terlebih dahulu.")
        st.switch_page("app.py")
        st.stop()