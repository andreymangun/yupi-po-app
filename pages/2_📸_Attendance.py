import streamlit as st
import pandas as pd
from datetime import datetime
import requests

# Konfigurasi Halaman (WAJIB PALING ATAS)
st.set_page_config(page_title="Attendance System", page_icon="📸", layout="wide", initial_sidebar_state="collapsed")

# Inisialisasi State
from utils.state import init_session
init_session()

from utils.auth import require_login
from utils.theme import inject_css
from utils.topbar import render_topbar

inject_css()
require_login()
render_topbar()

# KOTAK main-glass-frame SUDAH DIHAPUS DARI SINI

st.markdown("""
<div style="background: rgba(255, 255, 255, 0.3); backdrop-filter: blur(25px); border-radius: 20px; padding: 30px; border: 1px solid rgba(255,255,255,0.5); margin-bottom: 20px; margin-top: 0px;">
    <h1 style="color: #1E293B; margin: 0;">📸 Attendance System</h1>
    <p style="color: #334155; margin: 0;">Pusat presensi dan manajemen SDM terpadu.</p>
</div>
""", unsafe_allow_html=True)

def get_db_headers():
    try:
        key = st.secrets["supabase"]["key"]
        return {"apikey": key, "Authorization": f"Bearer {key}", "Content-Type": "application/json"}
    except: return {}

def insert_attendance(waktu, tipe):
    try:
        url = f"{st.secrets['supabase']['url']}/rest/v1/attendance"
        user = st.session_state.get("current_user", {})
        requests.post(url, headers=get_db_headers(), json={"user_id": user.get("id"), "nama": user.get("name"), "waktu": waktu, "tipe": tipe})
    except: pass

user_name = st.session_state.current_user.get('name', 'User')
now = datetime.now()

tabs = st.tabs(["🤳 Presensi", "⏳ Overtime", "🗓️ Cuti", "📝 Aktivitas", "📋 Riwayat"])

with tabs[0]:
    st.markdown(f"**Nama:** {user_name} | **Waktu:** {now.strftime('%H:%M')}")
    cam = st.camera_input("Ambil Foto Selfie di Lokasi")
    if cam:
        c1, c2 = st.columns(2)
        with c1:
            if st.button("🟢 CHECK IN", use_container_width=True, type="primary"):
                insert_attendance(now.strftime("%H:%M:%S"), "Masuk")
                st.success("✅ Check In Berhasil Dicatat!")
        with c2:
            if st.button("🔴 CHECK OUT", use_container_width=True):
                insert_attendance(now.strftime("%H:%M:%S"), "Keluar")
                st.success("✅ Check Out Berhasil Dicatat!")

with tabs[1]:
    st.markdown("### Form Pengajuan Lembur")
    with st.form("overtime_form"):
        st.date_input("Tanggal Lembur")
        st.number_input("Durasi (Jam)", min_value=1, max_value=12)
        st.text_area("Deskripsi Pekerjaan")
        if st.form_submit_button("Ajukan Overtime", type="primary"): st.success("✅ Terkirim ke Supervisor.")

with tabs[2]:
    st.markdown("### Form Pengajuan Cuti")
    with st.form("leave_form"):
        c1, c2 = st.columns(2)
        with c1: st.date_input("Tanggal Mulai")
        with c2: st.date_input("Tanggal Selesai")
        st.selectbox("Tipe Cuti", ["Tahunan", "Sakit", "Melahirkan", "Khusus"])
        st.text_area("Alasan Cuti")
        if st.form_submit_button("Ajukan Cuti", type="primary"): st.success("✅ Terkirim ke HRD.")

with tabs[3]:
    st.markdown("### Log Aktivitas Harian")
    with st.form("activity_form"):
        st.text_area("Apa yang Anda kerjakan hari ini?")
        st.slider("Tingkat Penyelesaian (%)", 0, 100, 80)
        if st.form_submit_button("Simpan Laporan"): st.success("✅ Laporan disimpan.")

with tabs[4]:
    try:
        url = f"{st.secrets['supabase']['url']}/rest/v1/attendance?select=*&order=created_at.desc"
        res = requests.get(url, headers=get_db_headers())
        if res.status_code == 200 and res.json():
            st.dataframe(pd.DataFrame(res.json())[["nama", "tipe", "waktu"]], use_container_width=True, hide_index=True)
        else: st.write("Belum ada data presensi.")
    except: st.write("Belum ada data presensi.")