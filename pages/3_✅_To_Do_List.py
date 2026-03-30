import streamlit as st
import pandas as pd
from datetime import datetime

# 1. KONFIGURASI HARUS PALING ATAS
st.set_page_config(page_title="Task Manager", page_icon="✅", layout="wide", initial_sidebar_state="collapsed")

# 2. INISIALISASI
if "logged_in" not in st.session_state: st.session_state["logged_in"] = False
if "current_user" not in st.session_state: st.session_state["current_user"] = {"name": "User", "role": "admin"}
if "global_chat_log" not in st.session_state: st.session_state["global_chat_log"] = [{"role": "assistant", "content": "Halo! Ada yang bisa saya bantu?"}]

if "kanban_tasks" not in st.session_state:
    st.session_state["kanban_tasks"] = [
        {"id": 1, "waktu": "2026-03-29 08:00", "tugas": "Follow up Vendor Yupi", "tipe": "General", "ref": "-", "prioritas": "High", "status": "To Do"},
        {"id": 2, "waktu": "2026-03-29 08:30", "tugas": "Cetak Delivery Note", "tipe": "Cetak DN", "ref": "DN-2024-001", "prioritas": "Medium", "status": "In Progress"}
    ]

from utils.auth import require_login
from utils.theme import inject_css
from utils.topbar import render_topbar

inject_css()
require_login()
render_topbar()

# KOTAK main-glass-frame SUDAH DIHAPUS DARI SINI

st.markdown("<h2 style='color: #1E293B; margin-top: 0px;'>✅ Task Manager (Kanban)</h2>", unsafe_allow_html=True)

def add_task(tugas, prio, tipe, ref):
    new_id = len(st.session_state["kanban_tasks"]) + 1
    st.session_state["kanban_tasks"].append({
        "id": new_id, "waktu": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "tugas": tugas, "tipe": tipe, "ref": ref, "prioritas": prio, "status": "To Do"
    })

def update_task_status(task_id, new_status):
    for t in st.session_state["kanban_tasks"]:
        if t["id"] == task_id: t["status"] = new_status; break

c_add, c_filter = st.columns([2, 1])

with c_add:
    with st.expander("➕ Tambah Tugas Baru (Isi Detail)", expanded=False):
        with st.form("add_task_form", clear_on_submit=True):
            t = st.text_area("Deskripsi Tugas", placeholder="Ketik tugas di sini... (Enter untuk baris baru)")
            c1, c2, c3 = st.columns(3)
            with c1: prio = st.selectbox("Prioritas", ["High", "Medium", "Low"])
            with c2: tipe = st.selectbox("Tipe Operasional", ["General", "Cetak PO", "Cetak DN", "Sourcing"])
            with c3: ref = st.text_input("No. Referensi (Jika Ada)", placeholder="Contoh: PO-001")
            
            if st.form_submit_button("Simpan Tugas (Klik)"):
                if t:
                    add_task(t, prio, tipe, ref); st.success("Tugas berhasil ditambahkan!"); st.rerun()

with c_filter:
    st.markdown("""
    <div style="background: rgba(255,255,255,0.7); backdrop-filter: blur(10px); padding: 15px; border-radius: 15px; border: 1px solid rgba(0,0,0,0.1);">
        <b style="color: #1E293B; margin-bottom:5px; display:block;">🔍 Filter Tampilan:</b>
    """, unsafe_allow_html=True)
    view_mode = st.radio("Pilih Mode", ["Kanban Board", "Tabel Lengkap (Detail)"], label_visibility="collapsed")
    st.markdown("</div>", unsafe_allow_html=True)

st.markdown("<hr style='margin: 15px 0; border-color:rgba(0,0,0,0.1);'>", unsafe_allow_html=True)

tasks = st.session_state["kanban_tasks"]

if view_mode == "Kanban Board":
    todo_list = [t for t in tasks if t.get("status") == "To Do"]
    prog_list = [t for t in tasks if t.get("status") == "In Progress"]
    done_list = [t for t in tasks if t.get("status") == "Done"]

    col_todo, col_prog, col_done = st.columns(3)
    s_kanban = "padding:12px; border-radius:12px; background: rgba(255,255,255,0.8); border: 1px solid rgba(0,0,0,0.1); text-align:center; box-shadow: 0 2px 5px rgba(0,0,0,0.05); margin-bottom: 10px;"

    with col_todo:
        st.markdown(f"<div style='{s_kanban}'><b style='color:#1E293B;'>TO DO ({len(todo_list)})</b></div>", unsafe_allow_html=True)
        for t in todo_list:
            with st.container(border=True):
                st.markdown(f"<b style='color:#8B0000;'>{t['tipe']}</b>: {t['tugas']}<br><small style='color:gray;'>Ref: {t['ref']} | Prio: {t['prioritas']}</small>", unsafe_allow_html=True)
                if st.button("Mulai ▶", key=f"start_{t['id']}", use_container_width=True):
                    update_task_status(t['id'], "In Progress"); st.rerun()

    with col_prog:
        st.markdown(f"<div style='{s_kanban}'><b style='color:#1E293B;'>IN PROGRESS ({len(prog_list)})</b></div>", unsafe_allow_html=True)
        for t in prog_list:
            with st.container(border=True):
                st.markdown(f"<b style='color:#8B0000;'>{t['tipe']}</b>: {t['tugas']}<br><small style='color:gray;'>Ref: {t['ref']} | Prio: {t['prioritas']}</small>", unsafe_allow_html=True)
                c_prev, c_next = st.columns(2)
                with c_prev:
                    if st.button("Batal", key=f"back_{t['id']}", use_container_width=True):
                        update_task_status(t['id'], "To Do"); st.rerun()
                with c_next:
                    if st.button("Selesai ✔", key=f"done_{t['id']}", use_container_width=True):
                        update_task_status(t['id'], "Done"); st.rerun()

    with col_done:
        st.markdown(f"<div style='{s_kanban}'><b style='color:#1E293B;'>DONE ({len(done_list)})</b></div>", unsafe_allow_html=True)
        for t in done_list:
            with st.container(border=True):
                st.markdown(f"<span style='text-decoration: line-through; color:gray;'>{t['tugas']}</span>", unsafe_allow_html=True)
                if st.button("Revisi", key=f"rev_{t['id']}", use_container_width=True):
                    update_task_status(t['id'], "In Progress"); st.rerun()

elif view_mode == "Tabel Lengkap (Detail)":
    st.markdown("<b style='color:#1E293B;'>Daftar Seluruh Tugas Operasional Terperinci</b>", unsafe_allow_html=True)
    if tasks:
        df_tasks = pd.DataFrame(tasks)
        st.dataframe(df_tasks[["id", "waktu", "tipe", "ref", "tugas", "prioritas", "status"]], use_container_width=True, hide_index=True, height=400)
    else:
        st.info("Belum ada tugas operasional.")