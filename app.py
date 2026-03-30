import streamlit as st
import pandas as pd
from datetime import datetime
import numpy as np

# 1. KONFIGURASI HARUS PALING ATAS
st.set_page_config(page_title="ServeOne ERP", page_icon="🏢", layout="wide", initial_sidebar_state="expanded")

# 2. INISIALISASI MEMORI BRUTAL
if "logged_in" not in st.session_state: st.session_state["logged_in"] = False
if "current_user" not in st.session_state: st.session_state["current_user"] = {"name": "User", "role": "admin"}
if "global_chat_log" not in st.session_state: st.session_state["global_chat_log"] = [{"role": "assistant", "content": "Halo! GPT-4o siap membantu."}]
if "kanban_tasks" not in st.session_state: st.session_state["kanban_tasks"] = []

if "op_step" not in st.session_state: st.session_state["op_step"] = 1
if "generated_po_bytes" not in st.session_state: st.session_state["generated_po_bytes"] = None
if "generated_po_filename" not in st.session_state: st.session_state["generated_po_filename"] = None
if "g_dn_bytes" not in st.session_state: st.session_state["g_dn_bytes"] = None
if "g_dn_file" not in st.session_state: st.session_state["g_dn_file"] = None
if "g_po_bytes" not in st.session_state: st.session_state["g_po_bytes"] = None
if "g_po_file" not in st.session_state: st.session_state["g_po_file"] = None
if "search_po_input" not in st.session_state: st.session_state["search_po_input"] = ""
if "copilot_history" not in st.session_state: st.session_state["copilot_history"] = [{"role": "assistant", "content": "Halo! Ketik nomor PO atau Vendor, saya akan merekap statusnya."}]
if "dn_counter_data" not in st.session_state: st.session_state["dn_counter_data"] = {}
if "operation_df" not in st.session_state: st.session_state["operation_df"] = None

from utils.auth import login
from utils.theme import inject_css
from utils.topbar import render_topbar
from utils.ai_engine import get_ai_response 

inject_css()

if not st.session_state["logged_in"]:
    st.markdown('<div style="height:10vh;"></div>', unsafe_allow_html=True)
    
    # Gunakan perbandingan kolom yang lebih adaptif (1:2:1)
    c1, c2, c3 = st.columns([1, 2, 1]) 
    with c2:
        st.markdown('<h1 class="login-logo-text">ServeOne</h1>', unsafe_allow_html=True)
        st.markdown('<p style="text-align: center; color: white; letter-spacing: 2px; font-weight: 500; margin-bottom: 2rem;">ENTERPRISE PORTAL</p>', unsafe_allow_html=True)
        with st.form("login_form"):
            u = st.text_input("Username")
            p = st.text_input("Password", type="password")
            if st.form_submit_button("Masuk Sistem", use_container_width=True, type="primary"):
                if login(u, p): 
                    st.session_state["logged_in"] = True; st.rerun()
                else: 
                    st.error("Kredensial salah.")
else:
    render_topbar()
    user = st.session_state["current_user"]
    col_main, col_ai = st.columns([3.2, 1.8], gap="large")
    
    with col_ai:
        # FRAME DIHILANGKAN
        st.markdown("<h3 style='color: #8B0000; margin-top: 0px; margin-bottom: 15px;'>🤖 Tanyadah (Under Development)</h3>", unsafe_allow_html=True)
        
        chat_container = st.container(height=520, border=True)
        for msg in st.session_state["global_chat_log"]:
            chat_container.chat_message(msg["role"]).write(msg["content"])
            
        if prompt := st.chat_input("Tanya strategi operasional..."):
            st.session_state["global_chat_log"].append({"role": "user", "content": prompt})
            chat_container.chat_message("user").write(prompt)
            ans = get_ai_response(prompt, user.get('name'), "Dashboard")
            chat_container.chat_message("assistant").write(ans)
            st.session_state["global_chat_log"].append({"role": "assistant", "content": ans})

    with col_main:
        # FRAME DIHILANGKAN
        greeting = "Pagi" if datetime.now().hour < 12 else ("Siang" if datetime.now().hour < 17 else "Sore")
        st.markdown(f"<h2 style='color: #1E293B; margin-top: 0px;'>Selamat {greeting}, {user.get('name')} ✨</h2><hr style='margin:10px 0; border-color: rgba(0,0,0,0.1);'>", unsafe_allow_html=True)
        
        st.markdown("""<style>
        div[data-testid="column"] div[data-testid="stButton"] button { background: rgba(255,255,255,0.9) !important; height: 90px; border-radius: 15px !important; border: 1px solid rgba(0,0,0,0.1) !important; transition: 0.3s !important; color: #1E293B !important; font-weight: bold !important; box-shadow: 0 4px 6px rgba(0,0,0,0.05); }
        div[data-testid="column"] div[data-testid="stButton"] button:hover { transform: translateY(-3px); border-color: #8B0000 !important; background: white !important; }
        </style>""", unsafe_allow_html=True)
        
        m1, m2, m3 = st.columns(3)
        with m1:
            if st.button("📦 OPERATION\nPO & DN", use_container_width=True): st.switch_page("pages/1_🚚_Operation.py")
        with m2:
            if st.button("📸 ATTENDANCE \n HR & Absen (Under Development)", use_container_width=True): st.switch_page("pages/2_📸_Attendance.py")
        with m3:
            if st.button("✅ TO DO LIST \n Kanban (Under Development)", use_container_width=True): st.switch_page("pages/3_✅_To_Do_List.py")

        st.markdown("<br>", unsafe_allow_html=True)
        
        c_left, c_right = st.columns([2, 1])
        with c_left:
            st.markdown("<h4 style='margin-top:0; color:#1E293B;'>📈 Tren Operasional</h4>", unsafe_allow_html=True)
            st.area_chart(pd.DataFrame(np.random.randn(15, 2), columns=['PO', 'DN']), height=180)

        with c_right:
            s = "padding: 10px; border-radius: 15px; background: rgba(255,255,255,0.7); border: 1px solid rgba(0,0,0,0.1); text-align: center; margin-bottom: 10px;"
            c1, c2 = st.columns(2)
            with c1: st.markdown(f"<div style='{s}'><small style='color:#475569; font-weight:bold;'>PO AKTIF</small><br><b style='font-size:1.5rem; color:#1E293B;'>12</b></div>", unsafe_allow_html=True)
            with c1: st.markdown(f"<div style='{s}'><small style='color:#475569; font-weight:bold;'>PO AKTIF</small><br><b style='font-size:1.5rem; color:#1E293B;'>12</b></div>", unsafe_allow_html=True)
            with c2: st.markdown(f"<div style='{s}'><small style='color:#475569; font-weight:bold;'>TASKS</small><br><b style='font-size:1.5rem; color:#1E293B;'>5</b></div>", unsafe_allow_html=True)
            c3, c4 = st.columns(2)
            with c3: st.markdown(f"<div style='{s}'><small style='color:#475569; font-weight:bold;'>ABSEN</small><br><b style='font-size:1.5rem; color:#1E293B;'>1</b></div>", unsafe_allow_html=True)
            with c4: st.markdown(f"<div style='{s}'><small style='color:#475569; font-weight:bold;'>LEAVE</small><br><b style='font-size:1.5rem; color:#1E293B;'>0</b></div>", unsafe_allow_html=True)