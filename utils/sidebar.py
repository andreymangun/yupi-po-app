import streamlit as st
from utils.auth import logout

def render_sidebar(active_page="dashboard"):
    with st.sidebar:
        st.markdown(f"""
            <div style="text-align: center; margin-bottom: 30px;">
                <h1 style="color: #8B0000; font-size: 2.2rem; font-weight: 900; margin: 0; text-shadow: 0 2px 4px rgba(255,255,255,0.8);">ServeOne</h1>
                <p style="color: #1E293B; font-size: 0.8rem; font-weight: bold; margin-top: -5px;">Enterprise Portal</p>
                <hr style="border-top: 1px solid rgba(255,255,255,0.4); margin-top: 15px;">
            </div>
        """, unsafe_allow_html=True)
        
        user = st.session_state.get("current_user", {})
        
        def nav_item(label, page_name, icon, is_active):
            active_style = "background: rgba(255,255,255,0.5); border-left: 4px solid #8B0000; font-weight: bold; color: #1E293B;" if is_active else "background: transparent; color: #334155; border-left: 4px solid transparent;"
            if st.button(f"{icon} {label}", use_container_width=True, key=f"nav_{page_name}"):
                if page_name == "dashboard": st.switch_page("app.py")
                else: st.switch_page(f"pages/{page_name}.py")
            
            st.markdown(f"""
                <style>
                div[data-testid="stSidebar"] div[data-testid="stVerticalBlock"] div[data-testid="stButton"] button[key="nav_{page_name}"] {{
                    {active_style}
                    border-top: none; border-right: none; border-bottom: none;
                    text-align: left; justify-content: flex-start;
                    padding: 10px 15px; border-radius: 0 10px 10px 0; margin-bottom: 8px;
                    transition: 0.3s;
                }}
                div[data-testid="stSidebar"] div[data-testid="stVerticalBlock"] div[data-testid="stButton"] button[key="nav_{page_name}"]:hover {{
                    background: rgba(255,255,255,0.7); color: #1E293B;
                }}
                </style>
            """, unsafe_allow_html=True)

        nav_item("Dashboard", "dashboard", "📊", active_page == "dashboard")
        st.markdown("<br>", unsafe_allow_html=True)
        nav_item("Operation", "1_🚚_Operation", "🚚", active_page == "operation")
        nav_item("Attendance", "2_📸_Attendance", "📸", active_page == "attendance")
        nav_item("To-Do List", "3_✅_To_Do_List", "✅", active_page == "to_do")
        
        st.markdown("<div style='flex-grow: 1; height: 120px;'></div><hr style='border-top: 1px solid rgba(255,255,255,0.4);'>", unsafe_allow_html=True)
        
        c1, c2 = st.columns([1, 4])
        with c1:
            st.markdown(f"<div style='background: rgba(139,0,0,0.8); color: white; border-radius: 50%; width: 40px; height: 40px; display: flex; align-items: center; justify-content: center; font-weight: bold; border: 1px solid rgba(255,255,255,0.5);'>{user.get('name', 'U')[0].upper()}</div>", unsafe_allow_html=True)
        with c2:
            st.markdown(f"<strong style='color:#1E293B;'>{user.get('name', 'User')}</strong><br><span style='font-size: 0.8rem; color: #475569; font-weight:bold;'>{user.get('role', 'user').capitalize()}</span>", unsafe_allow_html=True)
            
        if st.button("🚪 Logout", use_container_width=True, type="secondary", key="logout_btn"):
            logout()
            st.rerun()