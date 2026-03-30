import streamlit as st
import pandas as pd
from datetime import date

# 1. KONFIGURASI HARUS PALING ATAS
st.set_page_config(page_title="Operation System", page_icon="🚚", layout="wide", initial_sidebar_state="collapsed")

# 2. INISIALISASI MEMORI AMAN (Pencegah Error Merah)
from utils.state import init_session
init_session()

# Pastikan semua kunci yang dipakai di halaman ini benar-benar ada di memori
state_keys = {
    "op_step": 1, 
    "operation_df": None,
    "generated_po_bytes": None, 
    "generated_po_filename": None, 
    "g_dn_bytes": None, 
    "g_dn_file": None, 
    "g_po_bytes": None, 
    "g_po_file": None, 
    "search_po_input": "", 
    "copilot_history": [{"role": "assistant", "content": "Halo! Ketik nomor PO atau Vendor, saya akan merekap statusnya."}], 
    "dn_counter_data": {}
}
for key, val in state_keys.items():
    if key not in st.session_state:
        st.session_state[key] = val

# 3. IMPORTS CUSTOM
from utils.auth import require_login
from utils.theme import inject_css
from utils.topbar import render_topbar
from utils.guards import is_superuser

from services.operation_service import (
    init_operation_state, load_operation_data, clear_operation_data, 
    reset_generated_files, get_available_po_numbers, filter_po_dataframe, 
    get_available_vendors, filter_vendor_dataframe, get_display_dataframe, build_po_pdf_bytes
)
from pdf_engine.dn_pdf import generate_dn_pdf 

init_operation_state()
inject_css()
require_login()
render_topbar()

# 4. LAYOUT UTAMA
col_main, col_ai = st.columns([3.5, 1.5], gap="large")

with col_ai:
    st.markdown("<h3 style='color:#1E293B; margin-top:0px;'>🤖 Tanyadah (Under Development)</h3>", unsafe_allow_html=True)
    st.info("Pusat asisten terintegrasi untuk merekap status vendor.")
    
    chat_container = st.container(height=500, border=True)
    for msg in st.session_state["copilot_history"]:
        chat_container.chat_message(msg["role"]).write(msg["content"])
        
    prompt = st.chat_input("Tanya OpenClaw...")
    if prompt:
        st.session_state["copilot_history"].append({"role": "user", "content": prompt})
        st.session_state["copilot_history"].append({"role": "assistant", "content": f"⏳ Menganalisis '{prompt}'..."})
        st.rerun()

with col_main:
    st.markdown("<h2 style='color:#1E293B; margin-top:0px;'>🚚 Operation System</h2>", unsafe_allow_html=True)
    
    st.markdown("""
    <div style="width: 100%; overflow: hidden; background: linear-gradient(90deg, #8B0000, #b22222); color: white; padding: 12px; border-radius: 8px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); margin-bottom: 20px;">
        <div style="animation: slide 20s linear infinite; white-space: nowrap; font-weight: 500; letter-spacing: 0.5px;">
            ✨ <b>INFO SYSTEM:</b> Tiga Modul Enterprise (Workflow Log, Vendor Scoring, Inventory Forecast) telah diaktifkan di menu SUMMARY. ✨
        </div>
    </div>
    <style>@keyframes slide { 0% { transform: translateX(100%); } 100% { transform: translateX(-100%); } }</style>
    """, unsafe_allow_html=True)
    
    kategori_sheet = {
        "Raw Material (RM)": "https://docs.google.com/spreadsheets/d/e/2PACX-1vTS35p__c-snrnvGNu6KPiK4yB2SE_75ikNEsMkO-st-PEYTAJZvuzaDUXcdM9bI-so7VTyenZYx7GA/pub?gid=0&single=true&output=csv",
        "Packaging Material (PM)": "https://docs.google.com/spreadsheets/d/e/2PACX-1vQBC7ZDMg_vzQNWjjnr4EylVeLFZKlg7nFcvyuygOH-e3TEZS_H_E2E8xLlWX7326YJRnH37AdU-C1M/pub?gid=0&single=true&output=csv",
        "Spare Part/Consumable (SP)": "https://docs.google.com/spreadsheets/d/e/2PACX-1vS14-HlMoSqURQRmhyDdhOcQnneBZC48ccsFaGd5lDu39f9uSyBl3EIFeIGTjqmRoFoyESG3YHqZh58/pub?gid=0&single=true&output=csv"
    }

    pilihan_kategori = st.radio("Pilih Kategori Kebutuhan:", list(kategori_sheet.keys()), horizontal=True)
    csv_url = kategori_sheet[pilihan_kategori]
    
    if "Raw" in pilihan_kategori: cat_code = "RM"
    elif "Packaging" in pilihan_kategori: cat_code = "PM"
    else: cat_code = "SP"

    def get_site_name(df_source):
        site = "KRG" 
        try:
            for col in df_source.columns:
                col_lower = str(col).lower()
                if "site" in col_lower or "plant" in col_lower:
                    valid_rows = df_source[col].dropna()
                    if not valid_rows.empty:
                        site_val = str(valid_rows.iloc[0]).upper()
                        if "IDN" in site_val or "GUNUNG PUTRI" in site_val or "GNP" in site_val: site = "IDN"
                        elif "KRG" in site_val or "KARANGJATI" in site_val: site = "KRG"
                        break 
        except: pass
        return site

    def get_custom_dn_number():
        month_key = date.today().strftime("%m%Y")
        if month_key not in st.session_state["dn_counter_data"]: 
            st.session_state["dn_counter_data"][month_key] = {"RM": 1, "PM": 1, "SP": 1}
        current_num = st.session_state["dn_counter_data"][month_key].get(cat_code, 1)
        date_str = date.today().strftime('%d%m%Y')
        return f"DS/{cat_code}/{date_str}/{current_num:03d}"

    def bump_custom_dn_counter():
        month_key = date.today().strftime("%m%Y")
        if month_key in st.session_state["dn_counter_data"]:
            st.session_state["dn_counter_data"][month_key][cat_code] += 1

    col_load, col_clear = st.columns(2)
    with col_load:
        if st.button("Muat Data Kategori", use_container_width=True):
            with st.spinner("Memuat data dari Google Sheet..."):
                df_loaded = load_operation_data(csv_url)
                st.session_state["operation_df"] = df_loaded # Simpan langsung ke memori
            st.success(f"Data dimuat. Total {len(df_loaded)} baris.")
    with col_clear:
        if st.button("Clear Data Lokal", use_container_width=True):
            clear_operation_data()
            st.session_state["operation_df"] = None
            st.success("Data berhasil dibersihkan.")

    # ==========================================
    # PERBAIKAN UTAMA: AMBIL DATA DENGAN AMAN (.get)
    # ==========================================
    df = st.session_state.get("operation_df")
    
    if df is not None and not df.empty:
        if is_superuser(): opsi_mode = ["SUMMARY", "PURCHASE ORDER", "DELIVERY NOTE"]
        else: opsi_mode = ["PURCHASE ORDER", "DELIVERY NOTE"]
            
        mode_aktif = st.radio("Mode Tampilan Dokumen:", opsi_mode, horizontal=True)
        st.markdown("<br>", unsafe_allow_html=True)

        if mode_aktif in ["PURCHASE ORDER", "DELIVERY NOTE"]:
            date_col = None
            for col in df.columns:
                if 'date' in col.lower() or 'tgl' in col.lower():
                    date_col = col; break
                    
            available_dates = ["Semua Tanggal"]
            if date_col:
                try:
                    clean_dates = [str(d) for d in df[date_col].dropna().unique()]
                    def sort_key(date_str):
                        try: return pd.to_datetime(date_str + f"-{date.today().year}", format='%d-%b-%Y')
                        except:
                            try: return pd.to_datetime(date_str)
                            except: return pd.Timestamp('1970-01-01')
                    clean_dates_sorted = sorted(clean_dates, key=sort_key)
                    available_dates.extend(clean_dates_sorted)
                except Exception:
                    clean_dates = df[date_col].dropna().astype(str).unique().tolist()
                    available_dates.extend(sorted(clean_dates))

            col_paste, col_date = st.columns([2, 1])
            with col_paste: 
                st.text_input("🔍 Cari / Paste Nomor PO (Tekan Enter):", key="search_po_input", placeholder="Contoh: 410100137...")
                
            with col_date: 
                selected_date = st.selectbox("📅 Filter Tanggal:", available_dates)

            if selected_date != "Semua Tanggal" and date_col: filtered_df = df[df[date_col].astype(str) == selected_date]
            else: filtered_df = df.copy()
                
            po_numbers = get_available_po_numbers(filtered_df)
            if not po_numbers:
                st.warning("Tidak ada PO ditemukan pada kriteria tersebut.")
                st.stop()

            selected_po = None
            search_query = st.session_state["search_po_input"].strip().lower()
            
            if search_query:
                matching_pos = [po for po in po_numbers if search_query in str(po).lower()]
                if matching_pos:
                    selected_po = matching_pos[0]
                    st.success(f"✅ PO Ditemukan: {selected_po}")
                else:
                    st.error(f"❌ PO mengandung '{search_query}' tidak ditemukan di tabel data.")
                    st.stop()
            else:
                selected_po = st.selectbox("📦 Pilih PO dari Daftar:", po_numbers, index=0)

            po_df = filter_po_dataframe(filtered_df, selected_po)
            vendors = get_available_vendors(po_df)
            selected_vendor = st.selectbox("🏭 Pilih Vendor:", vendors, key="selected_ven")
            vendor_df = filter_vendor_dataframe(po_df, selected_vendor)

            st.markdown("---")

            if mode_aktif == "PURCHASE ORDER":
                st.markdown("### 🚀 Fast Track (Cetak Langsung PO)")
                if st.button("📄 CETAK LANGSUNG DOKUMEN (PO)", use_container_width=True, type="primary"):
                    df_for_pdf = vendor_df.copy().reset_index(drop=True)
                    df_for_pdf["Pajak PPN 11%"] = True
                    df_for_pdf["Pajak PPh 2%"] = False
                    df_for_pdf["Catatan / Remark"] = ""
                    with st.spinner("Mencetak PDF..."):
                        site_lbl = get_site_name(df_for_pdf)
                        st.session_state["generated_po_bytes"] = build_po_pdf_bytes(df_for_pdf)
                        st.session_state["generated_po_filename"] = f"[{site_lbl}] PO - {selected_po} - {selected_vendor}.pdf"
                if st.session_state.get("generated_po_bytes"):
                    st.download_button("⬇️ Download PO PDF", data=st.session_state["generated_po_bytes"], file_name=st.session_state["generated_po_filename"], mime="application/pdf", use_container_width=True, key="dl_fast_po")

            st.markdown("### ⚙️ Mode Lanjut (Custom Edit)")
            step1_col, step2_col, step3_col = st.columns(3)
            with step1_col:
                if st.button("Step 1 • Preview", use_container_width=True): st.session_state["op_step"] = 1
            with step2_col:
                if st.button("Step 2 • Pilih & Edit", use_container_width=True): st.session_state["op_step"] = 2
            with step3_col:
                if st.button("Step 3 • Generate", use_container_width=True): st.session_state["op_step"] = 3

            curr_step = st.session_state.get("op_step", 1)
            edit_key = f"edit_baru_{selected_po}_{mode_aktif}"

            if curr_step == 1:
                st.dataframe(get_display_dataframe(vendor_df).head(50), use_container_width=True, hide_index=True)

            elif curr_step == 2:
                edit_df = get_display_dataframe(vendor_df).copy().reset_index(drop=True)
                
                if edit_key not in st.session_state:
                    edit_df.insert(0, "Cetak", False)
                    if mode_aktif == "PURCHASE ORDER":
                        edit_df["PPN 11%"] = True
                        edit_df["PPh Potongan"] = False
                    if mode_aktif == "DELIVERY NOTE":
                        if cat_code == "RM":
                            edit_df["No Batch"], edit_df["Jml Batch"], edit_df["Expired Date"] = "", "", ""
                        elif cat_code == "PM":
                            edit_df["Coding"] = ""
                        edit_df["Catatan / Remark"] = ""
                    st.session_state[edit_key] = edit_df

                current_edit_df = st.session_state[edit_key].copy()
                
                col_cfg = {"Cetak": st.column_config.CheckboxColumn("Cetak Dok.", width="small")}
                if mode_aktif == "PURCHASE ORDER":
                    col_cfg["PPN 11%"] = st.column_config.CheckboxColumn("PPN 11%")
                    col_cfg["PPh Potongan"] = st.column_config.CheckboxColumn("PPh Potongan")
                if mode_aktif == "DELIVERY NOTE":
                    if cat_code == "RM":
                        col_cfg["No Batch"] = st.column_config.TextColumn("No Batch (Edit di bawah)", disabled=True)
                        col_cfg["Jml Batch"] = st.column_config.TextColumn("Jml Batch (Edit di bawah)", disabled=True)
                        col_cfg["Expired Date"] = st.column_config.TextColumn("Exp Date (Edit di bawah)", disabled=True)
                    elif cat_code == "PM":
                        col_cfg["Coding"] = st.column_config.TextColumn("Coding (Edit di bawah)", disabled=True)
                    col_cfg["Catatan / Remark"] = st.column_config.TextColumn("Remark (Edit di bawah)", disabled=True)

                st.write("**1. Daftar Item:**")
                st.info("Centang pada kolom 'Cetak Dok.' HANYA untuk item yang ingin dimasukkan ke PDF.")
                
                with st.form(f"frm_grid_{selected_po}_{mode_aktif}"):
                    edited_df = st.data_editor(current_edit_df, use_container_width=True, hide_index=True, column_config=col_cfg)
                    save_grid = st.form_submit_button("Simpan Status Centang", use_container_width=True)
                
                if save_grid:
                    st.session_state[edit_key] = edited_df
                    st.rerun()

                edited_df = st.session_state[edit_key]

                if mode_aktif == "DELIVERY NOTE":
                    st.write("**2. Isi Detail Per Item:**")
                    st.info("Pilih item dari opsi di bawah, isi detailnya (bisa Enter untuk multi-baris), lalu klik Simpan.")
                    
                    item_options = []
                    for idx, row in edited_df.iterrows():
                        po_sementara = str(row.get('PO SEMENTARA', f'Item {idx+1}'))
                        item_options.append(f"{idx} - {po_sementara}")
                        
                    selected_item_str = st.selectbox("🎯 Pilih Item yang akan diisi detailnya:", item_options)
                    
                    if selected_item_str:
                        selected_idx = int(selected_item_str.split(" - ")[0])
                        row_data = edited_df.loc[selected_idx]
                        
                        with st.form(f"frm_detail_per_item_{selected_idx}"):
                            new_batch, new_jml, new_exp, new_coding = "", "", "", ""
                            
                            if cat_code == "RM":
                                new_batch = st.text_area("No Batch", value=str(row_data.get("No Batch", "")).replace("nan",""), help="Tekan Enter untuk baris baru")
                                new_jml = st.text_area("Jml Batch", value=str(row_data.get("Jml Batch", "")).replace("nan",""))
                                new_exp = st.text_area("Expired Date", value=str(row_data.get("Expired Date", "")).replace("nan",""))
                            elif cat_code == "PM":
                                new_coding = st.text_area("Coding", value=str(row_data.get("Coding", "")).replace("nan",""))
                                
                            new_remark = st.text_area("Catatan / Remark", value=str(row_data.get("Catatan / Remark", "")).replace("nan",""))
                            
                            if st.form_submit_button("Simpan Detail ke Item Ini", type="primary"):
                                df_to_update = st.session_state[edit_key].copy()
                                
                                if cat_code == "RM":
                                    df_to_update.at[selected_idx, "No Batch"] = new_batch
                                    df_to_update.at[selected_idx, "Jml Batch"] = new_jml
                                    df_to_update.at[selected_idx, "Expired Date"] = new_exp
                                elif cat_code == "PM":
                                    df_to_update.at[selected_idx, "Coding"] = new_coding
                                
                                df_to_update.at[selected_idx, "Catatan / Remark"] = new_remark
                                
                                st.session_state[edit_key] = df_to_update
                                st.success(f"✅ Detail untuk '{selected_item_str.split(' - ')[1]}' berhasil disimpan! Silakan pilih item lain atau klik Step 3.")
                                st.rerun()

            elif curr_step == 3:
                df_pdf = vendor_df.copy().reset_index(drop=True)
                if edit_key not in st.session_state: 
                    st.warning("Harap atur dan simpan detail di Step 2 terlebih dahulu.")
                    st.stop()
                    
                e_df = st.session_state[edit_key]
                selected_indices = e_df.index[e_df["Cetak"] == True].tolist()
                
                if not selected_indices:
                    st.warning("Belum ada item yang dicentang 'Cetak Dok.' di Step 2.")
                    st.stop()

                if mode_aktif == "PURCHASE ORDER":
                    df_pdf["Pajak PPN 11%"] = e_df["PPN 11%"]; df_pdf["Pajak PPh 2%"] = e_df["PPh Potongan"]
                if mode_aktif == "DELIVERY NOTE":
                    df_pdf["Catatan / Remark"] = e_df["Catatan / Remark"].astype(str).replace("nan", "")
                    if cat_code == "RM":
                        df_pdf["No Batch"] = e_df["No Batch"].astype(str).replace("nan", ""); df_pdf["Jml Batch"] = e_df["Jml Batch"].astype(str).replace("nan", ""); df_pdf["Expired Date"] = e_df["Expired Date"].astype(str).replace("nan", "")
                    elif cat_code == "PM":
                        df_pdf["Coding"] = e_df["Coding"].astype(str).replace("nan", "")

                df_pdf = df_pdf.iloc[selected_indices].copy().reset_index(drop=True)

                if mode_aktif == "DELIVERY NOTE":
                    st.subheader("🧾 Pengaturan Kurir (DN)")
                    d1, d2, d3 = st.columns(3)
                    with d1: dn_ven = st.text_input("Vendor / Pengirim")
                    with d2: dn_pol = st.text_input("No. Polisi")
                    with d3: st.text_input("Preview DN", value=get_custom_dn_number(), disabled=True)
                    
                    if st.button("📄 Cetak Kustom (DN)", use_container_width=True, type="primary"):
                        c_dn = get_custom_dn_number()
                        pdf_bytes, site_lbl, ven_name = generate_dn_pdf(po_data=df_pdf, po_number=selected_po, dn_vendor=dn_ven, no_pol=dn_pol, dn_remarks="", dn_serveone_number=c_dn, category=cat_code)
                        
                        site_lbl = get_site_name(df_pdf)
                        st.session_state["g_dn_bytes"] = pdf_bytes
                        st.session_state["g_dn_file"] = f"[{site_lbl}] DN - {selected_po} - {selected_vendor}.pdf"
                        st.success("Dokumen berhasil dibuat.")
                        
                    if st.session_state.get("g_dn_bytes"):
                        if st.download_button("⬇️ Download DN", data=st.session_state["g_dn_bytes"], file_name=st.session_state["g_dn_file"], mime="application/pdf", use_container_width=True, key="dl_c_dn"):
                            bump_custom_dn_counter()
                            st.session_state["g_dn_bytes"] = None

                elif mode_aktif == "PURCHASE ORDER":
                    if st.button("📄 Cetak Kustom (PO)", use_container_width=True, type="primary"):
                        site_lbl = get_site_name(df_pdf)
                        st.session_state["g_po_bytes"] = build_po_pdf_bytes(df_pdf)
                        st.session_state["g_po_file"] = f"[{site_lbl}] PO - {selected_po} - {selected_vendor}.pdf"
                        st.success("Dokumen berhasil dibuat.")
                    if st.session_state.get("g_po_bytes"):
                        st.download_button("⬇️ Download PO", data=st.session_state["g_po_bytes"], file_name=st.session_state["g_po_file"], mime="application/pdf", use_container_width=True, key="dl_c_po")

        elif mode_aktif == "SUMMARY":
            st.subheader(f"📊 Enterprise Summary: {pilihan_kategori}")
            tab_sum, tab_rec, tab_app = st.tabs(["📊 Dashboard Utama", "💡 Sourcing Logic", "✅ Workflow Log"])
            with tab_sum:
                st.metric("Total Baris Data", len(df))
                st.dataframe(df.head(15), use_container_width=True)
            with tab_rec:
                col_item = df.columns[0]
                for c in df.columns:
                    if 'item' in c.lower() or 'nama' in c.lower(): col_item = c; break
                col_ven = df.columns[1]
                for c in df.columns:
                    if 'vendor' in c.lower(): col_ven = c; break
                try:
                    sourcing_df = df.groupby([col_item, col_ven]).size().reset_index(name='Frequency')
                    idx_best = sourcing_df.groupby(col_item)['Frequency'].idxmax()
                    best_vendors = sourcing_df.loc[idx_best].sort_values(by='Frequency', ascending=False)
                    df_rekomendasi = best_vendors.head(20).rename(columns={col_item: "Nama Item", col_ven: "Rekomendasi Vendor", "Frequency": "Skor"})
                    st.dataframe(df_rekomendasi, hide_index=True, use_container_width=True)
                except Exception as e:
                    st.warning(f"Tidak dapat memproses rekomendasi: {e}")
            with tab_app: st.info("Integrasi database tertunda.")