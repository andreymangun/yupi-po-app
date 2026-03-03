import streamlit as st
import pandas as pd
from fpdf import FPDF
from fpdf.enums import XPos, YPos
import requests
import io
import zipfile
from io import StringIO
import os
from datetime import date
import re

# ==========================================
# 1. KONFIGURASI HALAMAN (ERP UI)
# ==========================================
st.set_page_config(page_title="ServeOne ERP System", page_icon="🌐", layout="wide")

st.markdown("""
    <style>
    .block-container { padding-top: 2rem; padding-bottom: 2rem; }
    h1 { color: #0ea5e9; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; font-weight: 800; text-align: center; margin-bottom: 5px;}
    .subtitle { text-align: center; color: #64748b; margin-bottom: 30px; font-size: 1.1rem; }
    div[data-testid="stMetricValue"] { font-size: 1.8rem; color: #0ea5e9; font-weight: 700; }
    .stTextInput>div { width: 70%; margin: 0 auto; }
    .stTextInput label { display: flex; justify-content: center; font-size: 1rem; font-weight: 600; color: #334155;}
    .stButton>button { 
        background: linear-gradient(90deg, #0284c7 0%, #0369a1 100%); 
        color: white; border: none; border-radius: 6px; width: 100%; 
        padding: 0.6rem; font-weight: bold; transition: all 0.3s ease;
    }
    .stButton>button:hover { 
        background: linear-gradient(90deg, #0369a1 0%, #075985 100%); 
        box-shadow: 0 4px 12px rgba(2, 132, 199, 0.4); 
    }
    </style>
    """, unsafe_allow_html=True)

# ==========================================
# 2. PDF ENGINE
# ==========================================
class ServeonePO(FPDF):
    def header(self):
        if os.path.exists("logo.png"): self.image("logo.png", 10, 2, 45)
        self.set_font('Helvetica', 'B', 16)
        self.set_y(10)
        self.cell(0, 10, 'P/O Paper', align='C', new_x=XPos.LMARGIN, new_y=YPos.TOP)
        if os.path.exists("stamp.png"): self.image("stamp.png", 160, 8, 40)
        self.ln(20)

    def footer(self):
        self.set_y(-15)
        self.set_font('Helvetica', 'I', 8)
        self.cell(0, 10, f'Page {self.page_no()} of {{nb}}', align='C')

def clean_val(val):
    if pd.isna(val) or str(val).lower() == 'nan': return ""
    return str(val).encode('latin-1', 'replace').decode('latin-1').strip()

def safe_float(val):
    try:
        if isinstance(val, str): val = val.replace(',', '').replace('"', '').strip()
        return float(val) if val else 0.0
    except: return 0.0

def generate_po_pdf(po_data, po_number):
    pdf = ServeonePO()
    pdf.alias_nb_pages()
    pdf.add_page()
    info = po_data.iloc[0]
    
    vendor_name = clean_val(info.get('Vendor Name', 'Unknown Vendor'))

    # PO INFO
    pdf.set_font('Helvetica', 'B', 10)
    pdf.cell(30, 7, "PO YUPI :")
    pdf.set_font('Helvetica', '', 10)
    pdf.cell(0, 7, clean_val(po_number), new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    
    pdf.set_font('Helvetica', 'B', 10)
    pdf.cell(30, 7, "PO Date :")
    pdf.set_font('Helvetica', '', 10)
    pdf.cell(0, 7, date.today().strftime("%d/%m/%Y"), new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.ln(5)

    # CLIENT & VENDOR
    y_anchor = pdf.get_y()
    pdf.set_font('Helvetica', 'B', 9)
    pdf.cell(95, 5, "CLIENT:")
    pdf.set_xy(105, y_anchor)
    pdf.cell(95, 5, "VENDOR:")
    
    pdf.ln(6)
    pdf.set_font('Helvetica', '', 8)
    y_content = pdf.get_y()
    client_info = "PT. SERVEONE MRO INDONESIA\nC-2-2 LANTAI 2, MM2100\nKab. Bekasi, Prov. Jawa Barat"
    vendor_info = f"{vendor_name}\n{clean_val(info.get('Vendor Address'))}"
    
    pdf.set_xy(10, y_content)
    pdf.multi_cell(92, 4, client_info)
    pdf.set_xy(105, y_content)
    pdf.multi_cell(95, 4, vendor_info)
    pdf.set_y(max(pdf.get_y(), y_content + 15) + 5)

    # SITE LOGIC
    site_raw = ""
    for col in info.index:
        if 'SITE' in str(col).upper():
            site_raw = clean_val(info[col]).upper()
            break
    
    if 'IDN' in site_raw:
        site_label = "IDN"
        full_addr = "Jl. Pancasila IV No.9, Cicadas, Kec. Gn. Putri, Kabupaten Bogor, Jawa Barat 16964"
    else:
        site_label = "KRG"
        full_addr = "JI.Grompol-Jambangan, Kaliwuluh Lor, Kebakkramat, Karanganyar, Jawa Tengah"

    pdf.set_font('Helvetica', 'B', 9)
    pdf.cell(0, 6, f"DELIVERY ADDRESS: PT. YUPI INDO JELLY GUM Tbk ({site_label})", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.set_font('Helvetica', '', 8)
    pdf.multi_cell(0, 5, full_addr)
    pdf.ln(5)

    # TABLE HEADER
    pdf.set_fill_color(240, 240, 240)
    pdf.set_font('Helvetica', 'B', 7)
    cols = [
        ("Deliv Req", 20), ("SERVEONE PO", 30), ("Item Code", 20), 
        ("Item Name & Spec", 50), ("Qty", 10), ("Unit", 10), ("Price", 25), ("Amount", 25)
    ]
    for txt, w in cols:
        pdf.cell(w, 8, txt, border=1, fill=True, align='C')
    pdf.ln()

    # TABLE BODY
    pdf.set_font('Helvetica', '', 7)
    total_dpp = 0
    remark_list = []

    for _, row in po_data.iterrows():
        name_spec = f"{clean_val(row.get('Item name'))}\n{clean_val(row.get('Spec'))}"
        qty = safe_float(row.get("Ord. Q'ty"))
        price = safe_float(row.get('PURCHASE PRICE'))
        amount = qty * price
        total_dpp += amount
        
        rem = clean_val(row.get('REMARK YUPI'))
        if rem: remark_list.append(rem)

        lines = pdf.multi_cell(50, 4, name_spec, dry_run=True, output="LINES")
        row_h = max(10, len(lines) * 4 + 2)

        if pdf.get_y() + row_h > 240: pdf.add_page()

        x, y = pdf.get_x(), pdf.get_y()
        pdf.cell(20, row_h, clean_val(row.get('Req. Dlv Date')), border=1, align='C')
        pdf.cell(30, row_h, clean_val(row.get('PO SEMENTARA')), border=1, align='C')
        pdf.cell(20, row_h, clean_val(row.get('Item Yupi')), border=1, align='C')
        pdf.cell(50, row_h, "", border=1) 
        pdf.cell(10, row_h, f"{qty:,.0f}", border=1, align='C')
        pdf.cell(10, row_h, clean_val(row.get('Unit')), border=1, align='C')
        pdf.cell(25, row_h, f"{price:,.0f}", border=1, align='R')
        pdf.cell(25, row_h, f"{amount:,.0f}", border=1, align='R')

        pdf.set_xy(x + 70, y + 1) 
        pdf.multi_cell(50, 4, name_spec, border=0, align='L')
        pdf.set_xy(10, y + row_h)

    # FOOTER
    ppn = total_dpp * 0.11
    grand_total = total_dpp + ppn
    pdf.set_font('Helvetica', 'B', 8)
    for label, val in [("TOTAL AMOUNT", total_dpp), ("PPN (11%)", ppn), ("GRAND TOTAL", grand_total)]:
        pdf.set_x(150)
        pdf.cell(25, 5, label, border=1, align='R')
        pdf.cell(25, 5, f"{val:,.0f}", border=1, align='R', new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    pdf.ln(5)
    pdf.set_font('Helvetica', 'B', 9)
    pdf.cell(0, 6, "REMARKS:", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.set_font('Helvetica', '', 8)
    pdf.multi_cell(0, 5, ", ".join(set(remark_list)) if remark_list else "-")

    return bytes(pdf.output()), site_label, vendor_name

# ==========================================
# 3. INTERFACE & LOGIC UTAMA
# ==========================================
st.markdown("<h1>SERVEONE PO ENGINE</h1>", unsafe_allow_html=True)
st.markdown("<div class='subtitle'>Automated Document Processing System</div>", unsafe_allow_html=True)

link = st.text_input("Link CSV Google Sheets", placeholder="Masukkan link CSV di sini lalu tekan Enter...")
st.divider()

if link:
    try:
        with st.spinner("Menarik data dari server..."):
            res = requests.get(link)
            df_raw = pd.read_csv(StringIO(res.text), low_memory=False)
            
            mask = df_raw.apply(lambda row: row.astype(str).str.contains('PO YUPI').any(), axis=1)
            idx = df_raw[mask].index[0]
            df = df_raw.iloc[idx+1:].copy()
            df.columns = [str(c).strip() for c in df_raw.iloc[idx]]
            df = df.dropna(subset=['PO YUPI'])
            
            # Memperbaiki kolom numerik dan membersihkan nama Vendor
            df['Vendor Name'] = df.get('Vendor Name', pd.Series(['Unknown']*len(df))).fillna('Unknown').astype(str).str.strip()
            
            for col in ['PURCHASE PRICE', "Ord. Q'ty", 'AMOUNT', 'PURCHASE AMOUNT']:
                if col in df.columns:
                    df[col] = df[col].apply(safe_float)

            # --- PANEL FILTER ---
            st.markdown("### 🎛️ Control Panel")
            st.info("💡 **Smart Split & Filter:** Sistem otomatis memecah PO jika terdapat >1 Vendor, dan bisa mengabaikan item yang belum memiliki harga atau belum ditentukan vendornya.")
            
            col_filter1, col_filter2, col_filter3 = st.columns([1, 1.5, 1])
            
            # 1. Filter PO Date
            with col_filter1:
                date_cols = [c for c in df.columns if 'PO DATE' in c.upper()]
                date_col_name = date_cols[0] if date_cols else None
                if date_col_name:
                    available_dates = sorted(df[date_col_name].dropna().astype(str).unique())
                    selected_date = st.selectbox("📅 Filter by PO Date:", ["Semua Tanggal"] + available_dates)
                    if selected_date != "Semua Tanggal":
                        df = df[df[date_col_name] == selected_date]

            # 2. Smart Filter (Menjawab kendala Anda)
            with col_filter3:
                st.markdown("<br>", unsafe_allow_html=True)
                exclude_unready = st.checkbox("🚫 Sembunyikan Item Belum Siap (Vendor 'Serveone' / Harga 0)", value=True)
                
                if exclude_unready:
                    # Buang baris yang nama vendornya mengandung kata SERVEONE (case-insensitive)
                    df = df[~df['Vendor Name'].str.contains('SERVEONE|SERVE ONE', case=False, na=False)]
                    # Buang baris yang harganya masih 0
                    if 'PURCHASE PRICE' in df.columns:
                        df = df[df['PURCHASE PRICE'] > 0]

            # 3. Multi-Select PO
            with col_filter2:
                all_pos = sorted(df['PO YUPI'].unique().tolist())
                selected_pos = st.multiselect("🎯 Pilih Nomor PO:", all_pos, default=all_pos)

            if not selected_pos:
                st.warning("Data kosong. Mungkin karena terfilter oleh 'Sembunyikan Item Belum Siap' atau tidak ada PO terpilih.")
                st.stop()

            final_data = df[df['PO YUPI'].isin(selected_pos)]

            # --- PANEL METRIK ---
            col_met1, col_met2, col_met3 = st.columns(3)
            # Hitung unik dokumen yang akan ter-generate (Group by PO dan Vendor)
            total_documents = len(final_data.groupby(['PO YUPI', 'Vendor Name']))
            
            col_met1.metric("Dokumen PO Dicetak", total_documents, help="Jumlah PDF yang akan dihasilkan (sudah dipecah per Vendor)")
            col_met2.metric("Total Item Baris", len(final_data))
            
            if 'PURCHASE AMOUNT' in final_data.columns:
                estimasi_amount = final_data['PURCHASE AMOUNT'].sum()
            elif 'AMOUNT' in final_data.columns:
                estimasi_amount = final_data['AMOUNT'].sum()
            else:
                estimasi_amount = sum((final_data["Ord. Q'ty"] * final_data['PURCHASE PRICE']).fillna(0))
                
            col_met3.metric("Estimasi Total Nilai (DPP)", f"IDR {estimasi_amount:,.0f}")

            st.markdown("### 📊 Preview Data Tersortir")
            st.dataframe(final_data, width="stretch", height=250)
            st.divider()

            # --- EKSEKUSI GENERATE (DENGAN LOGIKA SPLIT VENDOR) ---
            st.markdown("### ⚙️ Eksekusi Dokumen")
            col_gen1, col_gen2 = st.columns(2)
            
            with col_gen1:
                st.info("Single Generate (Berdasarkan 1 PO)")
                if len(selected_pos) == 1:
                    po_num = selected_pos[0]
                    po_data_single_po = final_data[final_data['PO YUPI'] == po_num]
                    unique_vendors = po_data_single_po['Vendor Name'].unique()
                    
                    if len(unique_vendors) > 1:
                        st.warning(f"PO {po_num} ini memiliki **{len(unique_vendors)} Vendor berbeda**. Silakan unduh satu per satu:")
                    
                    for v_name in unique_vendors:
                        # Membersihkan nama vendor untuk nama file agar tidak error
                        safe_vname = re.sub(r'[^a-zA-Z0-9 \.\-]', '', v_name).strip()
                        po_data_vendor = po_data_single_po[po_data_single_po['Vendor Name'] == v_name]
                        
                        if st.button(f"📄 Generate PDF [{po_num}] - {safe_vname}", key=f"btn_{safe_vname}"):
                            po_bytes, site, vendor = generate_po_pdf(po_data_vendor, po_num)
                            st.download_button(
                                label=f"📥 Unduh {safe_vname}", 
                                data=po_bytes, 
                                file_name=f"[{site}] PO SERVEONE - {po_num} - {safe_vname}.pdf", 
                                mime="application/pdf",
                                key=f"dl_{safe_vname}"
                            )
                else:
                    st.button("📄 Generate Single PDF", disabled=True, help="Hanya aktif jika 1 PO dipilih")

            with col_gen2:
                st.info(f"Bulk Generate akan merangkum **{total_documents} Dokumen PDF** ke dalam satu file ZIP.")
                if st.button("📦 Generate Semua ke ZIP (Auto-Split)"):
                    zip_buffer = io.BytesIO()
                    progress_bar = st.progress(0)
                    
                    with st.spinner("Memproses dokumen per PO dan per Vendor..."):
                        with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                            # Mengelompokkan berdasarkan PO dan VENDOR
                            grouped_data = final_data.groupby(['PO YUPI', 'Vendor Name'])
                            
                            for idx, ((po_num, v_name), po_data_vendor) in enumerate(grouped_data):
                                po_bytes, site, vendor = generate_po_pdf(po_data_vendor, po_num)
                                safe_vname = re.sub(r'[^a-zA-Z0-9 \.\-]', '', vendor).strip()
                                
                                zip_file.writestr(f"[{site}] PO SERVEONE - {po_num} - {safe_vname}.pdf", po_bytes)
                                progress_bar.progress((idx + 1) / total_documents)
                        
                        st.success("Berhasil! File ZIP siap diunduh.")
                        st.download_button(
                            label="📥 Unduh Bulk PO (.zip)",
                            data=zip_buffer.getvalue(),
                            file_name=f"BULK_PO_SERVEONE_{date.today()}.zip",
                            mime="application/zip"
                        )

    except Exception as e:
        st.error(f"Sistem mendeteksi kesalahan pada struktur Google Sheets: {e}")