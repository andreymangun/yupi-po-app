import streamlit as st
import pandas as pd
from fpdf import FPDF
from fpdf.enums import XPos, YPos
import requests
import io
import zipfile
from io import StringIO
import os
from datetime import date, datetime
import re

# ==========================================
# 1. KONFIGURASI HALAMAN & CSS
# ==========================================
st.set_page_config(page_title="ServeOne ERP System", page_icon="🌐", layout="wide")

st.markdown("""
    <style>
    /* CSS LOGO CENTER & TRANSPARENT */
    [data-testid="stSidebar"] img {
        display: block;
        margin-left: auto;
        margin-right: auto;
        mix-blend-mode: multiply;
        margin-bottom: 20px;
    }
    .block-container { padding-top: 2rem; padding-bottom: 2rem; }
    h1 { color: #0ea5e9; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; font-weight: 800; text-align: center; margin-bottom: 5px;}
    .subtitle { text-align: center; color: #64748b; margin-bottom: 30px; font-size: 1.1rem; }
    div[data-testid="stMetricValue"] { font-size: 1.8rem; color: #0ea5e9; font-weight: 700; }
    .stTextInput>div { width: 70%; margin: 0 auto; }
    .stTextInput label { display: flex; justify-content: center; font-size: 1rem; font-weight: 600; color: #334155;}
    .btn-po>div>button { background: linear-gradient(90deg, #0284c7 0%, #0369a1 100%); color: white; border-radius: 6px; width: 100%; font-weight: bold; }
    .btn-dn>div>button { background: linear-gradient(90deg, #10b981 0%, #047857 100%); color: white; border-radius: 6px; width: 100%; font-weight: bold; }
    </style>
    """, unsafe_allow_html=True)

def clean_val(val):
    if pd.isna(val) or str(val).lower() == 'nan': return ""
    return str(val).encode('latin-1', 'replace').decode('latin-1').strip()

def safe_float(val):
    if pd.isna(val) or str(val).lower() == 'nan': return 0.0
    try:
        if isinstance(val, str): val = val.replace(',', '').replace('"', '').strip()
        return float(val) if val else 0.0
    except: return 0.0

def format_currency(val):
    """Format Uang: Memaksa 2 angka desimal di belakang koma untuk Price, Amount, dan Total"""
    try:
        val = float(val)
        return f"{val:,.2f}"
    except: return "0.00"

def format_qty(val):
    """Format Qty: Menghilangkan .00 jika angkanya bulat (misal: 1,000 Pcs)"""
    try:
        val = float(val)
        if val % 1 == 0: return f"{val:,.0f}"
        return f"{val:,.2f}"
    except: return "0"

# ==========================================
# 2. PDF ENGINE: PURCHASE ORDER
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

def generate_po_pdf(po_data, po_number):
    pdf = ServeonePO()
    pdf.alias_nb_pages()
    pdf.add_page()
    info = po_data.iloc[0]
    vendor_name = clean_val(info.get('Vendor Name', 'Unknown Vendor'))

    pdf.set_font('Helvetica', 'B', 10)
    pdf.cell(30, 7, "PO YUPI :")
    pdf.set_font('Helvetica', '', 10)
    pdf.cell(0, 7, clean_val(po_number), new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.set_font('Helvetica', 'B', 10)
    pdf.cell(30, 7, "PO Date :")
    pdf.set_font('Helvetica', '', 10)
    pdf.cell(0, 7, date.today().strftime("%d/%m/%Y"), new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.ln(5)

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

    site_raw = clean_val(info.get('SITE (IDN/KRG)', info.get('SITE'))).upper()
    site_label = "IDN" if 'IDN' in site_raw else "KRG"
    full_addr = "Jl. Pancasila IV, Desa/Kelurahan Cicadas, Kec Gunung Putri, Kab Bogor, Provinsi Jawa Barat, Kode Pos 16964 - Indonesia" if site_label == "IDN" else "Jl. Grompol Jambangan Km 5, Muringan, Desa Kaliwuluh, Kecamatan Kebak Kramat Kabupaten Karanganyar Provinsi Jawa Tengah, Indonesia"
    
    pdf.set_font('Helvetica', 'B', 9)
    pdf.cell(0, 6, f"DELIVERY ADDRESS: PT. YUPI INDO JELLY GUM Tbk ({site_label})", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.set_font('Helvetica', '', 8)
    pdf.multi_cell(0, 5, full_addr)
    pdf.ln(5)

    pdf.set_fill_color(240, 240, 240)
    pdf.set_font('Helvetica', 'B', 7)
    cols = [("Deliv Req", 20), ("SERVEONE PO", 25), ("Item Code", 15), ("Item Name & Spec", 55), ("Qty", 15), ("Unit", 10), ("Price", 25), ("Amount", 25)]
    for txt, w in cols:
        pdf.cell(w, 8, txt, border=1, fill=True, align='C')
    pdf.ln()

    pdf.set_font('Helvetica', '', 7)
    total_dpp = 0
    remark_list = []

    for _, row in po_data.iterrows():
        name_spec = f"{clean_val(row.get('Item name'))}\n{clean_val(row.get('Spec'))}"
        qty = safe_float(row.get("Ord. Q'ty"))
        price = safe_float(row.get('PURCHASE PRICE'))
        
        # LOGIKA PERBAIKAN: Memaksa Amount = Qty * Price
        amount = qty * price
        total_dpp += amount
        if clean_val(row.get('REMARK YUPI')): remark_list.append(clean_val(row.get('REMARK YUPI')))

        lines = pdf.multi_cell(55, 4, name_spec, dry_run=True, output="LINES")
        row_h = max(10, len(lines) * 4 + 2)

        if pdf.get_y() + row_h > 240: pdf.add_page()
        x, y = pdf.get_x(), pdf.get_y()
        
        pdf.cell(20, row_h, clean_val(row.get('Req. Dlv Date')), border=1, align='C')
        pdf.cell(25, row_h, clean_val(row.get('PO SEMENTARA')), border=1, align='C')
        pdf.cell(15, row_h, clean_val(row.get('Item Yupi')), border=1, align='C')
        pdf.cell(55, row_h, "", border=1) 
        pdf.cell(15, row_h, format_qty(qty), border=1, align='C')
        pdf.cell(10, row_h, clean_val(row.get('Unit')), border=1, align='C')
        pdf.cell(25, row_h, format_currency(price), border=1, align='R')
        pdf.cell(25, row_h, format_currency(amount), border=1, align='R')

        pdf.set_xy(x + 60, y + 1) 
        pdf.multi_cell(55, 4, name_spec, border=0, align='L')
        pdf.set_xy(10, y + row_h)

    # LOGIKA MATA UANG & PAJAK 
    curr_raw = clean_val(info.get('CURRENCY')).upper()
    if not curr_raw or curr_raw == 'NAN': curr_raw = 'IDR'
    
    tax_raw = clean_val(info.get('TAX TYPE', '')).upper()
    ppn_rate = 0.0 if tax_raw == 'FREE' else 0.11
    
    ppn = total_dpp * ppn_rate
    grand_total = total_dpp + ppn
    
    lbl_tot = f"TOTAL AMOUNT ({curr_raw})"
    lbl_ppn = f"PPN 11% ({curr_raw})"
    lbl_grd = f"GRAND TOTAL ({curr_raw})"

    pdf.set_font('Helvetica', 'B', 8)
    for label, val in [(lbl_tot, total_dpp), (lbl_ppn, ppn), (lbl_grd, grand_total)]:
        # PERBAIKAN ALIGNMENT: 140 Sejajar dengan sisi kiri Header 'Unit'
        pdf.set_x(140)
        # 35 (Lebar kolom Unit + Price), 25 (Lebar kolom Amount)
        pdf.cell(35, 5, label, border=1, align='R')
        pdf.cell(25, 5, format_currency(val), border=1, align='R', new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    pdf.ln(5)
    pdf.set_font('Helvetica', 'B', 9)
    pdf.cell(0, 6, "REMARKS:", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.set_font('Helvetica', '', 8)
    pdf.multi_cell(130, 5, ", ".join(set(remark_list)) if remark_list else "-")

    return bytes(pdf.output()), site_label, vendor_name

# ==========================================
# 3. PDF ENGINE: DELIVERY NOTE
# ==========================================
class ServeoneDN(FPDF):
    def header(self):
        if os.path.exists("logo.png"): self.image("logo.png", 10, 2, 45)
        self.set_font('Helvetica', 'B', 16)
        self.set_y(10)
        self.cell(0, 10, 'DELIVERY NOTE', align='C', new_x=XPos.LMARGIN, new_y=YPos.TOP)
        self.ln(15)
    def footer(self):
        self.set_y(-15)
        self.set_font('Helvetica', 'I', 8)
        self.cell(0, 10, f'Page {self.page_no()} of {{nb}}', align='C')

def generate_dn_pdf(po_data, po_number, dn_vendor, no_pol, packing_method, dn_memo):
    pdf = ServeoneDN()
    pdf.alias_nb_pages()
    pdf.add_page()
    info = po_data.iloc[0]
    
    site_raw = clean_val(info.get('SITE (IDN/KRG)', info.get('SITE'))).upper()
    site_label = "IDN" if 'IDN' in site_raw else "KRG"
    full_addr = "Jl. Pancasila IV, Desa/Kelurahan Cicadas, Kec Gunung Putri, Kab Bogor, Provinsi Jawa Barat, Kode Pos 16964 - Indonesia" if site_label == "IDN" else "Jl. Grompol-Jambangan, Kebakkramat, Karanganyar"
    vendor_name = clean_val(info.get('Vendor Name', 'Unknown Vendor'))

    pdf.set_font('Helvetica', 'B', 9)
    pdf.cell(0, 5, "Bill To / Ship To:", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.set_font('Helvetica', '', 9)
    pdf.cell(0, 5, f"PT. YUPI INDO JELLY GUM Tbk ({site_label})", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(0, 5, full_addr, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.ln(5)

    # PERBAIKAN: Nomor DN Auto Create by System
    svo_dn_number = f"SVO-DN-{date.today().strftime('%Y%m%d')}-{str(po_number)[-4:]}"

    pdf.set_font('Helvetica', '', 9)
    infos = [
        ("PO Yupi", f": {po_number}"),
        ("DN Serveone", f": {svo_dn_number}"),
        ("DN Vendor", f": {dn_vendor if dn_vendor else '_________________'}"),
        ("Delivery Date", f": {date.today().strftime('%d-%m-%Y')}"),
        ("No.POL/Kend", f": {no_pol if no_pol else '_________________'}")
    ]
    for lbl, val in infos:
        pdf.cell(30, 5, lbl)
        pdf.cell(0, 5, val, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.ln(5)

    pdf.set_fill_color(240, 240, 240)
    pdf.set_font('Helvetica', 'B', 8)
    
    # PERBAIKAN TABEL DN: Kolom SO Serveone ditambahkan (Total Lebar: 190)
    cols = [("Item ID", 20), ("SO Serveone", 25), ("Prod. Nm / Spec", 50), ("QTY", 15), ("Unit", 12), ("Packing", 15), ("Req Date", 18), ("Vendor", 35)]
    for txt, w in cols:
        pdf.cell(w, 8, txt, border=1, fill=True, align='C')
    pdf.ln()

    pdf.set_font('Helvetica', '', 8)
    for _, row in po_data.iterrows():
        name_spec = f"{clean_val(row.get('Item name'))}\n{clean_val(row.get('Spec'))}"
        qty = safe_float(row.get("Ord. Q'ty"))
        
        lines = pdf.multi_cell(50, 4, name_spec, dry_run=True, output="LINES")
        row_h = max(10, len(lines) * 4 + 2)

        if pdf.get_y() + row_h > 240: pdf.add_page()
        x, y = pdf.get_x(), pdf.get_y()
        
        pdf.cell(20, row_h, clean_val(row.get('Item Yupi')), border=1, align='C')
        pdf.cell(25, row_h, clean_val(row.get('PO SEMENTARA')), border=1, align='C')
        pdf.cell(50, row_h, "", border=1) 
        pdf.cell(15, row_h, format_qty(qty), border=1, align='C')
        pdf.cell(12, row_h, clean_val(row.get('Unit')), border=1, align='C')
        pdf.cell(15, row_h, packing_method if packing_method else "________", border=1, align='C')
        pdf.cell(18, row_h, clean_val(row.get('Req. Dlv Date')), border=1, align='C')
        
        pdf.set_font('Helvetica', '', 7)
        pdf.cell(35, row_h, vendor_name[:20], border=1, align='C')
        pdf.set_font('Helvetica', '', 8)

        # 45 = Lebar Item ID (20) + SO Serveone (25)
        pdf.set_xy(x + 45, y + 1) 
        pdf.multi_cell(50, 4, name_spec, border=0, align='L')
        pdf.set_xy(10, y + row_h)

    pdf.ln(10)
    pdf.set_font('Helvetica', 'B', 9)
    pdf.cell(0, 5, f"Delivery Memo: {dn_memo if dn_memo else '-'}", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    
    pdf.ln(10)
    pdf.set_font('Helvetica', '', 9)
    pdf.cell(95, 5, "Shipper,", align='C')
    pdf.cell(95, 5, "Received by,", align='C', new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.ln(20)
    pdf.cell(95, 5, "PT Serveone MRO Indonesia", align='C')
    pdf.cell(95, 5, "PT Yupi Indo Jelly Gum Tbk", align='C')

    return bytes(pdf.output()), site_label, vendor_name

# ==========================================
# 4. INTERFACE UTAMA
# ==========================================
with st.sidebar:
    if os.path.exists("logo.png"): st.image("logo.png", use_container_width=True) 
    st.header("⚙️ Setting Akses Tambahan")
    with st.expander("📧 Pengaturan SMTP (Email Sender)", expanded=False):
        st.info("Input App Password jika ingin mengirim email otomatis.")

st.markdown("<h1>SERVEONE PO ENGINE</h1>", unsafe_allow_html=True)
st.markdown("<div class='subtitle'>Integrated ERP System & Document Dispatcher</div>", unsafe_allow_html=True)

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
            
            raw_cols = [str(c).replace('\n', ' ').strip() for c in df_raw.iloc[idx]]
            new_cols = []
            seen = {}
            for c in raw_cols:
                if c in seen:
                    seen[c] += 1
                    new_cols.append(f"{c}_{seen[c]}")
                else:
                    seen[c] = 0
                    new_cols.append(c)
            df.columns = new_cols
            df = df.dropna(subset=['PO YUPI'])
            
            df['Vendor Name'] = df.get('Vendor Name', pd.Series(['Unknown']*len(df))).fillna('Unknown').astype(str).str.strip()
            for col in ['PURCHASE PRICE', "Ord. Q'ty", 'AMOUNT', 'PURCHASE AMOUNT']:
                if col in df.columns: df[col] = df[col].apply(safe_float)

            # --- PANEL FILTER UTAMA ---
            st.markdown("### 🎛️ Data Control Panel")
            col_filter1, col_filter2, col_filter3 = st.columns([1, 1.5, 1])
            with col_filter1:
                date_cols = [c for c in df.columns if 'PO DATE' in c.upper()]
                date_col_name = date_cols[0] if date_cols else None
                if date_col_name:
                    available_dates = sorted(df[date_col_name].dropna().astype(str).unique())
                    selected_date = st.selectbox("📅 Filter by PO Date:", ["Semua Tanggal"] + available_dates)
                    if selected_date != "Semua Tanggal":
                        df = df[df[date_col_name] == selected_date]

            with col_filter3:
                st.markdown("<br>", unsafe_allow_html=True)
                exclude_unready = st.checkbox("🚫 Sembunyikan Item Belum Siap", value=True)
                if exclude_unready:
                    df = df[~df['Vendor Name'].str.contains('SERVEONE|SERVE ONE', case=False, na=False)]
                    if 'PURCHASE PRICE' in df.columns: df = df[df['PURCHASE PRICE'] > 0]

            with col_filter2:
                all_pos = sorted(df['PO YUPI'].unique().tolist())
                selected_pos = st.multiselect("🎯 Pilih Nomor PO:", all_pos, default=all_pos)

            if not selected_pos:
                st.warning("Data kosong / Tidak ada PO terpilih.")
                st.stop()

            final_data = df[df['PO YUPI'].isin(selected_pos)]

            # METRIK DASHBOARD MINI
            col_met1, col_met2, col_met3 = st.columns(3)
            total_documents = len(final_data.groupby(['PO YUPI', 'Vendor Name']))
            col_met1.metric("Dokumen Dipecah (Split)", total_documents)
            col_met2.metric("Total Item Baris", len(final_data))
            
            estimasi_amount = sum((final_data["Ord. Q'ty"] * final_data['PURCHASE PRICE']).fillna(0))
            col_met3.metric("Estimasi DPP Terpilih", f"IDR {estimasi_amount:,.0f}")
            
            st.dataframe(final_data, width="stretch", height=200)
            st.divider()

            # ==========================================
            # PERBAIKAN TABS
            # ==========================================
            tab_engine, tab_dn, tab_dashboard = st.tabs(["📄 PO ENGINE", "🚚 DELIVERY NOTE ENGINE", "📊 ANALYTICS DASHBOARD"])

            # --- TAB 1: PO ENGINE ---
            with tab_engine:
                st.markdown("### ⚙️ Purchase Order Generator")
                col_gen1, col_gen2 = st.columns(2)
                with col_gen1:
                    st.info("🎯 **Single Generate** (Berdasarkan 1 PO)")
                    if len(selected_pos) == 1:
                        po_num = selected_pos[0]
                        po_data_single = final_data[final_data['PO YUPI'] == po_num]
                        unique_vendors = po_data_single['Vendor Name'].unique()
                        
                        for v_name in unique_vendors:
                            safe_vname = re.sub(r'[^a-zA-Z0-9 \.\-]', '', v_name).strip()
                            po_data_vendor = po_data_single[po_data_single['Vendor Name'] == v_name]
                            
                            st.write(f"**Vendor: {safe_vname}**")
                            po_bytes, site, vendor = generate_po_pdf(po_data_vendor, po_num)
                            st.markdown("<div class='btn-po'>", unsafe_allow_html=True)
                            st.download_button(f"📄 Download PO {po_num}", data=po_bytes, file_name=f"[{site}] PO - {po_num} - {safe_vname}.pdf", mime="application/pdf", key=f"po_{safe_vname}")
                            st.markdown("</div><hr style='margin:10px 0'>", unsafe_allow_html=True)
                    else:
                        st.write("Silakan pilih tepat 1 PO di menu Control Panel untuk mengaktifkan fitur ini.")

                with col_gen2:
                    st.info("📦 **Bulk Generate** (Menggabungkan Semua PO)")
                    st.write("Sistem akan otomatis memecah PO berdasarkan nama Vendor dan menjadikannya File ZIP.")
                    if st.button("📦 Download Semua PO (ZIP)"):
                        zip_buffer = io.BytesIO()
                        with st.spinner("Memproses semua PO..."):
                            with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                                for (po_num, v_name), group_data in final_data.groupby(['PO YUPI', 'Vendor Name']):
                                    po_bytes, site, vendor = generate_po_pdf(group_data, po_num)
                                    safe_vname = re.sub(r'[^a-zA-Z0-9 \.\-]', '', vendor).strip()
                                    zip_file.writestr(f"[{site}] PO - {po_num} - {safe_vname}.pdf", po_bytes)
                            st.success("File PO ZIP siap diunduh!")
                            st.download_button("📥 Klik Disini untuk Menyimpan ZIP", data=zip_buffer.getvalue(), file_name=f"BULK_PO_{date.today()}.zip", mime="application/zip")

            # --- TAB 2: DELIVERY NOTE ENGINE ---
            with tab_dn:
                st.markdown("### 🚚 Delivery Note (Surat Jalan)")
                
                col_dn1, col_dn2 = st.columns([1.2, 1])
                with col_dn1:
                    st.info("📝 **Delivery Note Settings** (Input Manual)")
                    with st.container(border=True):
                        dn_vendor = st.text_input("DN Vendor (Opsional):", placeholder="Misal: INV-9921")
                        no_pol = st.text_input("No. Polisi / Kendaraan:", placeholder="Misal: B 1234 XYZ")
                        packing_method = st.text_input("Packing Method:", placeholder="Misal: Pallet / Dus / Pcs")
                        dn_memo = st.text_area("Delivery Memo (Opsional):", placeholder="Instruksi tambahan pengiriman...")

                with col_dn2:
                    st.info("⚙️ **Generate Action**")
                    if len(selected_pos) == 1:
                        po_num = selected_pos[0]
                        po_data_single = final_data[final_data['PO YUPI'] == po_num]
                        for v_name in po_data_single['Vendor Name'].unique():
                            safe_vname = re.sub(r'[^a-zA-Z0-9 \.\-]', '', v_name).strip()
                            po_data_vendor = po_data_single[po_data_single['Vendor Name'] == v_name]
                            
                            dn_bytes, site, vendor = generate_dn_pdf(po_data_vendor, po_num, dn_vendor, no_pol, packing_method, dn_memo)
                            st.markdown("<div class='btn-dn'>", unsafe_allow_html=True)
                            st.download_button(f"🚚 Download DN [{safe_vname}]", data=dn_bytes, file_name=f"[{site}] DN - {po_num} - {safe_vname}.pdf", mime="application/pdf", key=f"dn_sgl_{safe_vname}")
                            st.markdown("</div><hr style='margin:10px 0'>", unsafe_allow_html=True)
                    else:
                        st.write("⬇️ Unduh seluruh DN secara massal (Bulk):")
                        
                    st.markdown("<br>", unsafe_allow_html=True)
                    if st.button("📦 Download Semua DN (ZIP)"):
                        zip_buffer = io.BytesIO()
                        with st.spinner("Memproses semua Delivery Note..."):
                            with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                                for (po_num, v_name), group_data in final_data.groupby(['PO YUPI', 'Vendor Name']):
                                    dn_bytes, site, vendor = generate_dn_pdf(group_data, po_num, dn_vendor, no_pol, packing_method, dn_memo)
                                    safe_vname = re.sub(r'[^a-zA-Z0-9 \.\-]', '', vendor).strip()
                                    zip_file.writestr(f"[{site}] DN - {po_num} - {safe_vname}.pdf", dn_bytes)
                            st.success("File DN ZIP siap!")
                            st.download_button("📥 Simpan DN Massal", data=zip_buffer.getvalue(), file_name=f"BULK_DN_{date.today()}.zip", mime="application/zip")

            # --- TAB 3: DASHBOARD ---
            with tab_dashboard:
                st.header("📈 Dashboard Analitik Pengeluaran")
                df_chart = final_data.copy()
                amount_col = 'PURCHASE AMOUNT' if 'PURCHASE AMOUNT' in df_chart.columns else 'AMOUNT'
                if amount_col in df_chart.columns:
                    col_chart1, col_chart2 = st.columns(2)
                    with col_chart1:
                        st.subheader("🏆 Top 10 Vendor")
                        st.bar_chart(df_chart.groupby('Vendor Name')[amount_col].sum().sort_values(ascending=False).head(10))
                    with col_chart2:
                        st.subheader("🏢 Distribusi per Site")
                        site_col = [c for c in df_chart.columns if 'SITE' in c.upper()][0]
                        df_chart['SITE_CODE'] = df_chart[site_col].astype(str).apply(lambda x: "IDN" if "IDN" in x.upper() else "KRG")
                        st.bar_chart(df_chart.groupby('SITE_CODE')[amount_col].sum())

    except Exception as e:
        st.error(f"Sistem mendeteksi kesalahan: {e}")