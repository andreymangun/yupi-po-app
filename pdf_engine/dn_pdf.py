from __future__ import annotations
import os
from datetime import date
from fpdf import FPDF
from fpdf.enums import XPos, YPos
from utils.formatters import clean_val, format_qty, safe_float

class ServeoneDN(FPDF):
    def header(self):
        if os.path.exists("logo.png"):
            self.image("logo.png", 10, 2, 45)
        self.set_font("Helvetica", "B", 18)
        self.set_y(12)
        self.cell(0, 10, "DELIVERY NOTE", align="C", new_x=XPos.LMARGIN, new_y=YPos.TOP)
        self.ln(18)

    # --- TAMBAHKAN FUNGSI FOOTER INI ---
    def footer(self):
        self.set_y(-15)
        self.set_font("Helvetica", "I", 8)
        self.set_text_color(128, 128, 128)
        # Mencetak "Page X of Y"
        self.cell(0, 10, f"Page {self.page_no()} of {{nb}}", align="C")
        self.set_text_color(0, 0, 0)

def generate_dn_pdf(po_data, po_number: str, dn_vendor: str, no_pol: str, dn_remarks: str, dn_serveone_number: str, category: str = "SP"):
    pdf = ServeoneDN()
    # Wajib panggil alias_nb_pages untuk perhitungan total halaman
    pdf.alias_nb_pages()
    pdf.add_page()
    
    info = po_data.iloc[0]

    site_raw = clean_val(info.get("SITE (IDN/KRG)", info.get("SITE"))).upper()
    site_label = "IDN" if "IDN" in site_raw else "KRG"
    full_addr = "Jl. Pancasila IV, Desa/Kelurahan Cicadas,\nKec Gunung Putri, Kab Bogor,\nProvinsi Jawa Barat, Indonesia" if site_label == "IDN" else "Jl. Grompol Jambangan Km 5, Muringan,\nDesa Kaliwuluh, Kecamatan Kebak Kramat\nKabupaten Karanganyar Provinsi Jawa Tengah, Indonesia"
    vendor_name = clean_val(info.get("Vendor Name", "Unknown Vendor"))

    y_anchor = pdf.get_y()
    
    # Bill To
    pdf.set_font("Helvetica", "B", 9)
    pdf.cell(90, 5, "Ship To:", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.set_font("Helvetica", "", 9)
    pdf.cell(90, 5, f"PT. YUPI INDO JELLY GUM Tbk ({site_label})", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.multi_cell(90, 5, full_addr)
    
    infos = [
        ("PO Yupi", po_number),
        ("DN ServeOne", dn_serveone_number),
        ("DN Vendor", dn_vendor if dn_vendor else '-'),
        ("Delivery Date", date.today().strftime('%d-%m-%Y')),
        ("No.POL/Kend", no_pol if no_pol else '-'),
        ("Vendor", vendor_name),
    ]
    curr_y = y_anchor
    pdf.set_font("Helvetica", "", 9)
    for lbl, val in infos:
        pdf.set_xy(110, curr_y)
        pdf.cell(28, 5, lbl)
        pdf.cell(2, 5, ":")
        pdf.set_xy(140, curr_y)
        pdf.multi_cell(60, 5, val)
        curr_y = pdf.get_y()
    
    pdf.set_y(max(pdf.get_y(), y_anchor + 35) + 5)

    pdf.set_fill_color(139, 0, 0)
    pdf.set_text_color(255, 255, 255)
    pdf.set_font("Helvetica", "B", 7)
    
    # --- PERBAIKAN TOTAL LEBAR KOLOM (Total ~190) ---
    if category == "RM":
        cols = [("NO", 6), ("SO", 25), ("ITEM ID", 18), ("PROD. NM / SPEC", 40), ("QTY", 10), ("UNIT", 9), ("NO BATCH", 18), ("JML BATCH", 15), ("EXP DATE", 16), ("REMARKS", 33)]
    elif category == "PM":
        cols = [("NO", 8), ("SO", 25), ("ITEM ID", 18), ("PROD. NM / SPEC", 53), ("QTY", 12), ("UNIT", 12), ("CODING", 20), ("REMARKS", 42)]
    else:
        cols = [("NO", 8), ("SO", 25), ("ITEM ID", 20), ("PROD. NM / SPEC", 70), ("QTY", 15), ("UNIT", 15), ("REMARKS", 37)]

    for txt, w in cols: 
        pdf.cell(w, 8, txt, border=1, fill=True, align="C")
    pdf.ln()

    pdf.set_text_color(0, 0, 0)
    pdf.set_font("Helvetica", "", 7)

    qty_col_name = None
    for c in po_data.columns:
        if 'qty' in c.lower() or "q'ty" in c.lower():
            qty_col_name = c; break

    for i, (_, row) in enumerate(po_data.iterrows(), start=1):
        prod_spec = f"{clean_val(row.get('Item name'))}\n{clean_val(row.get('Spec'))}"
        
        qty_raw_val = row.get(qty_col_name, row.get("Ord. Q'ty"))
        qty = safe_float(qty_raw_val)
        
        item_remark = str(row.get("Catatan / Remark", "")).strip()
        if item_remark.lower() in ["nan", "none", ""]: item_remark = "-"
        
        # Tarik data PO SEMENTARA
        po_sementara = str(row.get("PO SEMENTARA", "")).strip()
        if po_sementara.lower() in ["nan", "none", ""]: po_sementara = "-"

        # --- PERBAIKAN: Penyesuaian lebar data lurus dengan header ---
        if category == "RM":
            batch = str(row.get("No Batch", "")).strip()
            jml = str(row.get("Jml Batch", "")).strip()
            exp = str(row.get("Expired Date", "")).strip()
            if batch.lower() in ["nan", "none", ""]: batch = "-"
            if jml.lower() in ["nan", "none", ""]: jml = "-"
            if exp.lower() in ["nan", "none", ""]: exp = "-"
            widths = [6, 25, 18, 40, 10, 9, 18, 15, 16, 33]
            texts = [str(i), po_sementara, clean_val(row.get("Item Yupi")), prod_spec, format_qty(qty), clean_val(row.get("Unit")), batch, jml, exp, item_remark]
        elif category == "PM":
            coding_val = str(row.get("Coding", "")).strip()
            if coding_val.lower() in ["nan", "none", ""]: coding_val = "-"
            widths = [8, 25, 18, 53, 12, 12, 20, 42]
            texts = [str(i), po_sementara, clean_val(row.get("Item Yupi")), prod_spec, format_qty(qty), clean_val(row.get("Unit")), coding_val, item_remark]
        else:
            widths = [8, 25, 20, 70, 15, 15, 37]
            texts = [str(i), po_sementara, clean_val(row.get("Item Yupi")), prod_spec, format_qty(qty), clean_val(row.get("Unit")), item_remark]

        max_lines = 1
        for w, t in zip(widths, texts):
            lines = pdf.multi_cell(w, 4, t, dry_run=True, output="LINES")
            if len(lines) > max_lines: max_lines = len(lines)

        row_h = max(10, max_lines * 4 + 2)
        
        if pdf.get_y() + row_h > 270: 
             pdf.add_page()
        
        x_start, y_start = pdf.get_x(), pdf.get_y()

        for w, t in zip(widths, texts):
            pdf.set_xy(x_start, y_start)
            pdf.cell(w, row_h, "", border=1)
            pdf.set_xy(x_start, y_start + 1)
            
            # PROD NM dan REMARKS lurus kiri (L), sisanya tengah (C)
            align = "L" if t in [prod_spec, item_remark] else "C"
            pdf.multi_cell(w, 4, t, border=0, align=align)
            x_start += w

        pdf.set_xy(10, y_start + row_h)

    # Memaksa posisi Y untuk kotak TTD agar selalu konsisten
    sig_y = 230 
    
    # Jika sisa ruang tidak cukup untuk blok TTD, pindahkan ke halaman baru
    if pdf.get_y() > 210: 
        pdf.add_page()
        sig_y = 230
        
    pdf.set_y(sig_y)
    
    box_w = 65
    
    # KOTAK KIRI (Shipper)
    pdf.set_xy(15, sig_y)
    pdf.set_font("Helvetica", "B", 9)
    pdf.cell(box_w, 8, "Shipper", border=1, align="C") 
    
    pdf.set_xy(15, sig_y + 8)
    pdf.cell(box_w, 25, "", border="LR") 
    
    pdf.set_xy(15, sig_y + 33)
    pdf.cell(box_w, 8, "PT Serveone MRO Indonesia", border=1, align="C") 

    # KOTAK KANAN (Received By)
    x_right = 130
    
    pdf.set_xy(x_right, sig_y)
    pdf.set_font("Helvetica", "B", 9)
    pdf.cell(box_w, 8, "Received By", border=1, align="C")
    
    pdf.set_xy(x_right, sig_y + 8)
    pdf.cell(box_w, 25, "", border="LR")
    
    pdf.set_xy(x_right, sig_y + 33)
    pdf.cell(box_w, 8, "PT Yupi Indo Jelly Gum Tbk", border=1, align="C")

    return bytes(pdf.output()), site_label, vendor_name