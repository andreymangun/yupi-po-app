import pandas as pd

def clean_yupi_data(uploaded_file):
    try:
        # Coba baca dengan encoding standar, jika gagal gunakan latin-1
        try:
            df = pd.read_csv(uploaded_file, encoding='utf-8')
        except:
            df = pd.read_csv(uploaded_file, encoding='ISO-8859-1')

        # SKEPTICAL CHECK: Cari baris header yang mengandung 'PO YUPI'
        # Karena file Yupi biasanya punya judul di baris 1-2
        target_row = None
        for i in range(min(len(df), 10)):
            if 'PO YUPI' in df.iloc[i].values.astype(str):
                target_row = i
                break
        
        if target_row is not None:
            new_header = df.iloc[target_row]
            df = df[target_row + 1:].copy()
            df.columns = new_header

        # Bersihkan nama kolom dari \n dan spasi ganda
        df.columns = [str(c).replace('\n', ' ').replace('  ', ' ').strip() for c in df.columns]

        # Bersihkan kolom angka (menghapus koma, tanda kutip, dll)
        cols_to_fix = ["Ord. Q'ty", "Unit Price", "AMOUNT"]
        for col in cols_to_fix:
            if col in df.columns:
                df[col] = df[col].astype(str).str.replace(',', '').str.replace('"', '').replace('nan', '0')
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
                
        return df
    except Exception as e:
        raise ValueError(f"Gagal memproses CSV: {str(e)}")

def get_po_details(df, po_number):
    return df[df['PO YUPI'] == po_number]