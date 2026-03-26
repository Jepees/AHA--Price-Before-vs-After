import pandas as pd
import numpy as np
import zipfile
import tempfile
import os
import re
import glob
from datetime import datetime

def clean_number(series):
    """
    Membersihkan format angka dari teks menjadi tipe numerik (float).
    Contoh: '31.780' -> 31780.0
    """
    return (
        series
        .astype(str)
        .str.replace("IDR ", "", regex=False)
        .str.replace("Rp", "", regex=False)
        .str.replace(".", "", regex=False)   # Menghapus titik pemisah ribuan
        .str.replace(",", ".", regex=False)  # Mengubah koma menjadi titik desimal
        .str.strip()
        .apply(pd.to_numeric, errors="coerce")
    )

def _baca_file_input(file_path, delimiter=';'):
    """
    Membaca file input dan mengembalikan satu DataFrame.
    Mendukung: .csv, .xlsx, .zip (berisi kumpulan .xlsx/.csv)
    """
    file_path = str(file_path)
    ext = os.path.splitext(file_path)[1].lower()
    
    if ext == '.csv':
        return pd.read_csv(file_path, sep=delimiter, dtype=str)
    
    elif ext == '.xlsx':
        return pd.read_excel(file_path, dtype=str)
    
    elif ext == '.zip':
        dfs = []
        with tempfile.TemporaryDirectory() as tmp_dir:
            with zipfile.ZipFile(file_path, 'r') as zf:
                zf.extractall(tmp_dir)
            
            # Cari semua file xlsx dan csv di dalam arsip (termasuk subfolder)
            for pattern in ['**/*.xlsx', '**/*.csv']:
                for f in glob.glob(os.path.join(tmp_dir, pattern), recursive=True):
                    if f.endswith('.csv'):
                        dfs.append(pd.read_csv(f, sep=delimiter, dtype=str))
                    else:
                        dfs.append(pd.read_excel(f, dtype=str))
        
        if not dfs:
            raise ValueError(f"Tidak ditemukan file .xlsx/.csv di dalam arsip: {file_path}")
        
        return pd.concat(dfs, ignore_index=True)
    
    else:
        raise ValueError(f"Format file tidak didukung: {ext}. Gunakan .csv, .xlsx, atau .zip")

def proses_penjualan_shopee(file_path, delimiter=';'):
    """
    Menghitung Omzet Bersih dan Harga Efektif per SKU dari laporan pesanan Shopee.
    Mendukung input: .csv, .xlsx, .zip (berisi kumpulan .xlsx/.csv)
    Disesuaikan untuk membagi voucher/diskon secara proporsional berdasarkan nominal harga barang,
    bukan sekadar dibagi rata berdasarkan jumlah QTY barang.
    """
    
    df = _baca_file_input(file_path, delimiter=delimiter)
    
    # Membersihkan dan memastikan kolom numerik
    harga_diskon = clean_number(df["Harga Setelah Diskon"]).fillna(0)
    qty = clean_number(df["Jumlah"]).fillna(0)
    
    voucher_penjual = clean_number(df["Voucher Ditanggung Penjual"]).fillna(0)
    cashback_koin = clean_number(df["Cashback Koin"]).fillna(0)
    diskon_shopee = clean_number(df["Diskon Dari Shopee"]).fillna(0)

    # ---------------------------------------------------------
    # TAHAP 1: Hitung Harga Kotor per Baris & Total per Pesanan
    # ---------------------------------------------------------
    # Total kotor untuk SKU di baris ini (sebelum dipotong voucher pesanan)
    df["Total_Harga_Baris_Kotor"] = harga_diskon * qty
    
    # Jumlahkan total kotor seluruh barang dalam 1 No. Pesanan
    df["Total_Belanja_1_Pesanan"] = df.groupby("No. Pesanan")["Total_Harga_Baris_Kotor"].transform("sum")
    
    # ---------------------------------------------------------
    # TAHAP 2: Hitung Porsi/Proporsi masing-masing Baris
    # ---------------------------------------------------------
    # Berapa persen kontribusi harga baris ini terhadap total belanja 1 pesanan?
    df["Proporsi_Baris"] = np.where(
        df["Total_Belanja_1_Pesanan"] > 0, 
        df["Total_Harga_Baris_Kotor"] / df["Total_Belanja_1_Pesanan"], 
        0
    )
    
    # ---------------------------------------------------------
    # TAHAP 3: Hitung Potongan/Voucher secara Proporsional
    # ---------------------------------------------------------
    # Nilai Diskon Level Pesanan (Nilainya diulang di setiap baris oleh Shopee)
    # Catatan: Diskon Shopee ditambahkan karena uangnya dikembalikan oleh Shopee (merupakan hak penjual)
    potongan_level_pesanan = voucher_penjual + cashback_koin - diskon_shopee
    
    # Diskon yang dibebankan HANYA untuk baris ini
    df["Potongan_Proporsional_Baris"] = df["Proporsi_Baris"] * potongan_level_pesanan
    
    # ---------------------------------------------------------
    # TAHAP 4: Hitung Omzet Bersih & Harga Efektif
    # ---------------------------------------------------------
    # Cek apakah pesanan dibatalkan (Harga Setelah Diskon = 0 di laporan Shopee)
    is_cancelled = df["Harga Setelah Diskon"].astype(str).str.strip() == "0"
    
    # Omzet aktual yang didapat dari baris transaksi ini
    df["Omzet_Bersih_Baris"] = np.where(
        is_cancelled,
        0,
        df["Total_Harga_Baris_Kotor"] - df["Potongan_Proporsional_Baris"]
    )
    
    # Harga Efektif per unit untuk baris ini
    df["Harga_Efektif_Per_Unit"] = np.where(
        qty > 0,
        df["Omzet_Bersih_Baris"] / qty,
        np.nan
    )
    
    return df

def rekap_agregasi_produk(df):
    """
    Menjumlahkan seluruh data dari tiap baris menjadi ringkasan level Produk/SKU
    (Cocok untuk ditampilkan di laporan Omzet seperti di Excel)
    """
    # Kolom aggregasi
    df["Jumlah_Valid"] = np.where(df["Omzet_Bersih_Baris"] > 0, clean_number(df["Jumlah"]), 0)
    
    # Groupby SKU saja agar SKU yang sama tapi nama beda tetap digabung
    agregasi = df.groupby("Nomor Referensi SKU").agg(
        Total_Qty_Terjual=("Jumlah_Valid", "sum"),
        Total_Omzet_Produk=("Omzet_Bersih_Baris", "sum")
    ).reset_index()

    # Untuk setiap SKU, ambil Nama Produk yang paling pendek (jumlah karakter terkecil)
    nama_terpendek = (
        df.groupby("Nomor Referensi SKU")["Nama Produk"]
        .apply(lambda names: min(names.unique(), key=len))
        .reset_index()
    )
    agregasi = agregasi.merge(nama_terpendek, on="Nomor Referensi SKU", how="left")
    
    # Singkirkan produk yang Qty 0 (misal karena dibatalkan semua)
    agregasi = agregasi[agregasi["Total_Qty_Terjual"] > 0].copy()
    
    # Cari Rata-Rata Harga Jual Per Unit dari keseluruhan bulan ini
    agregasi["Rata2_Harga_Jual"] = np.where(
        agregasi["Total_Qty_Terjual"] > 0,
        agregasi["Total_Omzet_Produk"] / agregasi["Total_Qty_Terjual"],
        0
    )
    
    # Urutkan berdasarkan Omzet Terbesar (TOP Omzet)
    agregasi.sort_values(by="Total_Omzet_Produk", ascending=False, inplace=True)
    
    # ---------------------------------------------------------
    # FORMAT SESUAI PERMINTAAN
    # ---------------------------------------------------------
    
    # 1. Pilih & Susun Ulang Kolom
    agregasi = agregasi[[
        "Nomor Referensi SKU", 
        "Total_Omzet_Produk", 
        "Nama Produk", 
        "Total_Qty_Terjual", 
        "Rata2_Harga_Jual"
    ]].copy()
    
    # 2. Ubah Nama Kolom agar persis dengan tabel target
    agregasi.rename(columns={
        "Nomor Referensi SKU": "SKU",
        "Total_Omzet_Produk": "Total Omzet per Produk",
        "Total_Qty_Terjual": "Qty Terjual",
        "Rata2_Harga_Jual": "Rata2 Harga Jual"
    }, inplace=True)
    
    # 3. Ubah format angka menjadi pemisah ribuan dengan koma (,)
    agregasi["Qty Terjual"] = agregasi["Qty Terjual"].astype(int)
    agregasi["Total Omzet per Produk"] = agregasi["Total Omzet per Produk"].apply(lambda x: f"{x:,.0f}")
    agregasi["Rata2 Harga Jual"] = agregasi["Rata2 Harga Jual"].apply(lambda x: f"{x:,.0f}")
    
    return agregasi

def top_produk(file_path, urut_berdasarkan="omzet", top_persen=20, delimiter=';'):
    """
    Menampilkan top produk dari satu file pesanan.
    
    Parameters:
        file_path: path ke file (.csv, .xlsx, atau .zip)
        urut_berdasarkan: "omzet" atau "qty" (default: "omzet")
        top_persen: persentase top produk yang ditampilkan (default: 20)
        delimiter: delimiter untuk CSV (default: ';')
    
    Returns:
        DataFrame dengan kolom: SKU, Nama Produk, Total Omzet, Qty Terjual, Rata² Harga Jual {Bulan}
    """
    # 1. Proses file
    df = proses_penjualan_shopee(file_path, delimiter=delimiter)
    
    # 2. Ambil bulan dominan
    bulan = _get_bulan_dominan(df)
    
    # 3. Rekap numerik
    rekap = _rekap_numerik(df)
    
    # 4. Urutkan berdasarkan pilihan user
    urut = urut_berdasarkan.lower().strip()
    if urut == "qty":
        rekap.sort_values(by="Total_Qty_Terjual", ascending=False, inplace=True)
    elif urut == "omzet":
        rekap.sort_values(by="Total_Omzet_Produk", ascending=False, inplace=True)
    else:
        raise ValueError(f"urut_berdasarkan harus 'omzet' atau 'qty', bukan '{urut_berdasarkan}'")
    
    # 5. Hitung jumlah produk berdasarkan persentase
    top_n = max(1, int(len(rekap) * top_persen / 100))
    output = rekap.head(top_n).copy()
    
    # 6. Susun kolom & rename
    col_harga = f"Rata2 Harga Jual {bulan}"
    output = output[["Nomor Referensi SKU", "Nama Produk", "Total_Omzet_Produk",
                      "Total_Qty_Terjual", "Rata2_Harga_Jual"]].copy()
    
    output.rename(columns={
        "Nomor Referensi SKU": "SKU",
        "Total_Omzet_Produk": "Total Omzet per Produk",
        "Total_Qty_Terjual": "Qty Terjual",
        "Rata2_Harga_Jual": col_harga
    }, inplace=True)
    
    # 7. Format angka
    output["Qty Terjual"] = output["Qty Terjual"].astype(int)
    output["Total Omzet per Produk"] = output["Total Omzet per Produk"].apply(lambda x: f"{x:,.0f}")
    output[col_harga] = output[col_harga].apply(lambda x: f"{x:,.0f}")
    
    return output

def _get_bulan_dominan(df):
    """
    Mengambil bulan dan tahun yang paling dominan dari kolom 'Waktu Pesanan Dibuat'.
    Format kolom: 'YYYY-MM-DD HH:MM'
    Return: string seperti 'Feb 2026', 'Jan 2026', dll.
    """
    waktu = pd.to_datetime(df["Waktu Pesanan Dibuat"], errors="coerce")
    # Ambil bulan-tahun yang paling sering muncul
    bulan_tahun = waktu.dt.to_period("M")
    dominan = bulan_tahun.mode()[0]
    
    # Format ke nama bulan singkat + tahun (misal 'Feb 2026')
    bulan_map = {
        1: "Jan", 2: "Feb", 3: "Mar", 4: "Apr", 5: "May", 6: "Jun",
        7: "Jul", 8: "Aug", 9: "Sep", 10: "Oct", 11: "Nov", 12: "Dec"
    }
    return f"{bulan_map[dominan.month]} {dominan.year}"

def _rekap_numerik(df):
    """
    Versi rekap_agregasi_produk yang mengembalikan data NUMERIK (belum di-format).
    Dipakai internal oleh bandingkan_harga_before_after().
    """
    df["Jumlah_Valid"] = np.where(df["Omzet_Bersih_Baris"] > 0, clean_number(df["Jumlah"]), 0)
    
    # Groupby SKU saja
    agregasi = df.groupby("Nomor Referensi SKU").agg(
        Total_Qty_Terjual=("Jumlah_Valid", "sum"),
        Total_Omzet_Produk=("Omzet_Bersih_Baris", "sum")
    ).reset_index()

    # Nama produk terpendek per SKU
    nama_terpendek = (
        df.groupby("Nomor Referensi SKU")["Nama Produk"]
        .apply(lambda names: min(names.unique(), key=len))
        .reset_index()
    )
    agregasi = agregasi.merge(nama_terpendek, on="Nomor Referensi SKU", how="left")
    
    # Singkirkan produk yang Qty 0
    agregasi = agregasi[agregasi["Total_Qty_Terjual"] > 0].copy()
    
    # Harga rata-rata (NUMERIK, belum di-format)
    agregasi["Rata2_Harga_Jual"] = np.where(
        agregasi["Total_Qty_Terjual"] > 0,
        agregasi["Total_Omzet_Produk"] / agregasi["Total_Qty_Terjual"],
        0
    )
    
    # Urutkan berdasarkan Omzet Terbesar
    agregasi.sort_values(by="Total_Omzet_Produk", ascending=False, inplace=True)
    
    return agregasi

def bandingkan_harga_before_after(file_before, file_after, urut_berdasarkan="omzet", top_persen=20, delimiter=';'):
    """
    Membandingkan harga produk antara 2 periode (bulan) yang berbeda.
    
    Parameters:
        file_before: path file acuan (top produk)
        file_after: path file pembanding (hanya dilihat harganya)
        urut_berdasarkan: "omzet" atau "qty" (default: "omzet")
        top_persen: persentase top produk yang ditampilkan (default: 20)
        delimiter: delimiter untuk file CSV (default ';')
    
    Returns:
        DataFrame dengan kolom:
        - SKU, Total Omzet/Qty, Nama Produk
        - Rata² Harga Jual {bulan_before}
        - Rata² Harga Jual {bulan_after}
        - % Perubahan
    """
    # 1. Proses kedua file
    df_before = proses_penjualan_shopee(file_before, delimiter=delimiter)
    df_after = proses_penjualan_shopee(file_after, delimiter=delimiter)
    
    # 2. Ambil nama bulan dominan dari masing-masing file
    bulan_before = _get_bulan_dominan(df_before)
    bulan_after = _get_bulan_dominan(df_after)
    
    # 3. Rekap numerik kedua file
    rekap_before = _rekap_numerik(df_before)
    rekap_after = _rekap_numerik(df_after)
    
    # 4. Urutkan berdasarkan pilihan user
    urut = urut_berdasarkan.lower().strip()
    if urut == "qty":
        rekap_before.sort_values(by="Total_Qty_Terjual", ascending=False, inplace=True)
        col_metrik_src = "Total_Qty_Terjual"
        col_metrik_nama = "Qty Terjual"
    elif urut == "omzet":
        rekap_before.sort_values(by="Total_Omzet_Produk", ascending=False, inplace=True)
        col_metrik_src = "Total_Omzet_Produk"
        col_metrik_nama = "Total Omzet per Produk"
    else:
        raise ValueError(f"urut_berdasarkan harus 'omzet' atau 'qty', bukan '{urut_berdasarkan}'")
    
    # 5. Hitung jumlah produk berdasarkan persentase, ambil top dari BEFORE
    top_n = max(1, int(len(rekap_before) * top_persen / 100))
    top_before = rekap_before.head(top_n).copy()
    
    # 6. Siapkan data AFTER (hanya SKU & harga)
    harga_after = rekap_after[["Nomor Referensi SKU", "Rata2_Harga_Jual"]].copy()
    harga_after.rename(columns={"Rata2_Harga_Jual": "Harga_After"}, inplace=True)
    
    # 7. Left join BEFORE ke AFTER berdasarkan SKU
    hasil = top_before.merge(harga_after, on="Nomor Referensi SKU", how="left")
    
    # 8. Hitung % Perubahan
    hasil["Persen_Perubahan"] = np.where(
        hasil["Harga_After"].notna() & (hasil["Rata2_Harga_Jual"] > 0),
        ((hasil["Harga_After"] - hasil["Rata2_Harga_Jual"]) / hasil["Rata2_Harga_Jual"]) * 100,
        np.nan
    )
    
    # ---------------------------------------------------------
    # FORMAT OUTPUT
    # ---------------------------------------------------------
    
    # Nama kolom dinamis berdasarkan bulan dominan
    col_before = f"Rata2 Harga Jual {bulan_before}"
    col_after = f"Rata2 Harga Jual {bulan_after}"
    
    # Susun kolom output (hanya tampilkan kolom metrik yang dipilih)
    output = hasil[["Nomor Referensi SKU", col_metrik_src, "Nama Produk",
                     "Rata2_Harga_Jual", "Harga_After", "Persen_Perubahan"]].copy()
    
    output.rename(columns={
        "Nomor Referensi SKU": "SKU",
        col_metrik_src: col_metrik_nama,
        "Rata2_Harga_Jual": col_before,
        "Harga_After": col_after,
        "Persen_Perubahan": "% Perubahan"
    }, inplace=True)
    
    # Format angka menjadi pemisah ribuan
    if urut == "qty":
        output[col_metrik_nama] = output[col_metrik_nama].astype(int)
    else:
        output[col_metrik_nama] = output[col_metrik_nama].apply(lambda x: f"{x:,.0f}")
    output[col_before] = output[col_before].apply(lambda x: f"{x:,.0f}")
    
    # Kolom AFTER: "tidak tersedia" jika NaN, format angka jika ada
    output[col_after] = output[col_after].apply(
        lambda x: "tidak tersedia" if pd.isna(x) else f"{x:,.0f}"
    )
    
    # Hitung weighted average % perubahan (berbobot omzet, hanya yang valid)
    mask_valid = hasil["Persen_Perubahan"].notna()
    if mask_valid.any():
        rata2_perubahan = (
            (hasil.loc[mask_valid, "Persen_Perubahan"] * hasil.loc[mask_valid, "Total_Omzet_Produk"]).sum()
            / hasil.loc[mask_valid, "Total_Omzet_Produk"].sum()
        )
    else:
        rata2_perubahan = np.nan
    col_persen = f"% Perubahan ({rata2_perubahan:.2f}%)" if not pd.isna(rata2_perubahan) else "% Perubahan"
    output.rename(columns={"% Perubahan": col_persen}, inplace=True)
    
    # Kolom % Perubahan: format dengan 2 desimal + tanda %, atau "-" jika tidak tersedia
    output[col_persen] = output[col_persen].apply(
        lambda x: "-" if pd.isna(x) else f"{x:.2f}%"
    )
    
    return output

# ---------------------------------------------------------
# PARSE NAMA FILE SHOPEE
# ---------------------------------------------------------
_BULAN_MAP = {1: "Jan", 2: "Feb", 3: "Mar", 4: "Apr", 5: "May", 6: "Jun",
              7: "Jul", 8: "Aug", 9: "Sep", 10: "Oct", 11: "Nov", 12: "Dec"}

def parse_shopee_filename(filename):
    """
    Parse nama file Shopee dan kembalikan (kode_brand, display_name).
    
    Mendukung 2 format:
      1. Format Shopee asli: KodeBrand_Order.all.YYYYMMDD_YYYYMMDD.ext
         → kode_brand = "KodeBrand", display = "KodeBrand SHO (1-28 Feb 26).ext"
      2. Format sudah di-rename: "KodeBrand SHO (1-28 Feb 26).ext"
         → kode_brand = "KodeBrand", display = nama file apa adanya
    """
    # Coba format Shopee asli
    match = re.search(r"^(.+?)_Order\.all\.(\d{8})_(\d{8})", filename)
    if match:
        kode_brand = match.group(1).strip().upper()
        tgl_awal = datetime.strptime(match.group(2), "%Y%m%d").date()
        tgl_akhir = datetime.strptime(match.group(3), "%Y%m%d").date()
        bulan_kode = _BULAN_MAP.get(tgl_akhir.month, "???")
        tahun_2d = str(tgl_akhir.year)[-2:]
        periode = f"({tgl_awal.day}-{tgl_akhir.day} {bulan_kode} {tahun_2d})"
        ext = os.path.splitext(filename)[1]
        display_name = f"{kode_brand} SHO {periode}{ext}"
        return kode_brand, display_name
    
    # Fallback: format lama (spasi sebagai pemisah)
    nama_tanpa_ext = filename.rsplit('.', 1)[0]
    kode_brand = nama_tanpa_ext.split(' ')[0] if ' ' in nama_tanpa_ext else nama_tanpa_ext
    return kode_brand, filename


if __name__ == "__main__":
    # Contoh penggunaan
    file_contoh = "data contoh/GBU SHO (1-28 Feb 26).xlsx"
    
    # 1. Proses file pesanan Shopee   
    df_detail = proses_penjualan_shopee(file_contoh)
    
    # 2. Proses rekap data menjadi bentuk tabel Omzet Produk (seperti di gambarmu)
    df_rekap = rekap_agregasi_produk(df_detail)
    
    print("\n--- Top 5 Produk Berdasarkan Omzet Terbaik ---")
    pd.set_option("display.max_columns", None)
    pd.set_option("display.width", 200)
    print(df_rekap.head(5).to_string(index=False))
