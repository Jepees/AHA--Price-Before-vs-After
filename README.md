# 📊 Analisa Harga Shopee

Tool web untuk menganalisa top produk dan membandingkan perubahan harga antar bulan dari data pesanan Shopee.

## 🌐 Cara Akses

Buka link berikut di browser:
```
https://[nama-app].streamlit.app
```

## 📁 Format Nama File

Sebelum upload, **tambahkan kode brand di depan nama file** yang diunduh dari Shopee:

| Asli dari Shopee | Setelah di-rename |
|---|---|
| `Order.all.20260201_20260228.xlsx` | `GBU_Order.all.20260201_20260228.xlsx` |
| `Order.all.20260101_20260131.zip` | `GBU_Order.all.20260101_20260131.zip` |

> **Catatan:** Cukup tambahkan `KodeBrand_` di awal nama file. Sistem akan otomatis mengenali periode dan kode brand dari nama file tersebut.

File yang sudah di-rename ke format standar (misal `GBU SHO (1-28 Feb 26).xlsx`) juga tetap bisa digunakan.

## 📖 Cara Penggunaan

### Mode 1: 🔄 Bandingkan Harga Before vs After

Membandingkan harga produk antara **2 bulan** yang berbeda.

1. Pilih tab **"🔄 Bandingkan Harga Before vs After"**
2. Upload **File BEFORE** (bulan acuan, misal Januari)
3. Upload **File AFTER** (bulan pembanding, misal Februari)
4. Pilih **Urutkan berdasarkan**: Omzet atau Qty
5. Atur **Top produk (%)**: persentase produk teratas yang ingin ditampilkan
6. Tabel akan otomatis muncul menampilkan:
   - SKU dan Nama Produk
   - Total Omzet / Qty Terjual (sesuai pilihan)
   - Rata-rata harga jual di bulan BEFORE dan AFTER
   - Persentase perubahan harga tiap produk
   - Produk yang tidak ada di bulan AFTER ditandai **"tidak tersedia"** (merah)
7. Gunakan slider atau ubah sorting — **tabel langsung berubah tanpa proses ulang**
8. Klik **📥 Download CSV** atau **📥 Download Excel (XLSX)** untuk mengunduh hasilnya

### Mode 2: 📋 Top Produk (1 File)

Melihat produk terlaris dari **1 file** pesanan.

1. Pilih tab **"📋 Top Produk"**
2. Upload file pesanan
3. Pilih **Urutkan berdasarkan**: Omzet atau Qty
4. Atur **Top produk (%)**
5. Tabel top produk akan muncul
6. Download hasil dalam format **CSV** atau **Excel (XLSX)**

## 📁 Format File yang Didukung

| Format | Keterangan |
|--------|------------|
| `.xlsx` | File Excel langsung dari Shopee |
| `.csv` | File CSV (delimiter `;`) |
| `.zip` | File ZIP berisi kumpulan file `.xlsx`/`.csv` |

## ⚠️ Catatan Penting

- File yang di-upload **tidak disimpan di server** — hanya diproses di memori, lalu dihapus
- Pastikan file yang di-upload adalah **file pesanan Shopee** dengan kolom standar
- Untuk mode **Bandingkan Harga**, kedua file harus dari **toko yang sama**
- File Excel yang diunduh sudah memiliki **format angka** (pemisah ribuan & persen) dan bisa langsung dihitung dengan rumus
