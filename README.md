# 📊 Analisa Harga Shopee

Tool web untuk menganalisa top produk dan membandingkan perubahan harga antar bulan dari data pesanan Shopee.

## 🌐 Cara Akses

Buka link berikut di browser:
```
https://aha-price-before-vs-after-sdhbggurio124.streamlit.app/
```

## 📖 Cara Penggunaan

### Mode 1: 🔄 Bandingkan Harga Before vs After

Membandingkan harga produk antara **2 bulan** yang berbeda.

1. Pilih tab **"🔄 Bandingkan Harga Before vs After"**
2. Pilih **Urutkan berdasarkan**: Omzet atau Qty
3. Atur **Top produk (%)**: persentase produk teratas yang ingin ditampilkan
4. Upload **File BEFORE** (bulan acuan, misal Januari)
5. Upload **File AFTER** (bulan pembanding, misal Februari)
6. Tabel akan otomatis muncul menampilkan:
   - SKU dan Nama Produk
   - Total Omzet / Qty Terjual (sesuai pilihan)
   - Rata-rata harga jual di bulan BEFORE dan AFTER
   - Persentase perubahan harga tiap produk
   - Produk yang tidak ada di bulan AFTER ditandai **"tidak tersedia"** (merah)
7. Gunakan slider atau ubah sorting — **tabel langsung berubah tanpa proses ulang**
8. Klik **📥 Download sebagai CSV** untuk mengunduh hasilnya

### Mode 2: 📋 Top Produk (1 File)

Melihat produk terlaris dari **1 file** pesanan.

1. Pilih tab **"📋 Top Produk"**
2. Pilih **Urutkan berdasarkan**: Omzet atau Qty
3. Atur **Top produk (%)**
4. Upload file pesanan
5. Tabel top produk akan muncul

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
