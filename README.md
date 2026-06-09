# SRR Billing Generator

Generator kwitansi & surat penagihan semi-otomatis untuk KJPP Suwendho Rinaldy dan Rekan.

## Fitur
- 🤖 Ekstraksi otomatis dari PDF/DOCX proposal menggunakan Claude AI
- ✏️ Edit manual semua field sebelum generate
- 📊 Output kwitansi (.xlsx) dengan format SRR
- 📝 Output surat tagihan (.docx) dengan placeholder Jinja-style
- 📦 Download ZIP untuk batch (banyak proyek sekaligus)
- 🔢 Nomor surat & kwitansi otomatis (YYMMDD.SEQ/SRR-JK/...)
- 💰 Kalkulasi DPP + PPN 12% otomatis (formula KJPP SRR: DPP = IJ × 11/12)

## Deploy ke Streamlit Cloud

1. Push repo ke GitHub
2. Buka [share.streamlit.io](https://share.streamlit.io)
3. Connect repo, set `app.py` sebagai main file
4. Di **Secrets**, tambahkan:
   ```toml
   ANTHROPIC_API_KEY = "sk-ant-..."
   ```

## Penggunaan Template Custom

### Template Surat Tagihan (.docx)
Buka file .docx Anda di Word, ganti teks yang ingin diisi otomatis dengan placeholder:

| Placeholder | Isi |
|---|---|
| `{{NOMOR_SURAT}}` | Nomor surat otomatis |
| `{{TANGGAL_SURAT}}` | Tanggal tagihan (format Indonesia) |
| `{{NAMA_KLIEN}}` | Nama lengkap klien |
| `{{ALAMAT1}}` | Alamat baris 1 |
| `{{ALAMAT2}}` | Alamat baris 2 |
| `{{KOTA_POS}}` | Kota + kode pos |
| `{{UP}}` | Jabatan penerima (Up.) |
| `{{JENIS_PEKERJAAN}}` | Jenis pekerjaan |
| `{{NOMOR_PROPOSAL}}` | Nomor proposal |
| `{{TGL_PROPOSAL}}` | Tanggal proposal |
| `{{TERBILANG}}` | Jumlah terbilang Rupiah |
| `{{TOTAL_ANGKA}}` | Jumlah dalam angka (Rp X.XXX.XXX) |
| `{{RECEIVER}}` | Nama penandatangan |
| `{{NAMA_KLIEN_SINGKAT}}` | Kode/singkatan klien |

### Template Kwitansi (.xlsx)
Upload template .xlsx Anda. Aplikasi akan menulis data ke posisi cell yang sesuai dengan template kwitansi SRR standar.

> **Catatan**: Jika tidak upload template, aplikasi menggunakan template default SRR.

## Struktur Proyek

```
srr_billing_generator/
├── app.py                    # Main Streamlit app
├── requirements.txt
└── .streamlit/
    └── config.toml
```

## Formula Perhitungan
Sesuai pola KJPP SRR (terverifikasi dari contoh kwitansi):
```
DPP PPN  = Imbalan Jasa × (11/12)
PPN 12%  = DPP × 12%
Total    = Imbalan Jasa + PPN 12%
```
