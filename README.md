# Surat Penagihan SRR - Generator

Aplikasi Streamlit untuk membuat surat penagihan KJPP Suwendho Rinaldy dan Rekan.

## Fitur

- Ambil data pemberi tugas dari Google Sheets secara otomatis
- Dropdown pemberi tugas (diurutkan A-Z)
- Upload template surat penagihan (.docx)
- Pengisian otomatis semua placeholder `{{...}}`
- Kalkulasi otomatis Fee, DPP, PPN, Jumlah
- Konversi angka ke terbilang Bahasa Indonesia
- Simpan sementara beberapa dokumen sekaligus
- Download individual atau ZIP semua dokumen

## Cara Deploy ke Streamlit Cloud

1. Push folder ini ke GitHub repository
2. Buka https://streamlit.io/cloud
3. New app → pilih repo → set `app.py` sebagai entrypoint
4. Deploy

## Cara Jalankan Lokal

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Placeholder Template

| Placeholder | Sumber |
|---|---|
| `{{Nomor_Srt}}` | Auto: YYMMDD.NNN |
| `{{Kode_PT}}` | Input manual |
| `{{Tgl_Srt}}` | Input tanggal |
| `{{pemberi_tugas}}` | Google Sheets |
| `{{alamat_1}}`, `{{alamat_2}}` | Google Sheets |
| `{{kota}}`, `{{kode_pos}}` | Google Sheets |
| `{{up}}` | Google Sheets |
| `{{penugasan}}` | Google Sheets |
| `{{Tagih_ke}}` / `{{tagih_ke}}` | Input manual |
| `{{no_proposal}}` | Google Sheets |
| `{{tanggal_proposal}}` | Google Sheets |
| `{{proposed_fee}}` | Google Sheets |
| `{{persentase}}` | Input manual |
| `{{Fee_Tagih}}` | Auto: proposed_fee × persentase |
| `{{DPP}}` | Auto: Fee × 11/12 |
| `{{PPN}}` | Auto: 12% × DPP |
| `{{Jumlah}}` | Auto: Fee + PPN |
| `{{Jumlah_Terbilang}}` | Auto: terbilang(Jumlah) |
| `{{Bank}}` | Dropdown |
| `{{Norek}}` | Dropdown |
| `{{title_Up}}` / `{{title_up}}` | Input manual |

## Rekening Bank (dapat disesuaikan di app.py)

- BCA - 747.051.7171
- Bank Mandiri - 122.000.637.7309
- BNI - 0965.7887.32
- BRI - 0059.01.006699.30.7
