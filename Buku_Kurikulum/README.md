# Buku_Kurikulum

Sistem audit sederhana untuk mengecek ada/tidaknya bagian pada dokumen Buku Kurikulum berdasarkan `config/RULES.xlsx`.

## Struktur Folder

```text
Buku_Kurikulum/
├── config/
│   └── RULES.xlsx
├── debug/
│   ├── extracted_toc.json
│   ├── rules_loaded.json
│   └── validation_result.json
├── input/
│   └── input.pdf
├── output/
│   └── laporan.pdf
├── PDF/
│   └── pdfgenerator.py
├── requirements.txt
└── validate_buku_kurikulum.py
```

## Cara Pakai

```bash
pip install -r requirements.txt
python validate_buku_kurikulum.py
```

Atau dengan path manual:

```bash
python validate_buku_kurikulum.py --pdf input/input.pdf --rules config/RULES.xlsx --output output/laporan.pdf
```

## Catatan Sistem

- Warna hijau pada kolom `FIELD` di `RULES.xlsx` dibaca sebagai wajib.
- Warna kuning pada kolom `FIELD` di `RULES.xlsx` dibaca sebagai opsional.
- Bagian opsional yang tidak ditemukan tidak ditampilkan pada laporan PDF.
- Header dan sub-header diambil dari `DAFTAR ISI`.
- Jika `DAFTAR ISI` tidak ditemukan, sistem akan memberi status bahwa daftar isi tidak ditemukan.
- Sistem hanya mengecek keberadaan bagian, bukan kebenaran isi.
- Header/sub-header tambahan yang ada di Daftar Isi tetapi tidak ada di rules akan ditampilkan sebagai temuan tambahan.
