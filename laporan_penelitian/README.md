# Sistem Audit Laporan Penelitian

Sistem ini mengecek ada atau tidaknya bagian dokumen berdasarkan `config/RULES.xlsx`.

## Struktur Folder

```text
laporan_penelitian/
├── config/
│   └── RULES.xlsx
├── debug/
├── input/
│   └── input.pdf
├── output/
│   └── laporan.pdf
├── PDF/
│   └── pdfgenerator.py
├── requirements.txt
└── validate_laporan_penelitian.py
```

## Format RULES.xlsx

Rules memakai dua kolom:

- `FIELD`: nama bagian yang dicek.
- `Parent`: parent bagian tersebut jika bagian itu adalah sub-header.

Warna pada cell `FIELD`:

- Hijau = wajib ada.
- Kuning = opsional.

Bagian opsional yang tidak ditemukan tidak akan ditampilkan di laporan.

## Cara Run

```bash
cd laporan_penelitian
pip install -r requirements.txt
python validate_laporan_penelitian.py
```

Default input:

```text
input/input.pdf
```

Default output:

```text
output/laporan.pdf
```

Run dengan path manual:

```bash
python validate_laporan_penelitian.py --pdf input/nama_file.pdf --rules config/RULES.xlsx --output output/laporan.pdf
```

## Catatan

- Sistem hanya mengecek keberadaan bagian, bukan menilai kualitas isi.
- Header dan sub-header utama diambil dari `DAFTAR ISI`.
- Jika `DAFTAR ISI` tidak ditemukan, sistem akan memberi status bahwa daftar isi tidak ditemukan.
- Temuan tambahan dari daftar isi tetap ditampilkan sebagai informasi.
