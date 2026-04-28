# PDF to Excel Converter (Bank Statements)

Script Python untuk mengkonversi mutasi rekening PDF dari bank BNI menjadi file Excel `.xlsx` yang sudah terstruktur dan siap dianalisis.

## Fitur Utama

- Extract data dari PDF menggunakan `tabula-py`
- Handle multiline transaction
- Pisahkan Debit/DB dan Credit/CR
- Bersihkan format angka
- Pisahkan summary rekening
- Export ke Excel dengan 2 sheet:
  - `Transaksi`
  - `Summary`
- Auto adjust column width di Excel

## Struktur Project

```text
project/
│
├── pdf_file/              # folder file PDF input
├── excel_file/            # folder output Excel
│
├── convert_mutasi_bni.py
│
└── README.md
