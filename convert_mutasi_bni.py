import re
from pathlib import Path

import pandas as pd
import tabula


# =========================
# PDF CONFIG
# =========================
PDF_AREA = [125, 15, 760, 830] # [top, left, bottom, right]
PDF_COLUMNS = [105, 250, 610, 700] 


# =========================
# HELPER FUNCTIONS
# =========================
def clean_text(value):
    if pd.isna(value):
        return ""

    text = str(value).replace("\r", " ").replace("\n", " ").strip()
    return re.sub(r"\s+", " ", text)


def append_text(base, extra):
    base = clean_text(base)
    extra = clean_text(extra)

    if not extra:
        return base
    if not base:
        return extra

    return f"{base} {extra}"


def parse_number(value):
    """
    Convert:
    '+4,000,000' -> 4000000
    '-96,600'    -> -96600
    '6,119,110'  -> 6119110
    """
    text = clean_text(value).upper()

    if text in ("", "NAN"):
        return pd.NA

    text = text.replace(",", "").replace("+", "").strip()

    try:
        number = float(text)
        if number.is_integer():
            return int(number)
        return number
    except ValueError:
        return pd.NA


def is_date(text):
    """
    Example:
    02 Mar 2026
    """
    text = clean_text(text)
    return bool(re.match(r"^\d{2}\s+[A-Za-z]{3}\s+\d{4}$", text))


def is_time(text):
    """
    Example:
    16:53:00 WIB
    """
    text = clean_text(text)
    return bool(re.match(r"^\d{2}:\d{2}:\d{2}\s+WIB$", text))


def extract_nominal_saldo(text):
    """
    Extract nominal and saldo from text ending with:
    -96,600 3,930,095
    +4,000,000 7,716,795
    """
    text = clean_text(text)

    pattern = r"([+-]?\d[\d,]*(?:\.\d+)?)\s+(\d[\d,]*(?:\.\d+)?)$"
    match = re.search(pattern, text)

    if not match:
        return text, "", ""

    nominal = match.group(1)
    saldo = match.group(2)
    remaining_text = text[:match.start()].strip()

    return remaining_text, nominal, saldo


def split_db_cr(nominal):
    number = parse_number(nominal)

    if pd.isna(number):
        return pd.Series([pd.NA, pd.NA], index=["DB", "CR"])

    if number < 0:
        return pd.Series([abs(number), pd.NA], index=["DB", "CR"])

    return pd.Series([pd.NA, number], index=["DB", "CR"])


# =========================
# READ PDF
# =========================
def read_pdf_table(pdf_file):
    dfs = tabula.read_pdf(
        str(pdf_file),
        pages="all",
        stream=True,
        guess=False,
        area=PDF_AREA,
        columns=PDF_COLUMNS,
        pandas_options={"header": None},
        multiple_tables=False,
    )

    if not dfs:
        raise ValueError("No table found. Try adjusting PDF_AREA / PDF_COLUMNS.")

    df = pd.concat(dfs, ignore_index=True)

    print(f"Jumlah kolom hasil Tabula: {df.shape[1]}")

    rows = []

    for _, row in df.iterrows():
        values = [clean_text(v) for v in row.tolist()]
        values = values + [""] * (5 - len(values))

        tanggal_waktu = values[0]
        rincian = ""
        nominal = ""
        saldo = ""

        if df.shape[1] == 3:
            # usually: tanggal/waktu | rincian | nominal+saldo
            tanggal_waktu = values[0]
            rincian = values[1]
            right_text = values[2]

            remaining, extracted_nominal, extracted_saldo = extract_nominal_saldo(right_text)

            if remaining:
                rincian = append_text(rincian, remaining)

            nominal = extracted_nominal
            saldo = extracted_saldo

            if not nominal and not saldo:
                remaining, extracted_nominal, extracted_saldo = extract_nominal_saldo(rincian)
                rincian = remaining
                nominal = extracted_nominal
                saldo = extracted_saldo

        elif df.shape[1] == 4:
            # usually: tanggal/waktu | rincian | nominal | saldo
            tanggal_waktu = values[0]
            rincian = values[1]
            nominal = values[2]
            saldo = values[3]

            if not nominal and not saldo:
                remaining, extracted_nominal, extracted_saldo = extract_nominal_saldo(rincian)
                rincian = remaining
                nominal = extracted_nominal
                saldo = extracted_saldo

        else:
            # if Tabula returns 5+ columns:
            # tanggal | waktu | rincian | nominal | saldo
            tanggal_waktu = append_text(values[0], values[1])
            rincian = values[2]
            nominal = values[3]
            saldo = values[4]

        rows.append({
            "TanggalWaktu": tanggal_waktu,
            "Rincian": rincian,
            "Nominal": nominal,
            "Saldo": saldo,
        })

    result = pd.DataFrame(rows, columns=["TanggalWaktu", "Rincian", "Nominal", "Saldo"])

    for col in result.columns:
        result[col] = result[col].apply(clean_text)

    return result


# =========================
# CLEANING
# =========================
def remove_garbage_rows(df):
    garbage_regex = (
        r"Laporan Mutasi Rekening|"
        r"Periode:|"
        r"Otoritas Jasa Keuangan|"
        r"peserta penjaminan|"
        r"Tanggal & Waktu|"
        r"Rincian Transaksi|"
        r"Nominal \(IDR\)|"
        r"Saldo \(IDR\)|"
        r"Informasi Lainnya|"
        r"Apabila terdapat kesalahan|"
        r"BNI dapat sewaktu|"
        r"Dokumen ini dibuat|"
        r"^\d+\s+dari\s+\d+$"
    )

    def is_garbage(row):
        joined = clean_text(" ".join(row.astype(str).tolist()))

        if not joined:
            return True

        return bool(re.search(garbage_regex, joined, flags=re.IGNORECASE))

    return df[~df.apply(is_garbage, axis=1)].reset_index(drop=True)


def split_summary_rows(df):
    summary_rows = []
    drop_indexes = set()

    keywords = [
        "Saldo Awal",
        "Total Pemasukan",
        "Total Pengeluaran",
        "Saldo Akhir",
    ]

    for i in range(len(df)):
        row = df.iloc[i]

        cells = [
            clean_text(row["TanggalWaktu"]),
            clean_text(row["Rincian"]),
            clean_text(row["Nominal"]),
            clean_text(row["Saldo"]),
        ]

        joined = clean_text(" ".join([c for c in cells if c]))

        for keyword in keywords:
            if keyword.upper() in joined.upper():
                amount = ""

                # case 1: keyword and amount in the same row
                same_row_text = joined.upper().replace(keyword.upper(), "").strip()
                if same_row_text:
                    amount = same_row_text

                # case 2: amount is in another column on same row
                for cell in cells:
                    if cell != keyword and parse_number(cell) is not pd.NA:
                        amount = cell
                        break

                # case 3: amount is in the next few rows
                if not amount:
                    for j in range(i + 1, min(i + 4, len(df))):
                        next_joined = clean_text(" ".join(df.iloc[j].astype(str).tolist()))
                        parsed = parse_number(next_joined)

                        if not pd.isna(parsed):
                            amount = next_joined
                            drop_indexes.add(j)
                            break

                summary_rows.append({
                    "Keterangan": keyword,
                    "Amount": abs(parse_number(amount)) if not pd.isna(parse_number(amount)) else pd.NA,
                })

                drop_indexes.add(i)
                break

    df_summary = pd.DataFrame(summary_rows, columns=["Keterangan", "Amount"])

    # remove duplicate summary rows, keep first occurrence
    df_summary = df_summary.drop_duplicates(subset=["Keterangan"], keep="first").reset_index(drop=True)

    df_trans = df.drop(index=list(drop_indexes), errors="ignore").reset_index(drop=True)

    # remove table saldo awal / saldo akhir rows if still left
    def is_table_saldo_row(row):
        joined = clean_text(" ".join(row.astype(str).tolist()))
        return bool(
            re.match(
                r"^Saldo Awal\s+[\d,+-]+$|^Saldo Akhir\s+[\d,+-]+$",
                joined,
                flags=re.IGNORECASE
            )
        )

    df_trans = df_trans[~df_trans.apply(is_table_saldo_row, axis=1)].reset_index(drop=True)

    return df_trans, df_summary


def merge_transactions(df):
    """
    BNI format:
    02 Mar 2026
    16:53:00 WIB
    Pembayaran Qris
    ALGO J942 AFM RS POLRI JAKARTA TIMURID  -96,600  3,930,095
    """
    records = []
    current = None

    for _, row in df.iterrows():
        tanggal_waktu = clean_text(row["TanggalWaktu"])
        rincian = clean_text(row["Rincian"])
        nominal = clean_text(row["Nominal"])
        saldo = clean_text(row["Saldo"])

        joined = clean_text(" ".join([tanggal_waktu, rincian, nominal, saldo]))

        if not joined:
            continue

        # start new transaction
        if is_date(tanggal_waktu):
            if current is not None:
                records.append(current)

            current = {
                "Tanggal": tanggal_waktu,
                "Jam": "",
                "Keterangan": "",
                "Mutasi": "",
                "Saldo": "",
            }

            if rincian:
                current["Keterangan"] = append_text(current["Keterangan"], rincian)

            if nominal:
                current["Mutasi"] = nominal

            if saldo:
                current["Saldo"] = saldo

            continue

        if current is None:
            continue

        # time row
        if is_time(tanggal_waktu):
            current["Jam"] = tanggal_waktu

            if rincian:
                current["Keterangan"] = append_text(current["Keterangan"], rincian)

            if nominal:
                current["Mutasi"] = nominal

            if saldo:
                current["Saldo"] = saldo

            continue

        # continuation row
        if tanggal_waktu:
            current["Keterangan"] = append_text(current["Keterangan"], tanggal_waktu)

        if rincian:
            current["Keterangan"] = append_text(current["Keterangan"], rincian)

        if nominal:
            if current["Mutasi"] == "":
                current["Mutasi"] = nominal
            else:
                current["Keterangan"] = append_text(current["Keterangan"], nominal)

        if saldo:
            if current["Saldo"] == "":
                current["Saldo"] = saldo
            else:
                current["Keterangan"] = append_text(current["Keterangan"], saldo)

    if current is not None:
        records.append(current)

    return pd.DataFrame(records, columns=["Tanggal", "Jam", "Keterangan", "Mutasi", "Saldo"])


def finalize_transactions(df):
    df = df.copy()

    df[["DB", "CR"]] = df["Mutasi"].apply(split_db_cr)

    df["Mutasi"] = df["Mutasi"].apply(
        lambda x: abs(parse_number(x)) if not pd.isna(parse_number(x)) else pd.NA
    )

    df["Saldo"] = df["Saldo"].apply(parse_number)

    df = df[["Tanggal", "Jam", "Keterangan", "Mutasi", "DB", "CR", "Saldo"]].copy()

    return df


# =========================
# EXPORT
# =========================
def auto_fit_columns(sheet):
    for col_cells in sheet.columns:
        col_letter = col_cells[0].column_letter
        max_length = max(
            len(str(cell.value)) if cell.value is not None else 0
            for cell in col_cells
        )
        sheet.column_dimensions[col_letter].width = max_length + 2


def export_to_excel(df_final, df_summary, output_file):
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        df_final.to_excel(writer, sheet_name="Transaksi", index=False)
        df_summary.to_excel(writer, sheet_name="Summary", index=False)

        auto_fit_columns(writer.sheets["Transaksi"])
        auto_fit_columns(writer.sheets["Summary"])


# =========================
# MAIN
# =========================
def main():
    file_name = input("Enter the filename without extension: ").strip()

    pdf_dir = Path("pdf_file")
    excel_dir = Path("excel_file")
    excel_dir.mkdir(parents=True, exist_ok=True)

    pdf_file = pdf_dir / f"{file_name}.pdf"
    output_file = excel_dir / f"{file_name}.xlsx"

    if not pdf_file.exists():
        raise FileNotFoundError(f"PDF file not found: {pdf_file}")

    print("Membaca PDF BNI...")
    df_raw = read_pdf_table(pdf_file)

    print("Membersihkan header/footer...")
    df_raw = remove_garbage_rows(df_raw)

    print("Memisahkan summary...")
    df_trans, df_summary = split_summary_rows(df_raw)

    print("Menggabungkan multiline transaksi...")
    df_merged = merge_transactions(df_trans)

    print("Finalisasi DB/CR dan saldo...")
    df_final = finalize_transactions(df_merged)

    print("Menulis ke Excel...")
    export_to_excel(df_final, df_summary, output_file)

    print(f"Data successfully written to {output_file}")


if __name__ == "__main__":
    main()