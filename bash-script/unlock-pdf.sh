#!/bin/bash

# Cek apakah qpdf sudah terinstall
if command -v qpdf >/dev/null 2>&1; then
  echo "qpdf sudah terinstall ✔"
else
  echo "qpdf belum terinstall, menginstall sekarang..."

  # Deteksi package manager
  if command -v apt >/dev/null 2>&1; then
    sudo apt update && sudo apt install -y qpdf
  elif command -v yum >/dev/null 2>&1; then
    sudo yum install -y qpdf
  elif command -v dnf >/dev/null 2>&1; then
    sudo dnf install -y qpdf
  elif command -v pacman >/dev/null 2>&1; then
    sudo pacman -Sy qpdf --noconfirm
  else
    echo "Package manager tidak dikenali. Install qpdf manual."
    exit 1
  fi
fi

# Input user
read -p "Input file PDF: " input_pdf
read -p "Password: " password
echo
read -p "Output file: " output_pdf

# Validasi file
input_dir="/mnt/c/Users/harin/pdf_to_excel/bash-script/input"
output_dir="/mnt/c/Users/harin/pdf_to_excel/bash-script/output"


input_path="$input_dir/$input_pdf"
output_path="$output_dir/$output_pdf"

if [ ! -f "$input_path" ]; then
  echo "File tidak ditemukan di $input_path"
  exit 1
fi

qpdf --password="$password" --decrypt "$input_path" "$output_path"

echo "Selesai → $output_path"
