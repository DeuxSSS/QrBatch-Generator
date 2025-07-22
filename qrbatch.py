import qrcode
from PIL import Image
import pandas as pd
import os
import sys
from tkinter import Tk, filedialog  
from colorama import Fore, init

def buat_folder_output(folder_path):
    """Membuat folder output jika belum ada"""
    try:
        os.makedirs(folder_path, exist_ok=True)
        print(f"Folder output dibuat di: {os.path.abspath(folder_path)}")
    except Exception as e:
        print(f"Gagal membuat folder: {str(e)}")
        sys.exit(1)

def proses_logo(logo_path, size_logo=100):
    """Memproses logo untuk mempertahankan transparansi"""
    try:
        logo = Image.open(logo_path).convert("RGBA")
        logo = logo.resize((size_logo, size_logo), Image.LANCZOS)
        
        # Proses transparansi
        logo_data = logo.getdata()
        new_data = []
        for item in logo_data:
            if item[3] < 10:
                new_data.append((255, 255, 255, 0))
            else:
                new_data.append(item)
        
        logo.putdata(new_data)
        return logo
    except Exception as e:
        print(f"Error memproses logo: {str(e)}")
        return None

def buat_qr_dengan_logo(data, output_path, logo=None, warna_qr="#000000", warna_bg='white'):
    """Membuat QR code dengan logo di tengah"""
    try:
        qr = qrcode.QRCode(
            version=2,
            error_correction=qrcode.constants.ERROR_CORRECT_H,
            box_size=10,
            border=4,
        )
        qr.add_data(str(data))
        qr.make(fit=True)
        
        img_qr = qr.make_image(fill_color=warna_qr, back_color=warna_bg).convert('RGB')
        
        if logo:
            pos = ((img_qr.size[0] - logo.size[0]) // 2, 
                   (img_qr.size[1] - logo.size[1]) // 2)
            img_qr.paste(logo, pos, logo)
        
        img_qr.save(output_path, quality=95)
        print(f"Berhasil membuat: {os.path.basename(output_path)}")
        return True
    except Exception as e:
        print(f"Gagal membuat QR untuk {data}: {str(e)}")
        return False

def baca_data_siswa(file_path, kolom_nisn, kolom_nama):
    """Membaca data siswa dari file Excel/CSV"""
    try:
        if file_path.endswith('.xlsx'):
            df = pd.read_excel(file_path, engine='openpyxl')
        elif file_path.endswith('.csv'):
            df = pd.read_csv(file_path)
        else:
            print("Format file tidak didukung. Gunakan .xlsx atau .csv")
            return None
        
        if kolom_nisn not in df.columns:
            print(f"Kolom '{kolom_nisn}' tidak ditemukan dalam file")
            return None
            
        if kolom_nama not in df.columns:
            print(f"Kolom '{kolom_nama}' tidak ditemukan dalam file")
            return None
            
        return df[[kolom_nisn, kolom_nama]].dropna()
    except Exception as e:
        print(f"Error membaca file data: {str(e)}")
        return None

def pilih_file(judul, tipe_file):
    """Membuka file explorer untuk memilih file"""
    root = Tk()
    root.withdraw()  # Sembunyikan jendela tkinter utama
    root.attributes('-topmost', True)  # Jadikan di atas jendela lain
    
    if tipe_file == 'excel':
        filetypes = [("Excel files", "*.xlsx"), ("CSV files", "*.csv"), ("All files", "*.*")]
    elif tipe_file == 'logo':
        filetypes = [("Image files", "*.png *.jpg *.jpeg *.bmp"), ("All files", "*.*")]
    else:
        filetypes = [("All files", "*.*")]
    
    file_path = filedialog.askopenfilename(
        title=judul,
        filetypes=filetypes
    )
    return file_path if file_path else None

def pilih_folder(judul):
    """Membuka file explorer untuk memilih folder"""
    root = Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    folder_path = filedialog.askdirectory(title=judul)
    return folder_path if folder_path else None

def main():
    init()  
    print(Fore.GREEN + """
========================================================================
  ___  ____       ____ _____ _   _ _____ ____      _  _____ ___  ____  
 / _ \|  _ \     / ___| ____| \ | | ____|  _ \    / \|_   _/ _ \|  _ \ 
| | | | |_) |   | |  _|  _| |  \| |  _| | |_) |  / _ \ | || | | | |_) |
| |_| |  _ <    | |_| | |___| |\  | |___|  _ <  / ___ \| || |_| |  _ < 
 \__\_\_| \_\    \____|_____|_| \_|_____|_| \_\/_/   \_\_| \___/|_| \_\  
========================================================================  
CSV/XLSX ONLY                                           Made by: DeuxSSS""")
    print(Fore.RED + "\n=== PROGRAM PEMBUAT QR CODE NISN DENGAN LOGO ===")
    print("Versi 2.1 - Format Nama File NISN_Nama\n")
    
    # Pilih file data siswa via GUI
    print(Fore.RESET + "Silakan pilih file data siswa (Excel/CSV)...")
    file_data = pilih_file("Pilih File Data Siswa (Excel/CSV)", 'excel')
    if not file_data:
        print("\nERROR: File data tidak dipilih!")
        return
    
    # Input nama kolom (tetap manual karena bergantung pada struktur file)
    kolom_nisn = input("\nMasukkan nama kolom yang berisi NISN: ").strip()
    kolom_nama = input("Masukkan nama kolom yang berisi Nama Siswa: ").strip()
    
    # Pilih folder output via GUI
    print("\nSilakan pilih folder untuk menyimpan QR Code...")
    folder_output = pilih_folder("Pilih Folder Output")
    if not folder_output:
        print("\nERROR: Folder output tidak dipilih!")
        return
    
    # Pilih logo via GUI (opsional)
    print("\nSilakan pilih file logo (kosongkan jika tanpa logo)...")
    logo_path = pilih_file("Pilih File Logo (PNG/JPG)", 'logo')
    
    # Proses logo jika ada
    logo = None
    if logo_path:
        logo = proses_logo(logo_path)
        if not logo:
            return
    
    # Buat folder output
    buat_folder_output(folder_output)
    
    # Baca data siswa
    print("\nMemproses data siswa...")
    data_siswa = baca_data_siswa(file_data, kolom_nisn, kolom_nama)
    if data_siswa is None:
        return
    
    # Buat QR code
    print(f"\nMembuat {len(data_siswa)} QR code...")
    for _, row in data_siswa.iterrows():
        nisn = row[kolom_nisn]
        nama = row[kolom_nama]
        nama_file = "".join(x for x in nama if x.isalnum() or x in (' ', '_')).strip()
        output_path = os.path.join(folder_output, f"{nama_file}_{nisn}.png")
        buat_qr_dengan_logo(
            data=nisn,
            output_path=output_path,
            logo=logo,
            warna_qr="#000000",
            warna_bg='white'
        )
    
    print("\nProses selesai!")
    print(f"QR code tersimpan di: {os.path.abspath(folder_output)}")

if __name__ == "__main__":
    main()
