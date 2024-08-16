import pandas as pd
import glob
import os

# Tentukan folder yang berisi file Excel (gunakan raw string untuk path)
folder_path = r'd:/Downloads/#TEMPORARY FILES/BAPK BP3IP/blanko cop januari 2023 - agustus 2024/all/*.xlsx'

# Buat writer untuk menulis file Excel baru
output_file = 'file_gabungan_all.xlsx'
with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    # Loop untuk setiap file Excel
    for file in glob.glob(folder_path):
        # Ambil nama file tanpa path
        file_name = os.path.basename(file).split('.')[0]
        
        # Loop untuk setiap sheet di file Excel
        xls = pd.ExcelFile(file)
        for sheet_name in xls.sheet_names:
            # Baca sheet dan tulis ke file Excel baru
            df = pd.read_excel(file, sheet_name=sheet_name)
            
            # Sanitasi nama sheet (hilangkan karakter tidak valid)
            sanitized_sheet_name = f"{sheet_name}_{file_name}".replace('/', '_').replace('\\', '_').replace('*', '_').replace('[', '_').replace(']', '_').replace(':', '_').replace('?', '_')
            
            # Batasi panjang nama sheet hingga maksimal 31 karakter
            sanitized_sheet_name = sanitized_sheet_name[:31]
            
            # Simpan ke sheet baru
            df.to_excel(writer, sheet_name=sanitized_sheet_name, index=False)

print(f"File Excel berhasil digabungkan dan disimpan ke {output_file}")