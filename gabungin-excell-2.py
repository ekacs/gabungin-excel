import pandas as pd
import glob
import os

# Tentukan folder yang berisi file Excel
folder_path = r'd:/target-folder/*.xlsx'

# Buat writer untuk menulis file Excel baru
output_file = 'file_gabungan_target.xlsx'

try:
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        sheet_written = False  # Flag untuk mengecek apakah ada sheet yang ditulis
        # Loop untuk setiap file Excel
        for file in glob.glob(folder_path):
            try:
                # Ambil nama file tanpa path
                file_name = os.path.basename(file).split('.')[0]

                # Loop untuk setiap sheet di file Excel
                xls = pd.ExcelFile(file)
                for sheet_name in xls.sheet_names:
                    # Baca sheet dan tulis ke file Excel baru
                    df = pd.read_excel(file, sheet_name=sheet_name)
                    
                    if not df.empty:  # Hanya tulis sheet jika tidak kosong
                        # Sanitasi nama sheet (hilangkan karakter tidak valid)
                        sanitized_sheet_name = f"{sheet_name}_{file_name}".replace('/', '_').replace('\\', '_').replace('*', '_').replace('[', '_').replace(']', '_').replace(':', '_').replace('?', '_')
                        
                        # Batasi panjang nama sheet hingga maksimal 31 karakter
                        sanitized_sheet_name = sanitized_sheet_name[:31]
                        
                        # Simpan ke sheet baru
                        df.to_excel(writer, sheet_name=sanitized_sheet_name, index=False)
                        sheet_written = True
                        
            except Exception as e:
                print(f"Terjadi kesalahan saat memproses file {file}: {e}")

        if not sheet_written:
            # Tambahkan sheet dummy jika tidak ada sheet yang ditulis
            pd.DataFrame().to_excel(writer, sheet_name='Sheet1', index=False)

    print(f"File Excel berhasil digabungkan dan disimpan ke {output_file}")

except Exception as e:
    print(f"Terjadi kesalahan saat menyimpan file gabungan: {e}")