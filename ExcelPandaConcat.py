import pandas as pd

# Daftar file Excel yang akan digabungkan
files = ['L.K_JULI_2013_Klaten.xlsx', 'L.K_SEPT_2013_Klaten.xlsx', 'L.K_OKT_2013_Klaten.xlsx','L.K_NOV_2013_Klaten.xlsx','L.K_DES_2013_Klaten.xlsx']

# Nama sheet yang akan digabungkan
sheet_name = 'Kas'

# Membuat DataFrame kosong untuk menampung data gabungan
combined_data = pd.DataFrame()

# Melakukan iterasi pada setiap file
for file in files:
    # Membaca file Excel dan mengambil sheet yang diinginkan
    data = pd.read_excel(file, sheet_name=sheet_name)
    
    # Menggabungkan data pada setiap iterasi
    combined_data = pd.concat([combined_data, data], ignore_index=True)

# Menyimpan data gabungan menjadi file Excel
combined_data.to_excel('concatdata.xlsx', index=False)

print("Data berhasil digabungkan dan disimpan dalam file 'concatdata.xlsx'")
