import openpyxl

def update_excel_file(file_path, sheet_name):
    # Excel dosyasını yükle
    workbook = openpyxl.load_workbook(file_path)
    
    # Belirtilen sayfayı seç
    if sheet_name in workbook.sheetnames:
        sheet = workbook[sheet_name]
    else:
        print(f"'{sheet_name}' adında bir sayfa bulunamadı.")
        return
    
    # G ve H sütunlarını kontrol et ve boş hücreleri 0 ile doldur
    for row in range(1, 2341):  # 1'den 307'ye kadar (307 dahil)
        for col in ['G', 'H']:
            cell = sheet[f"{col}{row}"]
            if cell.value is None:
                cell.value = 0

    # Dosyayı kaydet
    workbook.save(file_path)
    print(f"{file_path} dosyası başarıyla güncellendi.")

# Kullanım örneği
file_path = "C:\\Users\\furkan.cakir\\Desktop\\FurkanPRS\\Kodlar\\Planlama\\Exceller\\KILAVUZ VE STOKLAR\\OITM - Items1.xlsx" # Excel dosyasının yolu
sheet_name = 'proso'            # Excel sayfa adı
update_excel_file(file_path, sheet_name)
