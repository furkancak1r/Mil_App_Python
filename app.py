import tkinter as tk
import win32com.client as win32
import os

# Excel application'ı başlat
excel = win32.gencache.EnsureDispatch('Excel.Application')
excel.Visible = True  # Excel penceresini görünür yap

# A2'de "Notlar" yazısını ekleyen fonksiyon


def add_notes_title(worksheet):
    worksheet.Cells(2, 1).Value = "Notlar"  # Hücreye "Notlar" yazısı ekle
    worksheet.Cells(2, 1).Font.Bold = True  # Yazıyı kalın yap
    worksheet.Cells(2, 1).HorizontalAlignment = -4108  # Ortala hizala
    worksheet.Cells(2, 4).Value = "Ürün Adeti"  # Hücreye "Notlar" yazısı ekle
    worksheet.Cells(2, 4).Font.Bold = True  # Yazıyı kalın yap
    worksheet.Cells(2, 4).HorizontalAlignment = -4108  # Ortala hizala
    worksheet.Range("B2:C2").Merge()  # B2 ve C2 hücrelerini birleştir

    # "A1:E1" aralığındaki hücreleri birleştir ve fontunu kalın yap
    merged_cell = worksheet.Range("A1:E1")
    merged_cell.Merge()
    merged_cell.Font.Bold = True


# A3'den D'deki en son satıra kadar olan hücrelere kenarlık eklemek için fonksiyon


def add_border_to_range(worksheet, start_cell, end_cell):
    range_to_border = worksheet.Range(start_cell, end_cell)
    range_to_border.Borders.LineStyle = 1  # Kenarlık çizgilerini ince olarak ayarla


# A3'den E3'e kadar olan verileri bir diziye ekleyin
header = ["Malzeme Kodu", "Malzeme Açıklaması",
          "Birim Sarf Miktar", "Toplam Sarf Miktar", "Birim"]

# Excel dosyasını oluşturmak için fonksiyon


def create_excel():
    excel_filename = excel_filename_entry.get()
    copied_text = root.clipboard_get()  # Kopyalanan metni al
    current_directory = os.getcwd()  # Python dosyasının bulunduğu dizin
    excel_file_path = os.path.join(
        current_directory, excel_filename)  # Excel dosyasının tam yolu
    workbook = excel.Workbooks.Add()
    worksheet = workbook.Worksheets(1)
    worksheet.Range("A:E").VerticalAlignment = -4108  # Dikeyde ortala
    worksheet.Range("A:B").HorizontalAlignment = -4108  # A kolonunu ortala
    worksheet.Range("B:C").HorizontalAlignment = - \
        4131  # B kolonunu sola dayalı
    worksheet.Range("C:D").HorizontalAlignment = - \
        4108  # C kolonunu ortala
    worksheet.Range("D:E").HorizontalAlignment = - \
        4108  # A kolonunu ortala hizala
    worksheet.Range("A3:E3").HorizontalAlignment = -4108
    worksheet.Range("A3:E3").VerticalAlignment = -4108
    worksheet.Rows[2].RowHeight = 100  # 2. satırın yüksekliğini 100 yap

    # Başlık verilerini A3'den D3'e yerleştirin
    for col, header_text in enumerate(header, 1):
        cell = worksheet.Cells(3, col)
        cell.Value = header_text

    # Kopyalanan metni satır satır Excel'e ekleyin, A4 hücresinden başlayarak
    row = 4  # Başlangıç satırı
    for line in copied_text.split("\n"):  # Kopyalanan metni satıra göre ayır
        values = line.split("\t")  # Satırdaki değerleri tab ile ayır
        col = 1  # Başlangıç sütunu

        for value in values:  # Satırdaki her değer için
            if values != ['']:  # tab'dan kalan son boşluğu es geçmek için
                # Öncelikle hücre biçimini metin olarak ayarla çünkü diğer türlü uzun sayılarda virgül yok oluyor
                worksheet.Cells(row, 4).NumberFormat = "@"

                if col == 1:
                    worksheet.Cells(row, 1).Value = value  # Hücreye değeri yaz
                    if not value:
                        # Kırmızı rengi temsil eden değer
                        worksheet.Cells(row, 1).Interior.Color = 255
                elif col == 2:
                    worksheet.Cells(row, 2).Value = value  # Hücreye değeri yaz
                    if not value:
                        # Kırmızı rengi temsil eden değer
                        worksheet.Cells(row, 2).Interior.Color = 255
                elif col == 3:
                    worksheet.Cells(row, 4).Value = value  # Hücreye değeri yaz
                    worksheet.Cells(row, 4).NumberFormat = "0.00"

                    if not value:
                        # Kırmızı rengi temsil eden değer
                        worksheet.Cells(row, 4).Interior.Color = 255
                elif col == 4:
                    worksheet.Cells(row, 5).Value = value  # Hücreye değeri yaz
                    if not value:
                        # Kırmızı rengi temsil eden değer
                        worksheet.Cells(row, 5).Interior.Color = 255

                col += 1  # Sütunu bir artır
        # Bir sonraki satıra geçmeden önce kontrol et
        if values:
            row += 1  # Satırı bir artır

    # A3'den D'deki en son satıra kadar olan hücrelere kenarlık ekleyin
    add_border_to_range(worksheet, "A1", "E" + str(row - 2))
    worksheet.Columns.AutoFit()

    # A2'de "Notlar" yazısını ekleyin
    add_notes_title(worksheet)

    workbook.SaveAs(excel_file_path)  # Excel dosyasını belirtilen yere kaydet

    # Giriş alanını ve düğmeyi kaldır
    excel_filename_label.pack_forget()
    excel_filename_entry.pack_forget()
    create_button.pack_forget()

    # "Excel Oluşturuldu" yazısını göster
    excel_created_label = tk.Label(root, text="Excel dosyası oluşturuldu!")
    excel_created_label.pack()

    # Programı 2 saniye sonra kapat
    root.after(1500, lambda: root.destroy())


# Tkinter penceresini oluştur
root = tk.Tk()
root.geometry("400x200")
root.title("Excel Oluştur")

# "Evet" düğmesi
yes_button = tk.Button(root, text="Evet", command=lambda: [confirmation_label.pack_forget(
), yes_button.pack_forget(), excel_filename_label.pack(), excel_filename_entry.pack(), create_button.pack()])

# Kullanıcıya reçeteyi kopyaladığından emin mi diye sor
confirmation_label = tk.Label(
    root, text="Reçeteyi başlıksız olarak kopyaladığınızdan emin misiniz?")
confirmation_label.pack()

# Excel dosyasının adı için etiket
excel_filename_label = tk.Label(root, text="Excel Dosyası Adı:")

# Excel dosyası adı için giriş alanı
excel_filename_entry = tk.Entry(root)

# "Oluştur" düğmesi
create_button = tk.Button(root, text="Oluştur", command=create_excel)

# "Evet" düğmesini ekleyin
yes_button.pack()

# Tkinter penceresini başlat
root.mainloop()
