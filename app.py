import tkinter as tk
import win32com.client as win32
import os

# Excel application'ı başlat
excel = win32.gencache.EnsureDispatch('Excel.Application')
excel.Visible = True  # Excel penceresini görünür yap

# A3'den E3'e kadar olan verileri bir diziye ekleyin
header = ["Malzeme Kodu", "Malzeme Açıklaması",
          "Birim Sarf Miktarı", "Toplam Sarf Miktarı", "Birim"]

# A2'de "Notlar" yazısını ekleyen fonksiyon


def add_notes_title(worksheet):
    worksheet.Cells(2, 1).Value = "Notlar"  # Hücreye "Notlar" yazısı ekle
    worksheet.Cells(2, 1).Font.Bold = True  # Yazıyı kalın yap
    worksheet.Cells(2, 1).HorizontalAlignment = -4108  # Ortala hizala
    worksheet.Cells(2, 4).Value = "Ürün Adeti"  # Hücreye "Notlar" yazısı ekle
    worksheet.Cells(2, 4).Font.Bold = True  # Yazıyı kalın yap
    worksheet.Cells(2, 4).HorizontalAlignment = -4108  # Ortala hizala
    worksheet.Range("B2:C2").Merge()  # B2 ve C2 hücrelerini birleştir

# A3'den D'deki en son satıra kadar olan hücrelere kenarlık eklemek için fonksiyon


def add_border_to_range(worksheet, start_cell, end_cell):
    range_to_border = worksheet.Range(start_cell, end_cell)
    range_to_border.Borders.LineStyle = 1  # Kenarlık çizgilerini ince olarak ayarla


def create_excel():
    try:
        copied_text = root.clipboard_get()  # Kopyalanan metni al
    except tk.TclError:
        copied_text = ""

    if not copied_text:
        warning_label.config(text="Lütfen ürünleri kopyalayın!")
    else:
        warning_label.destroy()  # Label'ı kaldır
        approval_label.config(text="Excel oluşturuluyor...")
        approval_label.place(relx=0.5, rely=0.72, anchor="center")
        root.update()  # Arayüzü güncelle
        create_excelfn(copied_text)  # Excel oluştur
        approval_label.config(text="Excel oluşturuldu!")  # Sonucu göster
        approval_label.place(relx=0.5, rely=0.4, anchor="center")


# Excel dosyasını oluşturmak için fonksiyon
def create_excelfn(copied_text):
    excel_filename = excel_filename_entry.get()
    excel_product_count = excel_product_count_entry.get()

    current_directory = os.getcwd()  # Python dosyasının bulunduğu dizin
    excel_file_path = os.path.join(
        current_directory, excel_filename)  # Excel dosyasının tam yolu

    # Excel dosyasını oluştur
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

    # Ürün adetini E2 hücresine yaz ve fontu kalın yap
    worksheet.Cells(2, 5).Value = excel_product_count
    worksheet.Cells(2, 5).Font.Bold = True

    # Başlık verilerini A3'den D3'e yerleştirin
    for col, header_text in enumerate(header, 1):
        cell = worksheet.Cells(3, col)
        cell.Value = header_text
        cell.Font.Bold = True

    # A ve E sütunlarını birleştir ve Excel dosya adını içeren hücreyi oluştur
    worksheet.Range("A1:E1").Merge()
    worksheet.Range("A1").Value = excel_filename
    worksheet.Range("A1").Font.Bold = True
    worksheet.Range("A1").HorizontalAlignment = -4108  # Ortala hizala

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
                    if value:
                        # Eğer 'value' boş değilse işlem yap
                        # Virgülü nokta ile değiştirip ondalık sayıya çevir
                        value_float = float(value.replace(",", "."))
                        worksheet.Cells(row, 4).Value = value_float
                        try:
                            # Virgülü nokta ile değiştirip ondalık sayıya çevir
                            excel_product_count_float = float(
                                excel_product_count.replace(",", "."))
                            worksheet.Cells(
                                row, 3).Value = value_float / excel_product_count_float
                        except ValueError:
                            pass
                    else:
                        # 'value' boşsa hata işleme veya başka bir işlem yapabilirsiniz
                        worksheet.Cells(row, 3).Interior.Color = 255
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
    excel_filename_label.place_forget()
    excel_filename_entry.place_forget()
    excel_product_count_label.place_forget()
    excel_product_count_entry.place_forget()
    create_button.place_forget()

    # Programı 2 saniye sonra kapat
    root.after(1500, lambda: root.destroy())


def create_root():
    root = tk.Tk()
    window_width = 400
    window_height = 200
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()

    # Pencereyi ekranın ortasına konumlandır
    x = (screen_width - window_width) // 2
    y = (screen_height - window_height-150) // 2
    root.geometry(f"{window_width}x{window_height}+{x}+{y}")
    root.minsize(window_width, window_height)  # Minimum boyutu ayarla

    root.title("Excel Oluştur")
    return root


def create_warning_label(root):
    warning_label = tk.Label(root, text="", fg="red")
    warning_label.place(relx=0.5, rely=0.72, anchor="center")
    return warning_label


def create_approval_label(root):
    approval_label = tk.Label(root, text="", fg="green")
    return approval_label


def create_excel_filename_label(root):
    excel_filename_label = tk.Label(root, text="Excel Dosya Adı:")
    return excel_filename_label


def create_excel_product_count_label(root):
    excel_product_count_label = tk.Label(root, text="Ürün Adeti:")
    return excel_product_count_label


def create_excel_filename_entry(root):
    excel_filename_entry = tk.Entry(root)
    return excel_filename_entry


def create_excel_product_count_entry(root):
    def validate_input(P):
        # Kullanıcının girdiği değeri değerlendir
        if P == "" or P.isdigit():
            return True
        else:
            return False

    vcmd = root.register(validate_input)
    excel_product_count_entry = tk.Entry(
        root, validate="key", validatecommand=(vcmd, "%P"))
    return excel_product_count_entry


def create_create_button(root, create_excel):
    create_button = tk.Button(root, text="Oluştur", command=create_excel)
    # Başlangıçta düğmeyi devre dışı bırak
    create_button.config(state="disabled")

    def check_and_enable_button(event):
        excel_filename = excel_filename_entry.get()
        excel_product_count = excel_product_count_entry.get()
        if not excel_filename or not excel_product_count:
            create_button.config(state="disabled")
        else:
            create_button.config(state="normal")

    excel_filename_entry.bind(
        "<KeyRelease>", check_and_enable_button)
    excel_product_count_entry.bind(
        "<KeyRelease>", check_and_enable_button)

    return create_button


# Tkinter penceresini oluştur
root = create_root()


# Uyarı etiketleri
warning_label = create_warning_label(root)
approval_label = create_approval_label(root)


# Excel dosyasının adı için etiket
excel_filename_label = create_excel_filename_label(root)

# Excel Ürün adeti labelı
excel_product_count_label = create_excel_product_count_label(root)

# Excel dosyası adı için
excel_filename_entry = create_excel_filename_entry(root)

# Excel Ürün adeti için giriş alanı
excel_product_count_entry = create_excel_product_count_entry(root)

# "Oluştur" düğmesi
create_button = create_create_button(root, create_excel)

excel_filename_label.place(relx=0.5, rely=0.2, anchor="center")
excel_filename_entry.place(relx=0.5, rely=0.3, anchor="center")
excel_product_count_label.place(relx=0.5, rely=0.5, anchor="center")
excel_product_count_entry.place(relx=0.5, rely=0.6, anchor="center")
create_button.place(relx=0.5, rely=0.85, anchor="center")

# Tkinter penceresini başlat
root.mainloop()
