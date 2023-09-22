import tkinter as tk
import win32com.client as win32
import os
import json
from tkinter import ttk
from ttkthemes import  ThemedTk
# Excel application'ı başlat
excel = win32.gencache.EnsureDispatch('Excel.Application')
excel.Visible = True  # Excel penceresini görünür yap

# Sabitler
EXCEL_BORDER_STYLE = 1
EXCEL_TEXT_ALIGNMENT_CENTER = -4108
EXCEL_HORIZONTAL_ALIGNMENT_LEFT = -4131
EXCEL_VERTICAL_ALIGNMENT_CENTER = -4108
EXCEL_HORIZONTAL_ALIGNMENT_CENTER = -4108

# Excel başlık verileri
header = ["Malzeme Kodu", "Malzeme Açıklaması",
          "Birim Sarf Miktarı", "Toplam Sarf Miktarı", "Birim"]

y = 0.2
# A2'de "Notlar" yazısını ekleyen fonksiyon

def add_notes_title(worksheet):
    worksheet_range = worksheet.Range
    worksheet_cells = worksheet.Cells
    
    worksheet_cells(2, 1).Value = "Notlar"
    worksheet_cells(2, 1).Font.Bold = True
    worksheet_cells(2, 1).HorizontalAlignment = EXCEL_TEXT_ALIGNMENT_CENTER
    
    worksheet_cells(2, 4).Value = "Ürün Adeti"
    worksheet_cells(2, 4).Font.Bold = True
    worksheet_cells(2, 4).HorizontalAlignment = EXCEL_TEXT_ALIGNMENT_CENTER
    
    worksheet_range("B2:C2").Merge()

# A3'den D'deki en son satıra kadar olan hücrelere kenarlık eklemek için fonksiyon
def add_border_to_range(worksheet, start_cell, end_cell):
    range_to_border = worksheet.Range(start_cell, end_cell)
    borders = range_to_border.Borders
    borders.LineStyle = EXCEL_BORDER_STYLE


def load_words_to_remove(file_path):
 
    with open(file_path, 'r') as file:
        data = json.load(file)
        return data["words_to_remove"]
    
def remove_selected_words(data):
   
    # Belirtilen kelimeleri büyük harfe çevir
    words_to_remove = load_words_to_remove('milJsonFiles/sacSil.json')
    words_to_remove = [word.upper() for word in words_to_remove]
    
    # Veriyi satırlara böler
    lines = data.split("\n")
    
    # Temizlenmiş veriyi saklamak için bir liste oluştur
    cleaned_data = []
    
    # Satırları dolaş
    for line in lines:
        # Varsa kelimeleri kaldır
        if not any(word in line.upper() for word in words_to_remove):
            cleaned_data.append(line)
    
    # Temizlenmiş veriyi birleştir ve döndür
    return "\n".join(cleaned_data)


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
        approval_label.place(relx=0.5, rely=y-0.1, anchor="center")
        root.update()  # Arayüzü güncelle
        if sac_sil_flag.get(): # Eğer sac sil seçiliyse
            cleaned_text = remove_selected_words(copied_text) # Kopyalanan metinden belirtilen kelimeleri sil
            create_excelfn(cleaned_text) # Temizlenmiş veri ile Excel oluştur
        else: # Eğer sac sil seçili değilse
            create_excelfn(copied_text) # Kopyalanan metni olduğu gibi Excel'e yaz
        approval_label.config(text="Excel oluşturuldu!")  # Sonucu göster
        approval_label.place(relx=0.5, rely=y+0.2, anchor="center")

# Excel dosyasını oluşturmak için fonksiyon
def create_excelfn(copied_text):
    product_name = product_name_entry.get()
    order_name = order_number_entry.get()

    excel_product_count = excel_product_count_entry.get()

    current_directory = os.getcwd()  # Python dosyasının bulunduğu dizin
    os.mkdir(os.path.join(current_directory, order_name)) # Klasörü oluşturur

    excel_file_path = os.path.join(current_directory, order_name, order_name+" "+product_name) # Excel dosyasının tam yolu

    # Excel dosyasını oluştur
    workbook = excel.Workbooks.Add()
    worksheet = workbook.Worksheets(1)
    worksheet.Range("A:E").VerticalAlignment = EXCEL_VERTICAL_ALIGNMENT_CENTER
    worksheet.Range("A:B").HorizontalAlignment = EXCEL_HORIZONTAL_ALIGNMENT_CENTER
    worksheet.Range("B:C").HorizontalAlignment = EXCEL_HORIZONTAL_ALIGNMENT_LEFT
    worksheet.Range("C:D").HorizontalAlignment = EXCEL_HORIZONTAL_ALIGNMENT_CENTER
    worksheet.Range("D:E").HorizontalAlignment = EXCEL_HORIZONTAL_ALIGNMENT_CENTER
    worksheet.Range("A3:E3").HorizontalAlignment = EXCEL_TEXT_ALIGNMENT_CENTER
    worksheet.Range("A3:E3").VerticalAlignment = EXCEL_VERTICAL_ALIGNMENT_CENTER
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
    worksheet.Range("A1").Value = order_name+" "+product_name 
    worksheet.Range("A1").Font.Bold = True
    worksheet.Range("A1").HorizontalAlignment = EXCEL_TEXT_ALIGNMENT_CENTER

    # Metni bir defada parçalayarak işleme
    row = 4  # Başlangıç satırı
    lines = copied_text.split("\n")
    for line in lines:
        values = line.split("\t") # Satırdaki değerleri tab ile ayır
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
    product_name_label.place_forget()
    product_name_entry.place_forget()
    order_number_label.place_forget()
    order_number_entry.place_forget()
    excel_product_count_label.place_forget()
    excel_product_count_entry.place_forget()
    remove_sheet_metal_checkbox.place_forget()
    create_button.place_forget()

    # Programı 2 saniye sonra kapat
    root.after(1500, lambda: root.destroy())

def create_root():
    root = ThemedTk(theme='adapta', themebg=True)
    window_width = 600
    window_height = 400
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    print(root.get_themes())
    

  
    

    # Pencereyi ekranın ortasına konumlandır
    x = (screen_width - window_width) // 2
    y = (screen_height - window_height-150) // 2
    root.geometry(f"{window_width}x{window_height}+{x}+{y}")
    root.minsize(window_width, window_height)  # Minimum boyutu ayarla
    # Ayarlar düğmesini oluştur
    settings_button = ttk.Button(root, text="⚙️", command="toggle_settings")

    # Ayarlar düğmesini ana pencereye yerleştir
    settings_button.grid(row=0, column=1, padx=500, pady=10, sticky="e")

    root.title("Mil Excel & Pdf Oluşturma")
    return root

def create_warning_label(root):
    # Kırmızı renkli bir stil oluştur
    style = ttk.Style()
    style.configure("Red.TLabel", foreground="red")
    
    # Stili uygulanan bir Label widget'ı oluştur
    warning_label = ttk.Label(root, text="", style="Red.TLabel")
    warning_label.place(relx=0.5, rely=y-0.1, anchor="center")
    return warning_label

def create_approval_label(root):
    # Yeşil renkli bir stil oluştur
    style = ttk.Style()
    style.configure("Green.TLabel", foreground="green")
    
    # Stili uygulanan bir Label widget'ı oluştur
    approval_label = ttk.Label(root, text="", style="Green.TLabel")
    return approval_label

def set_font_style():
    style = ttk.Style()
    style.configure("Custom.TLabel",font=(12))
    return style
def create_product_name_label(root):
    style=set_font_style()
    # Varsayılan stil ile bir Label widget'ı oluştur
    product_name_label = ttk.Label(root, text="Ürün Adı:",style="Custom.TLabel")

    return product_name_label

def create_order_number_label(root):
    # Varsayılan stil ile bir Label widget'ı oluştur
    style=set_font_style()

    order_number_label = ttk.Label(root, text="Sipariş Numarası:",style="Custom.TLabel")
    return order_number_label

def create_excel_product_count_label(root):
    style=set_font_style()

    # Varsayılan stil ile bir Label widget'ı oluştur
    excel_product_count_label = ttk.Label(root, text="Ürün Adeti:",style="Custom.TLabel")
    return excel_product_count_label
def create_product_name_entry(root):
    product_name_entry = ttk.Entry(root)
    return product_name_entry

def create_order_number_entry(root):
    order_number_entry = ttk.Entry(root)
    return order_number_entry
def create_remove_sheet_metal_checkbox_entry(root):
    # "Sac Sil" butonuna tıklanıp tıklanmadığını takip eden değişken
    sac_sil_flag = tk.BooleanVar()
    sac_sil_flag.set(False)  # Başlangıçta "Sac Sil" butonu işaretsiz

    style = ttk.Style()
    style.configure("Custom.TCheckbutton", font=("Segoe UI", 12))  # Segoe UI fontu ve 12 punto olarak ayarla

    remove_sheet_metal_checkbox = ttk.Checkbutton(
        root, text="Sac Sil", variable=sac_sil_flag, style="Custom.TCheckbutton")

    return remove_sheet_metal_checkbox, sac_sil_flag

def create_excel_product_count_entry(root):
    def validate_input(P):
        # Kullanıcının girdiği değeri değerlendir
        if P == "" or P.isdigit():
            return True
        else:
            return False

    vcmd = root.register(validate_input)
    excel_product_count_entry = ttk.Entry(
        root, validate="key", validatecommand=(vcmd, "%P"))
    return excel_product_count_entry

def create_create_button(root, create_excel):
    
    style = ttk.Style()
    style.configure("Custom.TButton", font=("Segoe UI", 12))  # Segoe UI fontu ve 12 punto olarak ayarla
    create_button = ttk.Button(root, text="Oluştur", command=create_excel, style="Custom.TButton")
    # Başlangıçta düğmeyi devre dışı bırak
    create_button.config(state="disabled")

    def check_and_enable_button(event):
        product_name = product_name_entry.get()
        order_number = order_number_entry.get()
        excel_product_count = excel_product_count_entry.get()
        if not order_number or not product_name or not excel_product_count:
            create_button.config(state="disabled")
        else:
            create_button.config(state="normal")
    order_number_entry.bind(
        "<KeyRelease>", check_and_enable_button)
    product_name_entry.bind(
        "<KeyRelease>", check_and_enable_button)
    excel_product_count_entry.bind(
        "<KeyRelease>", check_and_enable_button)

    return create_button

# Tkinter penceresini oluştur
root = create_root()

# Uyarı etiketleri
warning_label = create_warning_label(root)
approval_label = create_approval_label(root)

# Sipariş numarası için etiket
order_number_label = create_order_number_label(root)

# Ürün adı için etiket
product_name_label = create_product_name_label(root)

# Excel Ürün adeti labelı
excel_product_count_label = create_excel_product_count_label(root)

# Sipariş numarası için giriş alanı
order_number_entry = create_order_number_entry(root)
order_number_entry.focus_set()  # order_number_entry'yi aktif hale getir

# Excel dosyası adı için
product_name_entry = create_product_name_entry(root)

# Excel Ürün adeti için giriş alanı
excel_product_count_entry = create_excel_product_count_entry(root)

# "Sac Sil" butonunu ve durumunu al
remove_sheet_metal_checkbox, sac_sil_flag = create_remove_sheet_metal_checkbox_entry(root)

# "Oluştur" düğmesi
create_button = create_create_button(root, create_excel)



order_number_label.place(relx=0.35, rely=y+0.1, anchor="center")
order_number_entry.place(relx=0.6, rely=y+0.1, anchor="center")
product_name_label.place(relx=0.4, rely=y+0.2, anchor="center")
product_name_entry.place(relx=0.6, rely=y+0.2, anchor="center")
excel_product_count_label.place(relx=0.39, rely=y+0.3, anchor="center")
excel_product_count_entry.place(relx=0.6, rely=y+0.3, anchor="center")
remove_sheet_metal_checkbox.place(relx=0.5, rely=y+0.425, anchor="center")

create_button.place(relx=0.5, rely=y+0.55, anchor="center")

# Tkinter penceresini başlat
root.mainloop()
