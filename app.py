import tkinter as tk
import win32com.client as win32
import os
import json
from tkinter import ttk
from ttkthemes import ThemedTk
import time


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


def fetch_json_data(json_file):
    try:
        with open(json_file, 'r', encoding='utf-8') as file:
            data = json.load(file)
        return data
    except Exception as e:
        print(f"Hata oluştu: {str(e)}")
        return None

def remove_selected_words(data):

    # Belirtilen kelimeleri büyük harfe çevir
    response = fetch_json_data('milJsonFiles/sacSil.json')
    words_to_remove = response["words_to_remove"]
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


def validate_copied_text(copied_text):
    # Metni küçük harfe çevirip sekme sayısını say
    tab_count = copied_text.lower().count("\t")
    lowercase_text = copied_text.lower()  # Metni küçük harfe çevir
    # Belirli kelimeleri metinde küçük harfle ara
    word_here = any(word in lowercase_text for word in [
                    "adet", "kg", "pk", "mt", "metre", "takım"])
    # Sekme sayısı 2'den fazla ve belirli kelime varsa True, aksi takdirde False döndür
    return tab_count > 2 and word_here

# Verileri alıp renklerine göre sıralayıp sıralanmış verileri dönen fonksiyon
# Verileri alıp istediğiniz sıralamaya göre sıralayıp sıralanmış verileri dönen fonksiyon


def sort_data_by_color(data):

    # İstenen sıralama
    order = ["8696052", "11992832", "65535", "13408767",
             "14395790", "9359529", "10092441", ""]

    # Verileri satırlara böl (boş satırları atla)
    lines = data.splitlines()

    # Her satır için son sütunun değerini bul
    values = [line.split("\t")[-1] for line in lines]

    # Değerleri istenen sıralamaya göre indeksle
    indices = [order.index(value) for value in values]

    # Satırları indekslere göre sırala
    sorted_lines = [line for _, line in sorted(zip(indices, lines))]

    # Sıralanmış verileri birleştir
    sorted_data = "\n".join(sorted_lines)
    # Sıralanmış verileri döndür
    return sorted_data


def apply_colors(text):
    response = fetch_json_data('milJsonFiles/renkler.json')
    colors = response["colors"]
    result = []

    lines = text.split("\n")

    for line in lines:
        values = line.split("\t")
        formatted_line = []

        if len(values) >= 4:  # En az 4 değeri olan satırları işle
            # 1. değeri al (arasın kelime) ve küçük harfe çevir
            keyword_to_search = values[1].lower()
            rgb_color = ""  # Varsayılan olarak boş renk

            for color, keywords in colors.items():
                for keyword in keywords:
                    # Anahtar kelimeleri - ile böl
                    parts = keyword.lower().split("-")
                    # Bölünen parçaların hepsinin values[1]'de olup olmadığını kontrol et
                    if all(part in keyword_to_search for part in parts):
                        rgb_color = color
                        break

            # 4. değeri eklemek
            values.append(rgb_color)

        formatted_line = values
        result.append("\t".join(formatted_line))
    result_excel_format = "\n".join(result)
    sorted_data_by_color = sort_data_by_color(result_excel_format)
    return sorted_data_by_color


def create_excel():
    try:
        copied_text = root.clipboard_get()  # Kopyalanan metni al
    except tk.TclError:
        copied_text = ""

    if not copied_text:
        warning_label.config(text="Lütfen ürünleri kopyalayın!")
        warning_label.place(relx=0.5, rely=y-0.1, anchor="center")

    if not validate_copied_text(copied_text):
        warning_label.config(text="Yanlış içerik kopyalanmış!")
        warning_label.place(relx=0.5, rely=y-0.1, anchor="center")

        return

    else:
        warning_label.destroy()  # Label'ı kaldır
        approval_label.config(text="Excel oluşturuluyor...")
        create_button.config(state="disabled")
        approval_label.place(relx=0.5, rely=y-0.1, anchor="center")
        root.update()  # Arayüzü güncelle
        if sac_sil_flag.get():  # Eğer sac sil seçiliyse
            # Kopyalanan metinden belirtilen kelimeleri sil
            cleaned_text = remove_selected_words(copied_text)
            result = apply_colors(cleaned_text)

            create_excelfn(result)  # Temizlenmiş veri ile Excel oluştur
        else:  # Eğer sac sil seçili değilse
            # Kopyalanan metni olduğu gibi Excel'e yaz
            # Veriyi temizle
            result = apply_colors(copied_text)
            # Yeni veriyi kopyala
            create_excelfn(result)
        approval_label.config(text="Excel oluşturuldu!")  # Sonucu göster
        approval_label.place(relx=0.5, rely=y+0.2, anchor="center")

# Excel dosyasını oluşturmak için fonksiyon


def create_excelfn(copied_text):
    # Excel application'ı başlat
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    excel.Visible = True  # Excel penceresini görünür yap

    product_name = product_name_entry.get()
    order_name = order_number_entry.get()

    excel_product_count = excel_product_count_entry.get()

    current_directory = os.getcwd()  # Python dosyasının bulunduğu dizin
    os.mkdir(os.path.join(current_directory, order_name))  # Klasörü oluşturur

    excel_file_path = os.path.join(
        current_directory, order_name, order_name+" "+product_name)  # Excel dosyasının tam yolu

    # Excel dosyasını oluştur
    workbook = excel.Workbooks.Add()
    worksheet = workbook.Worksheets(1)
    worksheet.Range("A:E").VerticalAlignment = EXCEL_VERTICAL_ALIGNMENT_CENTER
    worksheet.Range(
        "A:B").HorizontalAlignment = EXCEL_HORIZONTAL_ALIGNMENT_CENTER
    worksheet.Range(
        "B:C").HorizontalAlignment = EXCEL_HORIZONTAL_ALIGNMENT_LEFT
    worksheet.Range(
        "C:D").HorizontalAlignment = EXCEL_HORIZONTAL_ALIGNMENT_CENTER
    worksheet.Range(
        "D:E").HorizontalAlignment = EXCEL_HORIZONTAL_ALIGNMENT_CENTER
    worksheet.Range("A3:E3").HorizontalAlignment = EXCEL_TEXT_ALIGNMENT_CENTER
    worksheet.Range(
        "A3:E3").VerticalAlignment = EXCEL_VERTICAL_ALIGNMENT_CENTER
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
                elif col == 5:
                    # worksheet.Cells(row, 6).Value = value  # Hücreye değeri yaz
                    if value:
                        worksheet.Cells(row, 2).Interior.Color = value

                col += 1  # Sütunu bir artır
        # Bir sonraki satıra geçmeden önce kontrol et
        if values:
            row += 1  # Satırı bir artır

    # A3'den D'deki en son satıra kadar olan hücrelere kenarlık ekleyin
    add_border_to_range(worksheet, "A1", "E" + str(row - 1))
    worksheet.Columns.AutoFit()

    # A2'de "Notlar" yazısını ekleyin
    add_notes_title(worksheet)

    workbook.SaveAs(excel_file_path)  # Excel dosyasını belirtilen yere kaydet

    forget()

    # Programı 2 saniye sonra kapat
    root.after(1500, lambda: root.destroy())


def forget():
    product_name_label.place_forget()
    product_name_entry.place_forget()
    order_number_label.place_forget()
    order_number_entry.place_forget()
    excel_product_count_label.place_forget()
    excel_product_count_entry.place_forget()
    remove_sheet_metal_checkbox.place_forget()
    create_button.place_forget()
    settings_button.place_forget()
    colors_button.place_forget()
    sheet_remove_button.place_forget()
    liste.place_forget()
    yscrollbar.place_forget()
    root.update()


def create_root():
    root = ThemedTk(theme='adapta', themebg=True)
    window_width = 600
    window_height = 400
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    root.resizable(False, False)

    # Pencereyi ekranın ortasına konumlandır
    x = (screen_width - window_width) // 2
    y = (screen_height - window_height-150) // 2
    root.geometry(f"{window_width}x{window_height}+{x}+{y}")
    root.minsize(window_width, window_height)  # Minimum boyutu ayarla

    root.title("Mil Excel & Pdf Oluşturma")
    return root


def create_warning_label(root):
    # Kırmızı renkli bir stil oluştur
    style = ttk.Style()

    # Stili yalnızca warning_label için kullanmak için stil adına özel bir etiket stili tanımlayın
    style.configure("RedWarning.TLabel", foreground="red")

    # Stili uygulanan bir Label widget'ı oluştur
    warning_label = ttk.Label(root, text="", style="RedWarning.TLabel")
    return warning_label


def handle_home_button():
    home_button.place_forget()
    place()
    settings_button.place(relx=0.9, rely=y-0.1, anchor="center")
    colors_button.place_forget()
    sheet_remove_button.place_forget()
    settings_label.place_forget()
    liste.place_forget()
    remove_button.place_forget()
    add_button.place_forget()
    add_entry.place_forget()


def create_home_button(root):
    home_button = ttk.Button(root, text="🏠", command=handle_home_button)
    return home_button


def handle_settings_button():
    forget()
    # home düğmesini oluştur
    home_button.place(relx=0.9, rely=y-0.1, anchor="center")
    settings_label.place(relx=0.52, rely=y+0.05, anchor="center")
    sheet_remove_button.place(relx=0.4, rely=y+0.3, anchor="center")
    colors_button.place(relx=0.65, rely=y+0.3, anchor="center")
    settings_label.config(text="Ayarlar")


def create_settings_button(root):

    # Ayarlar düğmesini oluştur
    settings_button = ttk.Button(
        root, text="⚙️", command=handle_settings_button)

    # Ayarlar düğmesini ana pencereye yerleştir
    return settings_button


def selectItem(liste):
    # Seçilen öğenin id'sini al
    item_id = liste.focus()

    # Seçilen öğenin değerini al
    item_value = liste.item(item_id, "values")[0]

    # Seçilen öğenin değerini print et
    return item_value


def handle_sheet_remove_button():
    # "Sac Sil" düğmesine tıklandığında yapılacak işlemler
    sheet_remove_button.place_forget()
    colors_button.place_forget()
    settings_label.config(text="Sac Silme Ayarı")
    settings_label.place(relx=0.52, rely=y-0.1, anchor="center")
    liste.place(relx=0.4, rely=0.2, relwidth=0.5, relheight=0.6)
    remove_button.place(relx=0.65, rely=y+0.7, anchor="center")
    add_button.place(relx=0.25, rely=y+0.3, anchor="center")
    add_entry.place(relx=0.25, rely=y+0.2, anchor="center")

    # Bu, scrollbar'ın listenin içinde görünmesini sağlar
    yscrollbar.place(in_=liste, relx=0.95, relheight=1.0)


def add_item_to_json(json_file, item, key):
    # JSON verisini alın
    data = fetch_json_data(json_file)

    if data is not None:
        # Item'i belirtilen anahtarın altındaki listeye ekleyin (eğer öğe henüz eklenmemişse)
        if key in data and isinstance(data[key], list):
            if item not in data[key]:
                data[key].append(item)
                # JSON dosyasına item'i ekleyin (UTF-8 kodlaması kullanarak)
                try:
                    with open(json_file, 'w', encoding='utf-8') as file:
                        json.dump(data, file, indent=4, ensure_ascii=False)
                    return "Öğe başarıyla eklendi!"
                except Exception as e:
                    return f"Hata oluştu: {str(e)}"
            else:
                return "Öğe zaten ekli."
        else:
            # Anahtar yoksa veya anahtar bir liste değilse yeni bir liste oluşturun
            data[key] = [item]
            # JSON dosyasına item'i ekleyin (UTF-8 kodlaması kullanarak)
            try:
                with open(json_file, 'w', encoding='utf-8') as file:
                    json.dump(data, file, indent=4, ensure_ascii=False)
                return "Öğe başarıyla eklendi!"
            except Exception as e:
                return f"Hata oluştu: {str(e)}"
    else:
        return "Veri alınamadı."

# handle_add_button fonksiyonunu kullanırken konsola yazdırmayı unutmayın
def handle_add_button():
    # Kullanıcıdan girdiyi alın
    item = add_entry.get()
    warning_label.place_forget()
    approval_label.place_forget()

    # JSON dosyasına item'i ekleyin (örneğin, 'sacSil.json' dosyasına ekleyin)
    json_file = 'milJsonFiles/sacSil.json'  # JSON dosyasının adını buraya ekleyin
    key = 'words_to_remove'    # Anahtar adını buraya ekleyin
    result = add_item_to_json(json_file, item, key)

    # Kullanıcıya işlem sonucunu gösterin
    if result == "Öğe başarıyla eklendi!":
        approval_label.config(text=result, style="GreenApproval.TLabel")
        approval_label.place(relx=0.25, rely=y + 0.1, anchor="center")
        # 1.5 saniye sonra approval_label'ı gizle
        root.after(1500, lambda: approval_label.place_forget())        
        # Girdi alanını temizle
        add_entry.delete(0, 'end')
        response = fetch_json_data('milJsonFiles/sacSil.json')
        words_to_remove = response["words_to_remove"]
        update_list(liste,words_to_remove)
    elif result == "Öğe zaten ekli.":
        warning_label.config(text=result, style="RedWarning.TLabel")
        warning_label.place(relx=0.25, rely=y + 0.1, anchor="center")
        # 1.5 saniye sonra warning_label'ı gizle
        root.after(1500, lambda: warning_label.place_forget())
    else:
        warning_label.config(text=result, style="RedWarning.TLabel")
        warning_label.place(relx=0.25, rely=y + 0.1, anchor="center")

def update_list(liste, list_items):
    # Liste üzerindeki mevcut öğelerin id'lerini al
    children = liste.get_children()

    # List_items arrayindeki her öğe için
    for i, item in enumerate(list_items):
        # Eğer liste üzerinde karşılık gelen bir öğe varsa
        if i < len(children):
            # Öğenin değerini güncelle
            liste.item(children[i], values=(item))
        else:
            # Yoksa, yeni bir öğe ekle
            liste.insert("", "end", values=(item))

    # Eğer liste üzerinde fazla öğe varsa
    if len(children) > len(list_items):
        # Fazla olan öğeleri sil
        for j in range(len(list_items), len(children)):
            liste.delete(children[j])
    

def create_add_button(root, add_entry):
    style = ttk.Style()
    style.configure("Custom.TButton", font=("Segoe UI", 12))

    add_button = ttk.Button(
        root, text="Ekle", command=handle_add_button, style="Custom.TButton")
    add_button.config(state="disabled")

    def add_check_and_enable_button(event):
        add_entry_text = add_entry.get()
        if not add_entry_text:
            add_button.config(state="disabled")
        else:
            add_button.config(state="normal")

    add_entry.bind("<KeyRelease>", add_check_and_enable_button)

    return add_button


def handle_remove_button():
    item = selectItem(liste)
    print(item)


def create_remove_button(root):
    style = ttk.Style()
    # Segoe UI fontu ve 12 punto olarak ayarla
    style.configure("Custom.TButton", font=("Segoe UI", 12))

    remove_button = ttk.Button(
        root, text="Kaldır", command=handle_remove_button, style="Custom.TButton")
    return remove_button


def create_sheet_remove_button(root):
    style = ttk.Style()
    # Segoe UI fontu ve 12 punto olarak ayarla
    style.configure("Custom.TButton", font=("Segoe UI", 12))

    sheet_remove_button = ttk.Button(
        root, text="Sac Silme", command=handle_sheet_remove_button, style="Custom.TButton")
    return sheet_remove_button


def handle_colors_button():
    # "Colors" düğmesine tıklandığında yapılacak işlemler
    sheet_remove_button.place_forget()
    colors_button.place_forget()
    settings_label.config(text="Renk Ayarı")


def create_colors_button(root):
    style = ttk.Style()
    # Segoe UI fontu ve 12 punto olarak ayarla
    style.configure("Custom.TButton", font=("Segoe UI", 12))
    colors_button = ttk.Button(
        root, text="Renk", command=handle_colors_button, style="Custom.TButton")
    return colors_button


def create_approval_label(root):
    # Yeşil renkli bir stil oluştur
    style = ttk.Style()

    # Stili yalnızca approval_label için kullanmak için stil adına özel bir etiket stili tanımlayın
    style.configure("GreenApproval.TLabel", foreground="green")

    # Stili uygulanan bir Label widget'ı oluştur
    approval_label = ttk.Label(root, text="", style="GreenApproval.TLabel")
    return approval_label



def set_font_style():
    style = ttk.Style()
    style.configure("Custom.TLabel", font=(12))
    return style


def create_product_name_label(root):
    style = set_font_style()
    # Varsayılan stil ile bir Label widget'ı oluştur
    product_name_label = ttk.Label(
        root, text="Ürün Adı:", style="Custom.TLabel")

    return product_name_label


def create_order_number_label(root):
    # Varsayılan stil ile bir Label widget'ı oluştur
    style = set_font_style()

    order_number_label = ttk.Label(
        root, text="Sipariş Numarası:", style="Custom.TLabel")
    return order_number_label


def create_excel_product_count_label(root):
    style = set_font_style()

    # Varsayılan stil ile bir Label widget'ı oluştur
    excel_product_count_label = ttk.Label(
        root, text="Ürün Adeti:", style="Custom.TLabel")
    return excel_product_count_label


def create_product_name_entry(root):
    product_name_entry = ttk.Entry(root)
    return product_name_entry


def create_order_number_entry(root):
    order_number_entry = ttk.Entry(root)
    return order_number_entry


def create_add_entry(root):
    add_entry = ttk.Entry(root)
    return add_entry


def create_remove_sheet_metal_checkbox_entry(root):
    # "Sac Sil" butonuna tıklanıp tıklanmadığını takip eden değişken
    sac_sil_flag = tk.BooleanVar()
    sac_sil_flag.set(False)  # Başlangıçta "Sac Sil" butonu işaretsiz

    style = ttk.Style()
    # Segoe UI fontu ve 12 punto olarak ayarla
    style.configure("Custom.TCheckbutton", font=("Segoe UI", 12))

    remove_sheet_metal_checkbox = ttk.Checkbutton(
        root, text="Sac Sil", variable=sac_sil_flag, style="Custom.TCheckbutton")

    return remove_sheet_metal_checkbox, sac_sil_flag


def create_liste(root, list_items, text_header):
    # Liste penceresini oluştur (show parametresini "headings" olarak ayarla)
    liste = ttk.Treeview(root, columns=("Veriler"), show="headings", height=10)
    liste.heading("#1", text=text_header)

    # Liste öğelerini liste üzerinde görüntüle
    for item in list_items:
        liste.insert("", "end", values=(item))

    # Öğe seçildiğinde çağrılacak işlevi tanımla
    liste.bind("<<TreeviewSelect>>", lambda event: selectItem(liste))

    return liste



def create_yscrollbar(root, liste):
    yscrollbar = ttk.Scrollbar(root, orient="vertical", command=liste.yview)
    liste.configure(yscrollcommand=yscrollbar.set)
    return yscrollbar


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
    # Segoe UI fontu ve 12 punto olarak ayarla
    style.configure("Custom.TButton", font=("Segoe UI", 12))
    create_button = ttk.Button(
        root, text="Oluştur", command=create_excel, style="Custom.TButton")
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


def create_settings_label(root):
    # Özel stil ile bir Label widget'ı oluştur
    style = ttk.Style()
    style.configure("b.TLabel", font=("Segoe UI", 18))

    settings_label = ttk.Label(root, text="", style="b.TLabel")
    return settings_label


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
remove_sheet_metal_checkbox, sac_sil_flag = create_remove_sheet_metal_checkbox_entry(
    root)
add_entry = create_add_entry(root)
# "Oluştur" düğmesi
create_button = create_create_button(root, create_excel)
add_button = create_add_button(root, add_entry)
settings_button = create_settings_button(root)
home_button = create_home_button(root)
remove_button = create_remove_button(root)
# Ayarlar etkieti
settings_label = create_settings_label(root)

# Renkler Butonu
colors_button = create_colors_button(root)
# Sac sil butonu
sheet_remove_button = create_sheet_remove_button(root)


def listfn(root):
    response = fetch_json_data('milJsonFiles/sacSil.json')
    words_to_remove = response["words_to_remove"]
    liste = create_liste(root, words_to_remove, "Sac Sil Kelimeler")
    return liste


liste = listfn(root)
yscrollbar = create_yscrollbar(root, liste)


def place():

    order_number_label.place(relx=0.35, rely=y+0.1, anchor="center")
    order_number_entry.place(relx=0.6, rely=y+0.1, anchor="center")
    product_name_label.place(relx=0.4, rely=y+0.2, anchor="center")
    product_name_entry.place(relx=0.6, rely=y+0.2, anchor="center")
    excel_product_count_label.place(relx=0.39, rely=y+0.3, anchor="center")
    excel_product_count_entry.place(relx=0.6, rely=y+0.3, anchor="center")
    remove_sheet_metal_checkbox.place(relx=0.5, rely=y+0.425, anchor="center")
    create_button.place(relx=0.5, rely=y+0.55, anchor="center")
    settings_button.place(relx=0.9, rely=y-0.1, anchor="center")


place()
# Tkinter penceresini başlat
root.mainloop()
