import tkinter as tk
import win32com.client as win32
import os
import json
from elements.ttkElements import create_button, create_add_button, generate_create_button, item_place, place_list, create_label_with_style, create_entry, create_remove_sheet_metal_checkbox_entry, create_color_liste, create_liste, create_yscrollbar, create_root, create_add_color_button, create_scrolled_text
import subprocess
import sys

# Sabitler
EXCEL_BORDER_STYLE = 1
EXCEL_TEXT_ALIGNMENT_CENTER = -4108
EXCEL_HORIZONTAL_ALIGNMENT_LEFT = -4131
EXCEL_VERTICAL_ALIGNMENT_CENTER = -4108
EXCEL_HORIZONTAL_ALIGNMENT_CENTER = -4108

# Excel başlık verileri
header = ["Malzeme Kodu", "Malzeme Açıklaması",
          "Birim Sarf Miktarı", "Toplam Sarf Miktarı", "Birim"]
idx = None
item_value = None  # seçili rengin hex kodu
sheet_metals_path = r"\\SRV1\planlama\milApp\milJsonFiles\sacSil.json"
colors_path = r"\\SRV1\planlama\milApp\milJsonFiles\renkler.json"
# A2'de "Notlar" yazısını ekleyen fonksiyon


def add_notes_title(excel, worksheet, order_notes):
    worksheet_range = worksheet.Range
    worksheet_cells = worksheet.Cells

    worksheet_cells(2, 1).Value = "Notlar"
    worksheet_cells(2, 1).Font.Bold = True
    # EXCEL_TEXT_ALIGNMENT_CENTER için değeri 3
    worksheet_cells(2, 1).HorizontalAlignment = 3

    worksheet_cells(2, 4).Value = "Ürün Adeti"
    worksheet_cells(2, 4).Font.Bold = True
    # EXCEL_TEXT_ALIGNMENT_CENTER için değeri 3
    worksheet_cells(2, 4).HorizontalAlignment = 3

    # B2 ve C2 hücrelerini birleştir
    merge_range = worksheet_range("B2:C2")
    merge_range.Merge()

    # Eğer order_notes "Sipariş notları.." değilse
    if order_notes != "Sipariş notları..":
        # Birleştirilen hücreye order_notes ekleyin
        merge_range.Value = order_notes
        merge_range.WrapText = True
        worksheet.Rows[2].RowHeight = 185
    else:
        # Eğer order_notes "Sipariş notları.." ise yüksekliği ayarla

        worksheet.Rows[2].RowHeight = 100


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
    response = fetch_json_data(sheet_metals_path)
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


def validate_copied_text(copied_data):
    # Kopyalanan veriyi satırlara böl
    lines = copied_data.split("\n")
    # Sonuncu line'ı atla
    lines = lines[:-1]
    # Her satır için
    for line in lines:
        # Satırı tab karakterine göre böl
        columns = line.split("\t")
        # Eğer kolon sayısı 4 değilse
        if len(columns) != 4:
            # False dön ve fonksiyondan çık
            return False
    # Eğer tüm satırlarda kolon sayısı 4 ise
    # True dön
    return True

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
    response = fetch_json_data(colors_path)
    colors = response["colors"]
    result = []

    lines = text.split("\n")

    multiple_matches = []

    for line in lines:
        values = line.split("\t")
        formatted_line = []

        if len(values) >= 4:
            keyword_to_search = values[1].lower()
            rgb_colors = set()
            matched_keywords = {}

            for color, keywords in colors.items():
                for keyword in keywords:
                    if keyword.startswith("$") and keyword.endswith("$"):
                        full_matched_keyword = keyword[1:-1]
                        full_matched_keyword_lowered = full_matched_keyword.lower()
                        if full_matched_keyword_lowered in keyword_to_search.split():
                            matched_keywords[color] = full_matched_keyword
                            rgb_colors.add(color)

                    else:
                        parts = keyword.lower().split("*")
                        for part in parts:
                            if part.startswith("$") and part.endswith("$"):
                                full_matched_keyword = part[1:-1]
                                full_matched_keyword_lowered = full_matched_keyword.lower()
                                if full_matched_keyword_lowered in keyword_to_search.split():
                                    rgb_colors.add(color)

                                    matched_keywords[color] = full_matched_keyword
                        if all(part in keyword_to_search for part in parts):
                            rgb_colors.add(color)
                            matched_keywords[color] = keyword
                            break

            if len(rgb_colors) > 1:

                multiple_matches.append((keyword_to_search, matched_keywords))

            if not rgb_colors:
                rgb_colors.add("")
            values.append(rgb_colors.pop())

        formatted_line = values
        result.append("\t".join(formatted_line))

    result_excel_format = "\n".join(result)
    sorted_data_by_colora = sort_data_by_color(result_excel_format)

    if multiple_matches:
        message = "Excel ve Pdf oluşturulmayacaktır. Lütfen ayarlardan gerekli ayarlamaları yapın! Birden fazla renk eşleşti:\n"
        for keyword, matched_keywords in multiple_matches:
            message += f"{keyword} için:\n"
            for color, keyword in matched_keywords.items():
                color_mapping = {
                    "8696052": "#F4B084",
                    "11992832": "#00FFB6",
                    "65535": "#FFFF00",
                    "13408767": "#FF99CC",
                    "14395790": "#8AA9DB",
                    "9359529": "#A9D08E",
                    "10092441": "#99FF99"
                }
                color = color_mapping.get(color, color)
                message += f"{color}: {keyword}\n"
        
        write_to_txt(message)
        sys.exit(0)

    return sorted_data_by_colora


def write_to_txt(message):
    # Dosyanın konumunu belirleyin
    dosya_konumu = os.path.join(os.environ["USERPROFILE"], "Desktop")

    # Dosya adını bir değişkene atayın
    dosya_adi = "renkler hata 1"

    # Dosyanın var olup olmadığını kontrol edin
    while os.path.exists(os.path.join(dosya_konumu, dosya_adi + ".txt")):
        # Eğer dosya varsa, dosya adını bir artırın
        dosya_adi = dosya_adi[:-1] + str(int(dosya_adi[-1]) + 1)

    # Dosyayı yazma modunda açın
    with open(os.path.join(dosya_konumu, dosya_adi + ".txt"), "w") as dosya:
        # Dosyaya merhaba yazın
        dosya.write(message)
        # Dosyayı kapatın
        dosya.close()

    # Dosyayı varsayılan uygulamayla açın
    os.startfile(dosya_konumu + "\\" + dosya_adi + ".txt")


def validate_user_inputs(string):
    # Kullanılamayacak semboller listesi
    invalid_chars = ["\\", "/", ":", "*", "?", "\"", "<", ">", "|"]
    # Geçersiz sembollerin tutulacağı bir liste oluştur
    found_chars = []
    # Her bir sembol için
    for char in invalid_chars:
        # Eğer string içinde sembol varsa
        if char in string:
            # Sembolü bulunan listesine ekle
            found_chars.append(char)
    # Eğer bulunan listesi boş değilse
    if found_chars:
        # False ve bulunan listesini dön
        return [False, found_chars]
    # Eğer bulunan listesi boşsa
    else:
        # True ve boş liste dön
        return [True, []]


def create_excel():
    try:
        copied_text = root.clipboard_get()  # Kopyalanan metni al
    except tk.TclError:
        copied_text = ""

    if not copied_text:
        warning_label.config(text="Lütfen ürünleri kopyalayın!")
        item_place(warning_label, 0.5, 0.1)

    elif not validate_copied_text(copied_text):
        warning_label.config(text="Yanlış içerik kopyalanmış!")
        item_place(warning_label, 0.5, 0.1)

        return

    else:
        warning_label.place_forget()  # Label'ı kaldır
        order_name = order_number_entry.get()
        current_directory = os.path.expanduser("~/Desktop")
        product_name = product_name_entry.get()
        excel_product_count = excel_product_count_entry.get()
        order_name_validate = validate_user_inputs(order_name)
        product_name_validate = validate_user_inputs(product_name)
        order_notes = notes_scrolled_text.get("1.0", "end").strip()

        def excel_check():
            approval_label.config(text="Excel ve Pdf oluşturuluyor...")
            create_buttona.config(state="disabled")
            item_place(approval_label, 0.5, 0.1)
            root.update()  # Arayüzü güncelle

            # Klasörü oluştur

            if sac_sil_flag.get():  # Eğer sac sil seçiliyse
                # Kopyalanan metinden belirtilen kelimeleri sil
                cleaned_text = remove_selected_words(copied_text)
                result = apply_colors(cleaned_text)

                create_excelfn(result, order_name, product_name, excel_product_count,
                               excel_file_path, order_notes, pdf_file_path)  # Temizlenmiş veri ile Excel oluştur
            else:  # Eğer sac sil seçili değilse
                # Kopyalanan metni olduğu gibi Excel'e yaz
                # Veriyi temizle
                result = apply_colors(copied_text)
                # Yeni veriyi kopyala
                create_excelfn(result, order_name, product_name,
                               excel_product_count, excel_file_path, order_notes, pdf_file_path)
            approval_label.config(
                text="Excel ve Pdf oluşturuldu!")  # Sonucu göster
            item_place(approval_label, 0.5, 0.4)

        # Eğer order_name veya product_name içinde özel sembol varsa
        if not order_name_validate[0] or not product_name_validate[0]:

            # Eğer order_name içinde özel sembol varsa
            if not order_name_validate[0]:
                # Uyarı mesajında order_name ve bulunan semboller göster
                warning_label.config(
                    text=f"Sipariş numarasında kullanılamayacak semboller var: {', '.join(order_name_validate[1])}", style="RedWarning.TLabel")
                item_place(warning_label, 0.5, 0.1)

            # Eğer product_name içinde özel sembol varsa
            if not product_name_validate[0]:
                # Uyarı mesajında product_name ve bulunan semboller göster
                warning_label.config(
                    text=f"Ürün adında kullanılamayacak semboller var: {', '.join(product_name_validate[1])}", style="RedWarning.TLabel")
                item_place(warning_label, 0.5, 0.1)

            return None

        # Klasör yolu oluşturun
        folder_path = os.path.join(current_directory, order_name)
        excel_file_path = os.path.join(
            current_directory, order_name, order_name+" "+product_name)  # Excel dosyasının tam yolu
        pdf_file_path = os.path.splitext(excel_file_path)[0] + ".pdf"

        # Eğer klasör zaten varsa
        if os.path.exists(folder_path):
            # Uyarı mesajı göster
            if os.path.exists(pdf_file_path) or os.path.exists(excel_file_path):
                warning_label.config(
                    text="Bu isimde excel veya pdf zaten var!", style="RedWarning.TLabel")
                item_place(warning_label, 0.5, 0.1)
                return None
        # Eğer klasör yoksa
            excel_check()

        else:
            os.mkdir(folder_path)

            excel_check()

# Excel dosyasını oluşturmak için fonksiyon


def create_excelfn(copied_text, order_name, product_name, excel_product_count, excel_file_path, order_notes, pdf_file_path):

    # Excel application'ı başlat
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    excel.Visible = True  # Excel penceresini görünür yap

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
    add_notes_title(excel, worksheet, order_notes)
    worksheet.PageSetup.Zoom = False  # ölçeklendirme seçeneğini iptal et
    worksheet.PageSetup.FitToPagesWide = 1  # genişliği bir sayfaya sığdır
    # yüksekliği otomatik ayarla    workbook.SaveAs(excel_file_path)  # Excel dosyasını belirtilen yere kaydet
    worksheet.PageSetup.FitToPagesTall = False
    # Define the PDF file path with the same name as the Excel file
    workbook.SaveAs(excel_file_path)  # Excel dosyasını belirtilen yere kaydet

    # Export the Excel worksheet as a PDF
    worksheet.ExportAsFixedFormat(0, pdf_file_path)
    subprocess.Popen([pdf_file_path], shell=True)

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
    create_buttona.place_forget()
    settings_button.place_forget()
    colors_button.place_forget()
    sheet_remove_button.place_forget()
    liste.place_forget()
    yscrollbar.place_forget()
    notes_scrolled_text.place_forget()
    root.update()


def handle_home_button():
    home_button.place_forget()
    place()
    item_place(settings_button, 0.9, 0.1)
    colors_button.place_forget()
    sheet_remove_button.place_forget()
    settings_label.place_forget()
    liste.place_forget()
    remove_button.place_forget()
    add_button.place_forget()
    add_entry.place_forget()
    color_liste.place_forget()
    liste.delete(*liste.get_children())
    add_color_button.place_forget()
    add_color_entry.place_forget()
    add_color_label.place_forget()
    remove_color_button.place_forget()
    add_button.config(state="disabled")
    add_color_button.config(state="disabled")
    add_color_entry.delete(0, 'end')
    add_entry.delete(0, 'end')


def handle_settings_button():
    forget()
    item_place(home_button, 0.9, 0.1)
    item_place(settings_label, 0.49, 0.25)
    item_place(sheet_remove_button, 0.37, 0.5)
    item_place(colors_button, 0.62, 0.5)
    settings_label.config(text="Ayarlar")


def selectItem(liste):
    # Seçilen öğenin id'sini al
    item_id = liste.focus()

    # Eğer hiçbir öğe seçilmediyse None dön
    if item_id is None:
        return None

    # Seçilen öğenin değerlerini al
    item_values = liste.item(item_id, "values")

    # Eğer değerler boşsa veya boş bir liste ise None dön
    if not item_values:
        return None

    # Değerler listesinin ilk öğesini döndür
    item_value = item_values[0]

    # Seçilen öğenin değerini döndür
    return item_value


def handle_sheet_remove_button():
    # "Sac Sil" düğmesine tıklandığında yapılacak işlemler
    sheet_remove_button.place_forget()
    colors_button.place_forget()
    settings_label.config(text="Sac Silme Ayarı")
    liste.heading("#1", text="Sac Silme Kelimeler")
    response = fetch_json_data(sheet_metals_path)
    words_to_remove = response["words_to_remove"]
    update_list(liste, words_to_remove)
    place_list(liste, 0.4, 0.2, 0.5, 0.6)
    # Bu, scrollbar'ın listenin içinde görünmesini sağlar
    yscrollbar.place(in_=liste, relx=0.95, relheight=1.0)

    item_place(settings_label, 0.52, 0.1)
    item_place(remove_button, 0.65, 0.9)
    item_place(add_button, 0.22, 0.5)
    item_place(add_entry, 0.22, 0.4)


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

# sac ekleme için


def handle_add_button():
    # Kullanıcıdan girdiyi alın
    item = add_entry.get()
    warning_label.place_forget()
    approval_label.place_forget()

    # JSON dosyasına item'i ekleyin (örneğin, 'sacSil.json' dosyasına ekleyin)
    json_file = sheet_metals_path  # JSON dosyasının adını buraya ekleyin
    key = 'words_to_remove'    # Anahtar adını buraya ekleyin
    result = add_item_to_json(json_file, item, key)

    # Kullanıcıya işlem sonucunu gösterin
    if result == "Öğe başarıyla eklendi!":
        approval_label.config(text=result, style="GreenApproval.TLabel")
        item_place(approval_label, 0.25, 0.3)
        # 1.5 saniye sonra approval_label'ı gizle
        root.after(1500, lambda: approval_label.place_forget())
        # Girdi alanını temizle
        add_entry.delete(0, 'end')
        response = fetch_json_data(sheet_metals_path)
        words_to_remove = response["words_to_remove"]
        update_list(liste, words_to_remove)
    elif result == "Öğe zaten ekli.":
        warning_label.config(text=result, style="RedWarning.TLabel")
        item_place(warning_label, 0.25, 0.3)
        add_entry.delete(0, 'end')
        # 1.5 saniye sonra warning_label'ı gizle
        root.after(1500, lambda: warning_label.place_forget())
    else:
        warning_label.config(text=result, style="RedWarning.TLabel")
        item_place(warning_label, 0.25, 0.3)


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


def handle_remove_button():
    selected_item = selectItem(liste)
    warning_label.place_forget()
    approval_label.place_forget()

    if selected_item:
        json_file = sheet_metals_path
        key = 'words_to_remove'
        result = remove_item_from_json(json_file, selected_item, key)
        if result == f"Öğe '{selected_item}' başarıyla silindi.":
            approval_label.config(text=result, style="GreenApproval.TLabel")
            item_place(approval_label, 0.25, 0.3)
            # 1.5 saniye sonra approval_label'ı gizle
            root.after(1500, lambda: approval_label.place_forget())

            response = fetch_json_data(json_file)
            words_to_remove = response.get(key, [])
            update_list(liste, words_to_remove)
        else:
            warning_label.config(text=result, style="RedWarning.TLabel")
            item_place(warning_label, 0.25, 0.3)
            # 1.5 saniye sonra warning_label'ı gizle
            root.after(1500, lambda: warning_label.place_forget())
    else:
        warning_label.config(
            text="Lütfen listeden öğe seçiniz.", style="RedWarning.TLabel")
        item_place(warning_label, 0.25, 0.3)
        # 1.5 saniye sonra warning_label'ı gizle
        root.after(1500, lambda: warning_label.place_forget())


def remove_item_from_json(json_file, item, key):
    data = fetch_json_data(json_file)
    if key in data and isinstance(data[key], list) and item in data[key]:
        data[key].remove(item)
        try:
            with open(json_file, 'w', encoding='utf-8') as file:
                json.dump(data, file, ensure_ascii=False, indent=4)
            return f"Öğe '{item}' başarıyla silindi."
        except Exception as e:
            return f"Hata oluştu: {str(e)}"
    else:
        return f"Öğe '{item}' bulunamadı veya silinemedi."


def handle_colors_button():
    # "Colors" düğmesine tıklandığında yapılacak işlemler
    sheet_remove_button.place_forget()
    colors_button.place_forget()
    settings_label.config(text="Renk Ayarı")
    item_place(settings_label, 0.5, 0.1)

    place_list(color_liste, 0.25, 0.25, 0.5, 0.6)


def extract_last_digit_from_item_id(item_id):
    # item_id'den son karakteri alın ve tamsayıya dönüştürün
    last_digit = int(item_id[-1])
    return last_digit-1  # arrayin 0'dan başlaması dolayısıyla 1 çıkarıldı


def add_item_to_json_with_index(json_file, item):

    # JSON verisini alın
    data = fetch_json_data(json_file)

    if data is not None:
        if "colors" in data and isinstance(data["colors"], dict):
            color_dict = data["colors"]

            if idx is None or idx < 0 or idx >= len(color_dict):
                return "Geçersiz öğe sırası."

            color_keys = list(color_dict.keys())
            color_key_to_add = color_keys[idx]

            if isinstance(color_dict[color_key_to_add], list):
                if item not in color_dict[color_key_to_add]:
                    color_dict[color_key_to_add].append(item)

                    # JSON dosyasına güncellenmiş veriyi yazın (UTF-8 kodlaması kullanarak)
                    try:
                        with open(json_file, 'w', encoding='utf-8') as file:
                            json.dump(data, file, indent=4, ensure_ascii=False)
                        return "Öğe başarıyla eklendi!"
                    except Exception as e:
                        return f"Hata oluştu: {str(e)}"
                else:
                    return "Öğe zaten ekli."
            else:
                return "Belirtilen anahtar bir liste içermiyor."
        else:
            return "colors anahtarı bulunamadı veya bir sözlük değil."
    else:
        return "Veri alınamadı."


def handle_add_color_button():
    # Kullanıcıdan girdiyi alın
    item = add_color_entry.get()
    warning_label.place_forget()
    approval_label.place_forget()

    # JSON dosyasına item'i ekleyin (örneğin, 'sacSil.json' dosyasına ekleyin)
    json_file = colors_path  # JSON dosyasının adını buraya ekleyin

    result = add_item_to_json_with_index(json_file, item)

    # Kullanıcıya işlem sonucunu gösterin
    if result == "Öğe başarıyla eklendi!":
        approval_label.config(text=result, style="GreenApproval.TLabel")
        item_place(approval_label, 0.25, 0.3)
        # 1.5 saniye sonra approval_label'ı gizle
        root.after(1500, lambda: approval_label.place_forget())
        # Girdi alanını temizle
        add_color_entry.delete(0, 'end')
        add_color_button.config(state="disabled")

        update_list_with_index(liste, colors_path, idx)

    elif result == "Öğe zaten ekli.":
        warning_label.config(text=result, style="RedWarning.TLabel")
        item_place(warning_label, 0.25, 0.3)
        add_color_entry.delete(0, 'end')
        # 1.5 saniye sonra warning_label'ı gizle
        root.after(1500, lambda: warning_label.place_forget())
    else:
        warning_label.config(text=result, style="RedWarning.TLabel")
        item_place(warning_label, 0.25, 0.3)


def update_list_with_index(listbox, json_file, idx):
    response = fetch_json_data(json_file)
    if idx >= 0 and idx < len(response["colors"]):
        idx_key_value_array = list(response["colors"].values())[idx]
        update_list(listbox, idx_key_value_array)
        if item_value:
            # Tüm satırların arka plan rengini ayarla
            for row in listbox.get_children():
                listbox.item(row, tags=(item_value))
                listbox.tag_configure(item_value, background=item_value)
    else:
        print("Geçersiz İndeks")


def remove_item_by_color_index(json_file, index, item):
    try:
        data = fetch_json_data(json_file)  # JSON verisini al
        if data is None:
            return "Veri okunamadı veya hata oluştu."

        # Verinin "colors" bölümündeki sıralama indeksine göre erişin
        color_indices = list(data["colors"].keys())
        if index < 0 or index >= len(color_indices):
            return "Geçersiz sıralama indeksi."

        color_index = color_indices[index]

        if item in data["colors"][color_index]:
            data["colors"][color_index].remove(item)

            with open(json_file, 'w', encoding='utf-8') as file:
                json.dump(data, file, ensure_ascii=False, indent=4)

            return f"Öğe '{item}' başarıyla silindi."
        else:
            return f"Öğe '{item}' bulunamadı veya silinemedi."
    except Exception as e:
        return f"Hata oluştu: {str(e)}"


def handle_remove_color_button():
    selected_item = selectItem(liste)
    warning_label.place_forget()
    approval_label.place_forget()

    if selected_item:
        json_file = colors_path
        color_index = idx  # Seçilen öğe, renk indeksi olarak kullanılacak
        result = remove_item_by_color_index(
            json_file, color_index, selected_item)

        if result == f"Öğe '{selected_item}' başarıyla silindi.":
            approval_label.config(text=result, style="GreenApproval.TLabel")
            item_place(approval_label, 0.25, 0.3)
            # 1.5 saniye sonra approval_label'ı gizle
            root.after(1500, lambda: approval_label.place_forget())

            response = fetch_json_data(json_file)
            idx_key_value_array = list(response["colors"].values())[idx]
            update_list(liste, idx_key_value_array)
        else:
            warning_label.config(text=result, style="RedWarning.TLabel")
            item_place(warning_label, 0.25, 0.3)
            # 1.5 saniye sonra warning_label'ı gizle
            root.after(1500, lambda: warning_label.place_forget())
    else:
        warning_label.config(
            text="Lütfen listeden öğe seçiniz.", style="RedWarning.TLabel")
        item_place(warning_label, 0.25, 0.3)
        # 1.5 saniye sonra warning_label'ı gizle
        root.after(1500, lambda: warning_label.place_forget())


def on_select_color(event):
    global idx  # idx'i global değişken olarak tanımla
    global item_value
    # Olayın kaynağı olan liste widget'ını al
    color_liste = event.widget

    # Item ID'den indeksi al
    item_id = color_liste.focus()
    if item_id is None:
        return None

    idx = extract_last_digit_from_item_id(item_id)

    # Önceki sorgudan kalan arka plan rengini temizle
    for row in liste.get_children():
        liste.item(row, tags=())

    # Seçili öğenin değerini al
    item_value = color_liste.item(item_id)['values'][0]

    # İndekse göre liste güncellemesini yap
    update_list_with_index(liste, colors_path, idx)

    # Widget'ları düzenle
    place_list(liste, 0.4, 0.2, 0.5, 0.6)
    item_place(add_color_button, 0.22, 0.5)
    item_place(add_color_entry, 0.22, 0.4)
    item_place(remove_color_button, 0.65, 0.9)
    item_place(add_color_label, 0.22, 0.70)
    yscrollbar.place(in_=liste, relx=0.95, relheight=1.0)
    liste.heading("#1", text="Renkler")
    color_liste.place_forget()


# Tkinter penceresini oluştur
root = create_root()
order_number_entry = create_entry(root, "order_number_entry")
product_name_entry = create_entry(root, "product_name_entry")
add_color_entry = create_entry(root, "add_color_entry")
add_entry = create_entry(root, "add_entry")
excel_product_count_entry = create_entry(root, "excel_product_count_entry")

notes_scrolled_text = create_scrolled_text(root, "notes_scrolled_text", 29, 16)

order_number_entry.focus_set()  # order_number_entry'yi aktif hale getir
# Excel Ürün adeti için giriş alanı

home_button = create_button(root, "🏠", handle_home_button, False)
settings_button = create_button(root, "⚙️", handle_settings_button, False)
add_button = create_add_button(root, handle_add_button, add_entry)
add_color_button = create_add_color_button(
    root, handle_add_color_button, add_color_entry)

remove_button = create_button(root, "Kaldır", handle_remove_button, True)
remove_color_button = create_button(
    root, "Kaldır", handle_remove_color_button, True)

colors_button = create_button(root, "Renkler", handle_colors_button, True)
sheet_remove_button = create_button(
    root, "Sac Silme", handle_sheet_remove_button, True)
create_buttona = generate_create_button(
    root, create_excel, product_name_entry, order_number_entry, excel_product_count_entry)


product_name_label = create_label_with_style(
    root, "Ürün Adı:", "Custom.TLabel")
add_color_label = create_label_with_style(
    root, "Eğer bir satırda birden fazla kelimenin \neşleşmesini istiyorsanız kelime aralarına \n* koyun. Örnek: gözet*cam \n Tam eşleşme sağlamak için kelimenin \nbaşına ve sonuna $ sembolü ekleyin. \nÖrnek: $fan$", "Note.TLabel")
approval_label = create_label_with_style(root, "", "GreenApproval.TLabel")

settings_label = create_label_with_style(root, "", "b.TLabel")

# Uyarı etiketleri
warning_label = create_label_with_style(root, "", "RedWarning.TLabel")


excel_product_count_label = create_label_with_style(
    root, "Ürün Adeti:", "Custom.TLabel")

order_number_label = create_label_with_style(
    root, "Sipariş Numarası:", "Custom.TLabel")


color_liste = create_color_liste(root, on_select_color)


# "Sac Sil" butonunu ve durumunu al
remove_sheet_metal_checkbox, sac_sil_flag = create_remove_sheet_metal_checkbox_entry(
    root)
# Ayarlar etkieti


def listfn(root):
    response = fetch_json_data(sheet_metals_path)
    if response is not None:
        words_to_remove = response.get("words_to_remove")
        if words_to_remove:
            liste = create_liste(root, words_to_remove,
                                 "Sac Sil Kelimeler", selectItem)
            return liste
    return None


liste = listfn(root)

if liste is not None:
    yscrollbar = create_yscrollbar(root, liste)


def place():

    item_place(order_number_label, 0.15, 0.3)
    item_place(order_number_entry, 0.4, 0.3)
    item_place(product_name_label, 0.2, 0.4)
    item_place(product_name_entry, 0.4, 0.4)
    item_place(notes_scrolled_text, 0.75, 0.5)
    item_place(excel_product_count_label, 0.19, 0.5)
    item_place(excel_product_count_entry, 0.4, 0.5)
    item_place(remove_sheet_metal_checkbox, 0.3, 0.625)
    item_place(create_buttona, 0.3, 0.75)
    item_place(settings_button, 0.9, 0.1)


place()
# Tkinter penceresini başlat
root.mainloop()
# pyinstaller --onefile --noconsole --name Mil --icon=mil_icon.ico app.py
