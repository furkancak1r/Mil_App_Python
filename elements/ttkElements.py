import tkinter as tk
from tkinter import ttk
from ttkthemes import ThemedTk

"""ROOT"""
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



"""BUTTONS"""

def create_button(parent, text, command_function, check_style=True):
    style = ttk.Style()
    style.configure("Custom.TButton", font=("Segoe UI", 12))

    button = ttk.Button(parent, text=text, command=command_function,
                        style="Custom.TButton" if check_style else "TButton")
    return button


def create_add_button(root, command_function, add_entry):
    add_button = create_button(root, "Ekle", command_function)
    add_button.config(state="disabled")

    def add_check_and_enable_button(event):
        add_entry_text = add_entry.get()
        add_button.config(state="normal" if add_entry_text else "disabled")

    add_entry.bind("<KeyRelease>", add_check_and_enable_button)

    return add_button


def generate_create_button(root, create_excel, product_name_entry, order_number_entry, excel_product_count_entry):
    create_buttona = create_button(root, "Oluştur", create_excel)
    create_buttona.config(state="disabled")

    def check_and_enable_button(event):
        product_name = product_name_entry.get()
        order_number = order_number_entry.get()
        excel_product_count = excel_product_count_entry.get()
        create_buttona.config(
            state="normal" if order_number and product_name and excel_product_count else "disabled")

    order_number_entry.bind("<KeyRelease>", check_and_enable_button)
    product_name_entry.bind("<KeyRelease>", check_and_enable_button)
    excel_product_count_entry.bind("<KeyRelease>", check_and_enable_button)

    return create_buttona

"""PLACE"""
def item_place(item, relx, rely):
    item.place(relx=relx, rely=rely, anchor="center")


def place_list(liste, relx, rely, relwidth, relheight):
    liste.place(relx=relx, rely=rely, relwidth=relwidth, relheight=relheight)


"""LABEL"""
def create_label_with_style(parent, text, style_name):
    style = ttk.Style()
    style.configure("RedWarning.TLabel", foreground="red")
    style.configure("b.TLabel", font=("Segoe UI", 18))
    style.configure("GreenApproval.TLabel", foreground="green")
    style.configure("Custom.TLabel", font=(12))


    label = ttk.Label(parent, text=text, style=style_name)
    return label

"""ENTRY"""
def create_entry(root, entry_name):
    def validate_input(P):
        if P == "" or P.isdigit():
            return True
        else:
            return False

    if entry_name == "excel_product_count_entry":
        validate = "key"
        validatecommand = (root.register(validate_input), "%P")
    else:
        validate = None
        validatecommand = None
    
    entry = ttk.Entry(root, name=entry_name, validate=validate, validatecommand=validatecommand)
    return entry

"""CHECKBUTTON"""
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

"""LISTS"""
def create_color_liste(root,on_select_color):
    # Tkinter penceresi oluşturun
    color_codes = ['#F4B084', '#00FFB6', '#FFFF00',
                   '#FF99CC', '#8AA9DB', '#A9D08E', '#99FF99']

    # Liste penceresini oluştur (show parametresini "headings" olarak ayarla)
    color_liste = ttk.Treeview(root, columns=(
        "Renkler"), show="headings", height=len(color_codes))
    color_liste.heading("#1", text="Renkler")

    # Liste öğelerini liste üzerinde görüntüle
    for color_code in color_codes:
        # Renk kodunu kullanarak arka plan rengini ayarlayın
        color_liste.insert("", "end", values=(color_code), tags=(color_code))
        color_liste.tag_configure(color_code, background=color_code)

    # Öğe seçildiğinde çağrılacak işlevi tanımla
    color_liste.bind("<<TreeviewSelect>>", on_select_color)

    return color_liste

def create_liste(root, list_items, text_header,selectItem):
    # Liste penceresini oluştur (show parametresini "headings" olarak ayarla)
    liste = ttk.Treeview(root, columns=("Veriler"), show="headings", height=10)
    liste.heading("#1", text=text_header)

    # Liste öğelerini liste üzerinde görüntüle
    for item in list_items:
        liste.insert("", "end", values=(item))

    # Öğe seçildiğinde çağrılacak işlevi tanımla
    liste.bind("<<TreeviewSelect>>", lambda event: selectItem(liste))

    return liste

"""SCROLLBAR"""
def create_yscrollbar(root, liste):
    yscrollbar = ttk.Scrollbar(root, orient="vertical", command=liste.yview)
    liste.configure(yscrollcommand=yscrollbar.set)
    return yscrollbar