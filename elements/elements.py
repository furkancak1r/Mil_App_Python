import tkinter as tk
from tkinter import ttk

def create_button(parent, text, command_function, check_style=True):
    style = ttk.Style()
    style.configure("Custom.TButton", font=("Segoe UI", 12))

    button = ttk.Button(parent, text=text, command=command_function, style="Custom.TButton" if check_style else "TButton")
    return button

def item_place(item, relx, rely):
    item.place(relx=relx, rely=rely, anchor="center")

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
        create_buttona.config(state="normal" if order_number and product_name and excel_product_count else "disabled")

    order_number_entry.bind("<KeyRelease>", check_and_enable_button)
    product_name_entry.bind("<KeyRelease>", check_and_enable_button)
    excel_product_count_entry.bind("<KeyRelease>", check_and_enable_button)

    return create_buttona
