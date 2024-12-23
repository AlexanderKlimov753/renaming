import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import os
import pandas as pd
import openpyxl

def convert_xls_to_xlsx(file_path, folder_path):
    new_file_path = os.path.join(folder_path, "data.xlsx")
    df = pd.read_excel(file_path)
    df.to_excel(new_file_path, index=False)
    return new_file_path

def rename_files(folder_path, nomenclature_column, current_name_column):
    try:
        wb = openpyxl.load_workbook(os.path.join(folder_path, "data.xlsx"))
        sheet = wb.active
        for filename in os.listdir(folder_path):
            if filename.endswith(".rsw"):
                old_file_name = os.path.join(folder_path, filename)
                current_name = filename.split(".")[0]
                for row in sheet.iter_rows(min_row=2, values_only=True):
                    current_name_value = row[ord(current_name_column) - 65]
                    nomenclature_value = row[ord(nomenclature_column) - 65]
                    if current_name_value and nomenclature_value and current_name_value.lower() == current_name.lower():
                        new_file_name = os.path.join(folder_path, f"{nomenclature_value}.rsw")
                        os.rename(old_file_name, new_file_name)
                        break
        messagebox.showinfo("Готово", "Файлы успешно переименованы!")
    except Exception as e:
        messagebox.showerror("Ошибка", str(e))

def choose_excel_file():
    excel_file_path = filedialog.askopenfilename(filetypes=[("Файлы Excel", "*.xls")])
    if excel_file_path:
        folder_path = filedialog.askdirectory()
        if folder_path:
            nomenclature_column = nomenclature_entry.get().upper()
            current_name_column = current_name_entry.get().upper()
            converted_file_path = convert_xls_to_xlsx(excel_file_path, folder_path)
            rename_files(folder_path, nomenclature_column, current_name_column)

def choose_folder():
    folder_path = filedialog.askdirectory()
    if folder_path:
        rename_files(folder_path, None, None)

root = tk.Tk()
root.title("Выбор файлов")
root.geometry("300x250")

nomenclature_label = tk.Label(root, text="Колонка Nomenclature:")
nomenclature_label.pack(pady=5)
nomenclature_entry = tk.Entry(root)
nomenclature_entry.pack(pady=5)

current_name_label = tk.Label(root, text="Колонка текущего имени файла:")
current_name_label.pack(pady=5)
current_name_entry = tk.Entry(root)
current_name_entry.pack(pady=5)

excel_button = tk.Button(root, text="Выбрать Excel файл", command=choose_excel_file)
excel_button.pack(pady=10)

folder_button = tk.Button(root, text="Выбрать папку", command=choose_folder)
folder_button.pack(pady=10)

root.mainloop()