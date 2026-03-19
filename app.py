import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import tkinter as tk
from tkinter import messagebox
import chardet
import warnings
warnings.filterwarnings("ignore")

import os
import sys

def get_path(filename):
    if getattr(sys, 'frozen', False):
        return os.path.join(os.path.dirname(sys.executable), filename)
    return filename

# --- универсальное чтение файла ---
def read_file_safely(file_path):
    # 👉 если Excel
    if file_path.endswith(".xlsx") or file_path.endswith(".xls"):
        print("📊 Читаем Excel файл")
        return pd.read_excel(file_path)

    # 👉 если CSV
    with open(file_path, 'rb') as f:
        rawdata = f.read(100000)
        result = chardet.detect(rawdata)
        encoding = result['encoding']

    print(f"📄 Кодировка: {encoding}")

    try:
        return pd.read_csv(file_path, encoding=encoding, sep=None, engine='python')
    except:
        pass

    for enc in ["utf-8", "cp1252", "latin1"]:
        try:
            return pd.read_csv(file_path, encoding=enc, sep=None, engine='python')
        except:
            continue

    # последний шанс
    with open(file_path, encoding="latin1", errors="ignore") as f:
        return pd.read_csv(f, sep=None, engine='python')


# --- настройки ---
FILE_PATH = get_path("СВОД_затрат.xlsx")
SPREADSHEET_NAME = "СВОД_Ударник"
WORKSHEET_NAME = "Лист1"
CREDENTIALS_FILE = get_path("citric-cubist-490713-j3-4e852ad3799b.json")


# --- обновление таблицы ---
def update_sheet():
    try:
        scope = [
            "https://spreadsheets.google.com/feeds",
            "https://www.googleapis.com/auth/drive"
        ]

        creds = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, scope)
        client = gspread.authorize(creds)

        sheet = client.open(SPREADSHEET_NAME).worksheet(WORKSHEET_NAME)

        # 👉 читаем файл
        df = read_file_safely(FILE_PATH)

        # 👉 убираем NaN (важно для Google Sheets)
        df = df.fillna("")

        # 👉 обновляем таблицу
        sheet.clear()
        sheet.update([df.columns.values.tolist()] + df.values.tolist())

        messagebox.showinfo("Готово", "✅ Таблица обновлена!")

    except Exception as e:
        messagebox.showerror("Ошибка", str(e))


# --- интерфейс ---
root = tk.Tk()
root.title("Обновление Google Sheets")
root.geometry("300x150")

label = tk.Label(root, text="Обновить таблицу из файла", font=("Arial", 12))
label.pack(pady=10)

btn = tk.Button(root, text="Обновить", command=update_sheet, height=2, width=20)
btn.pack(pady=10)

root.mainloop()
