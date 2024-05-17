import tkinter as tk
from tkinter import messagebox, filedialog
import openpyxl
import os
import sys

def input_to_excel(file_path, data):
    wb = openpyxl.Workbook()
    sheet = wb.active

    for row_idx, row_data in enumerate(data):
        for col_idx, value in enumerate(row_data):
            sheet.cell(row=row_idx + 1, column=col_idx + 1).value = value

    wb.save(file_path)
    print(f"Data berhasil diinput ke {file_path}")

def submit_data():
    input_data = text_input.get("1.0", "end").strip()
    lines = input_data.split("\n")
    data = [line.split(",") for line in lines]

    if not excel_file_path:
        messagebox.showerror("Error", "Silakan pilih path untuk file Excel.")
    else:
        input_to_excel(excel_file_path, data)
        messagebox.showinfo("Sukses", "Data berhasil dimasukkan ke dalam file Excel.")

def run_as_admin():
    if os.name == 'nt': 
        try:
            if sys.argv[-1] != 'as_admin': 
                os.system(f'powershell -Command "Start-Process \\"python\\" -ArgumentList \\"{" ".join(sys.argv)} as_admin\\" -Verb RunAs"')
                sys.exit(0)
            else:
                create_gui() 
        except Exception as e:
            print("Gagal menjalankan program sebagai administrator:", e)
            sys.exit(1)
    else:
        create_gui() 

def browse_file():
    global excel_file_path
    excel_file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])

def create_gui():
    global root, text_input, submit_button, excel_file_path
    root = tk.Tk()
    root.title("Form Input Data")

    tk.Label(root, text="Path File Excel:").grid(row=0, column=0)
    tk.Button(root, text="Browse", command=browse_file).grid(row=0, column=1)

    text_input = tk.Text(root, height=10, width=40)
    text_input.grid(row=1, column=0, columnspan=2)

    submit_button = tk.Button(root, text="Submit", command=submit_data)
    submit_button.grid(row=2, column=0, columnspan=2)

    root.mainloop()

run_as_admin()
