import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk  # pip install pillow
import pandas as pd
import os

def convert_csv_to_xlsx():
    csv_file = filedialog.askopenfilename(
        title="Select CSV File",
        filetypes=[("CSV files", "*.csv")]
    )
    if not csv_file:
        return

    try:
        df = pd.read_csv(csv_file)
        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile=os.path.splitext(os.path.basename(csv_file))[0] + ".xlsx"
        )
        if save_path:
            df.to_excel(save_path, index=False)
            messagebox.showinfo("Success", f"File saved as: {save_path}")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to convert file:\n{e}")

# GUI setup
root = tk.Tk()
root.title("CSV to XLSX Converter")
root.geometry("420x450")
root.resizable(False, False)

# Load logo
logo_image = Image.open("nct_logo.png")
logo_image = logo_image.resize((150, 150), Image.LANCZOS)
logo = ImageTk.PhotoImage(logo_image)
logo_label = tk.Label(root, image=logo)
logo_label.pack(pady=10)

# Ghi chú dưới logo
note = tk.Label(root, text="💡 Chuyển file dữ liệu từ csv qua Excel 👇", font=("Arial", 10))
note.pack(pady=2)

# Nút chuyển đổi
btn_convert = tk.Button(root, text="Select CSV File", command=convert_csv_to_xlsx, font=("Arial", 12))
btn_convert.pack(pady=15)

# Ghi bản quyền
copyright = tk.Label(
    root,
    text="| Copyright 2025 | 🧠 Nghiên Cứu Thuốc |",
    font=("Arial", 9),
    fg="gray"
)
copyright.pack(side="bottom", pady=10)

root.mainloop()
