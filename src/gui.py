import tkinter as tk
from tkinter import filedialog, messagebox
from report import generate_report
from pathlib import Path

def select_input():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    input_entry.delete(0, tk.END)
    input_entry.insert(0, file_path)

def select_output():
    file_path = filedialog.asksaveasfilename(defaultextension=".xlsx")
    output_entry.delete(0, tk.END)
    output_entry.insert(0, file_path)

def run_report():
    input_path = Path(input_entry.get())
    output_path = Path(output_entry.get())

    if not input_path.exists():
        messagebox.showerror("Hata", "Geçerli bir giriş dosyası seçin.")
        return

    generate_report(input_path, output_path)
    messagebox.showinfo("Başarılı", "Rapor oluşturuldu!")

root = tk.Tk()
root.title("Excel AutoReport Pro")
root.geometry("400x200")

tk.Label(root, text="Giriş Excel Dosyası").pack()
input_entry = tk.Entry(root, width=50)
input_entry.pack()
tk.Button(root, text="Seç", command=select_input).pack()

tk.Label(root, text="Çıkış Rapor Dosyası").pack()
output_entry = tk.Entry(root, width=50)
output_entry.pack()
tk.Button(root, text="Kaydet", command=select_output).pack()

tk.Button(root, text="Rapor Oluştur", command=run_report).pack(pady=10)

root.mainloop()
