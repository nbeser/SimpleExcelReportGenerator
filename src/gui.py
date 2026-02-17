import tkinter as tk
from tkinter import filedialog, messagebox
from report import generate_report
from pathlib import Path

def select_input():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    input_entry.delete(0, tk.END)
    input_entry.insert(0, file_path)

def select_output():
    file_path = filedialog.asksaveasfilename(
        title="Save Report As",
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx")],
    )
    output_entry.delete(0, tk.END)
    output_entry.insert(0, file_path)

def run_report():
    input_path = Path(input_entry.get())
    output_path = Path(output_entry.get())

    if not input_path.exists():
        message_label.config(text="❌ Please select a valid input file.", fg="red")
        return

    if output_path.suffix.lower() != ".xlsx":
        message_label.config(text="❌ Output must end with .xlsx", fg="red")
        return

    try:
        generate_report(input_path, output_path)
        message_label.config(text="✅ Report Created Successfully!", fg="green")
    except Exception as e:
        message_label.config(text=f"❌ Error: {str(e)}", fg="red")

# GUI SETUP

root = tk.Tk()
root.title("Excel AutoReport Pro")
root.geometry("600x300")  
root.resizable(False, False) 

frame = tk.Frame(root, padx=20, pady=20)
frame.pack(expand=True, fill="both")

tk.Label(frame, text="Input Excel File (.xlsx)", anchor="w").pack(fill="x")
input_entry = tk.Entry(frame, width=60)
input_entry.pack(pady=5)
tk.Button(frame, text="Select Input", command=select_input).pack(pady=5)

tk.Label(frame, text="Output Excel File (.xlsx)", anchor="w").pack(fill="x")
output_entry = tk.Entry(frame, width=60)
output_entry.pack(pady=5)
tk.Button(frame, text="Select Output", command=select_output).pack(pady=5)

tk.Button(frame, text="Generate Report", command=run_report).pack(pady=15)

message_label = tk.Label(frame, text="", font=("Arial", 10))
message_label.pack()

root.mainloop()
