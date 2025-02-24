import os
import tkinter as tk
from tkinter import filedialog
import csv
import pandas as pd

def browse_file():
    file_path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv"), ("Excel Files", "*.xlsx;*.xls")])
    if file_path:
        entry.delete(0, tk.END)
        entry.insert(0, file_path)

def reset_app():
    entry.delete(0, tk.END)
    status_label.config(text="")

def convert_file():
    input_file = entry.get()
    if not input_file:
        status_label.config(text="No file selected")
        return

    documents_folder = os.path.join(os.path.expanduser("~"), "Documents")
    output_folder = os.path.join(documents_folder, "PSV_Files")
    os.makedirs(output_folder, exist_ok=True)
    output_file = os.path.join(output_folder, os.path.basename(input_file).rsplit('.', 1)[0] + '_converted.csv')
    
    try:
        if input_file.endswith('.csv'):
            with open(input_file, 'r', newline='') as csv_file:
                csv_reader = csv.reader(csv_file)
                with open(output_file, 'w', newline='') as psv_file:
                    psv_writer = csv.writer(psv_file, delimiter='|')
                    for row in csv_reader:
                        psv_writer.writerow(row)
        elif input_file.endswith('.xlsx') or input_file.endswith('.xls'):
            df = pd.read_excel(input_file)
            df.to_csv(output_file, sep='|', index=False)
        status_label.config(text=f"File converted successfully")
        root.after(5000, reset_app)  # Reset the application after 5 seconds
    except Exception as e:
        status_label.config(text=f"Error during conversion: {e}")

root = tk.Tk()
root.title("CSV/Excel to PSV Converter")
root.geometry("300x105")

frame = tk.Frame(root, padx=10, pady=10)
frame.pack(fill=tk.BOTH, expand=True)

entry = tk.Entry(frame, width=40)
entry.grid(row=1, column=0, padx=5, pady=5, sticky="ew")

browse_button = tk.Button(frame, text="Browse", command=browse_file)
browse_button.grid(row=1, column=1, padx=5, pady=5)

convert_button = tk.Button(frame, text="Convert", command=convert_file, width=40)
convert_button.grid(row=2, column=0, columnspan=2, pady=3)

status_label = tk.Label(frame, text="")
status_label.grid(row=3, column=0, columnspan=2, pady=0)

frame.columnconfigure(0, weight=1)
frame.columnconfigure(1, weight=0)

root.mainloop()