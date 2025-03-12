import os
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
import csv
import pandas as pd

def browse_file():
    OutputListBox.delete(0, tk.END)
    try:
        file_paths = filedialog.askopenfilenames(
            title="Select files",
            filetypes=[("Excel Files", "*.xlsx;*.xls"), ("CSV Files", "*.csv")]
        )
        file_dir = os.path.dirname(file_paths[0])
        InputPath.config(text=file_dir)
        global output_folder
        output_folder = file_dir + '/Output'
        OutputPath.config(text=output_folder)
        if file_paths:
            InputListBox.delete(0, tk.END)
            global file_path_list
            file_path_list = list(file_paths)    
            for count, file_path in enumerate(file_path_list):
                file_name = os.path.basename(file_path)
                InputListBox.insert(count, file_name)
            InputListBox.insert(tk.END,"Selected all files")
    except Exception as e:
        print(f"An error occurred: {e}")

def convert_files(): 
    global file_path_list
    for count, file in enumerate(file_path_list):
        input_file = file
        if not file:
            print("No file selected")
            return
        os.makedirs(output_folder, exist_ok=True)
        output_file = os.path.join(output_folder, os.path.basename(input_file).rsplit('.', 1)[0] + '.csv')
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
            print(f"File converted successfully")
            file_name = os.path.basename(output_file)
            OutputListBox.insert(count, file_name)
        except Exception as e:
            print(f"Error during conversion: {e}")
    OutputListBox.insert(tk.END, "Conversion completed")
    file_path_list = []
    InputListBox.delete(0, tk.END)

# root window
root = tk.Tk()

# configure the root window
root.resizable(False, False)
root.title('PSV v2')
root.configure(padx=10, pady=10)  # Add padding to the entire window


BrowseButton = tk.Button(root, text='Browse', command=browse_file)
ConvertButton = tk.Button(root, text='Convert', command=convert_files)
InputLabel = tk.Label(root, text='Input', fg="blue")
InputPath = tk.Label(root, text='')
InputListBox = tk.Listbox(root, height=10, width=40)
OutputLabel = tk.Label(root, text='Output', fg="red")
OutputPath = tk.Label(root, text='')
OutputListBox = tk.Listbox(root, height=10, width=40)

progressbar = ttk.Progressbar()

StatusLabel = tk.Label(root, text='Status: ')

def row_1(row):
    InputLabel.grid(row=row, column=0, sticky="w")
    InputPath.grid(row=row, column=1, padx=(0, 5), sticky="e")
    OutputLabel.grid(row=row, column=2, padx=(5, 0), sticky="w")
    OutputPath.grid(row=row, column=3, padx=(0, 0), sticky="e")

def row_2(row):
    InputListBox.grid(row=row, column=0, columnspan=2, padx=(0, 5), pady=5, sticky="ew")
    OutputListBox.grid(row=row, column=2, columnspan=2, padx=(5, 0), pady=5, sticky="ew")

def row_3(row):
    BrowseButton.grid(row=row, column=0, columnspan=2, padx=(0,5), pady=(5, 5), sticky="ew")
    ConvertButton.grid(row=row, column=2, columnspan=2, padx=(5,0), pady=(5, 5), sticky="ew")

def row_4(row):
    StatusLabel.grid(row=row, column=0, pady=(5, 5), sticky="w")
    progressbar.grid(row=row, column=1, columnspan=3, pady=(5, 5), sticky="ew")


row_1(0)
row_2(1)
row_3(2)
# row_4(3)


root.mainloop()
