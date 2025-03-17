import os
import tkinter as tk
from tkinter import filedialog
import csv
import pandas as pd
from tkinter import messagebox
from datetime import datetime


class ExcelPsvConverter:
    def __init__(self, root):
        self.root = root
        self.file_path_list = []
        self.output_folder = ""
        
        self._setup_window()
        self._create_widgets()
        self._layout_widgets()
        
    def _setup_window(self):
        self.root.title('Excel PSV Converter')
        self.root.configure(padx=10, pady=10)
        self.root.resizable(False, False)
        
        window_width = 500
        window_height = 400
        
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        
        center_x = int((screen_width - window_width) / 2)
        center_y = int((screen_height - window_height) / 2)
        
        self.root.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')
        
        try:
            self.root.wm_iconphoto(False, tk.PhotoImage(file='icon.png'))
        except Exception as e:
            print(f"Could not load icon: {e}")
    
    def _create_widgets(self):
        self.browse_delete_frame = tk.Frame(self.root)
        self.convert_frame = tk.Frame(self.root)
        
        self.browse_button = tk.Button(self.browse_delete_frame, text='Browse', command=self.browse_file)
        self.delete_button = tk.Button(self.browse_delete_frame, text='Delete', command=self.delete_files)
        self.convert_button = tk.Button(self.convert_frame, text='Convert', command=self.convert_files)
        
        self.input_label = tk.Label(self.root, text='Input', fg="blue")
        self.input_path = tk.Label(self.root, text='')
        self.output_label = tk.Label(self.root, text='Output', fg="red")
        self.output_path = tk.Label(self.root, text='')
        
        self.input_listbox = tk.Listbox(self.root, height=10, width=40)
        self.output_listbox = tk.Listbox(self.root, height=10, width=40, fg="green")

        self.log = tk.Text(self.root, width=40)
        
    def _layout_widgets(self):
        self.input_label.grid(row=0, column=0, sticky="w")
        self.input_path.grid(row=0, column=1, padx=(0, 5), sticky="e")
        self.output_label.grid(row=0, column=2, padx=(5, 0), sticky="w")
        self.output_path.grid(row=0, column=3, padx=(0, 0), sticky="e")
        
        self.input_listbox.grid(row=1, column=0, columnspan=2, padx=(0, 5), pady=5, sticky="ew")
        self.output_listbox.grid(row=1, column=2, columnspan=2, padx=(5, 0), pady=5, sticky="ew")
        
        self.browse_button.grid(row=0, column=0, padx=(0, 5), pady=(0, 0), sticky="ew")
        self.delete_button.grid(row=0, column=1, padx=(5, 0), pady=(0, 0), sticky="ew")
        
        self.convert_button.grid(row=0, column=0, pady=(0, 0), sticky="ew")
        
        self.browse_delete_frame.columnconfigure(0, weight=1)
        self.browse_delete_frame.columnconfigure(1, weight=1)
        self.convert_frame.columnconfigure(0, weight=1)
        
        self.browse_delete_frame.grid(row=2, column=0, columnspan=2, pady=(5, 5), padx=(0, 5), sticky="ew")
        self.convert_frame.grid(row=2, column=2, columnspan=2, pady=(5, 5), padx=(5, 0), sticky="ew")

        self.log.grid(row=3, column=0, columnspan=4, pady=(5, 0), padx=(0, 0), sticky="nsew")
        
        self.root.columnconfigure(0, weight=1)
        self.root.columnconfigure(1, weight=1)
        self.root.columnconfigure(2, weight=1)
        self.root.columnconfigure(3, weight=1)
        self.root.rowconfigure(3, weight=1)
    
    def log_datetime(self, message, color="black"):
        current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.log.tag_configure(color, foreground=color)
        self.log.insert(tk.END, f"{current_time}: {message}\n", color)
    
    def browse_file(self):
        self.output_listbox.delete(0, tk.END)
        
        try:
            file_paths = filedialog.askopenfilenames(
                title="Select files",
                filetypes=[("Excel Files", "*.xlsx;*.xls"), ("CSV Files", "*.csv")]
            )
            
            if not file_paths:
                return
                
            file_dir = os.path.dirname(file_paths[0])
            self.input_path.config(text=file_dir)
            
            self.output_folder = os.path.join(file_dir + '/Output')

            self.output_path.config(text=self.output_folder)
            
            self.input_listbox.delete(0, tk.END)
            self.file_path_list = list(file_paths)
            
            for count, file_path in enumerate(self.file_path_list):
                file_name = os.path.basename(file_path)
                self.input_listbox.insert(count, " " + file_name)
                
            self.log_datetime(f"{len(file_paths)} files selected", "blue")
            
        except Exception as e:
            print(f"An error occurred: {e}")

    def delete_files(self):
        selected_indices = self.input_listbox.curselection()
        if not selected_indices:
            self.log_datetime("No files selected to delete", "red")
            return
            
        for index in sorted(selected_indices, reverse=True):
                if index < len(self.file_path_list):
                    file_name = os.path.basename(self.file_path_list[index])
                    del self.file_path_list[index]
                    self.input_listbox.delete(index)
                    self.log_datetime(f"'{file_name}' removed from selection", "red")


    def convert_files(self):
        if not self.file_path_list:
            self.log_datetime("No files selected", "red")
            return
            
        os.makedirs(self.output_folder, exist_ok=True)
        
        for count, input_file in enumerate(self.file_path_list):
            output_file = os.path.join(
                self.output_folder, 
                os.path.basename(input_file).rsplit('.', 1)[0] + '.csv'
            )
            
            try:
                if input_file.endswith('.csv'):
                    self._convert_csv_to_psv(input_file, output_file)
                elif input_file.endswith('.xlsx') or input_file.endswith('.xls'):
                    self._convert_excel_to_psv(input_file, output_file)
                
                file_name = os.path.basename(output_file)
                self.output_listbox.insert(count, " " + file_name)
                
            except Exception as e:
                self.log_datetime(f"Error converting '{os.path.basename(input_file)}'", "red")

        self.log_datetime("Conversion completed", "green")
        self.file_path_list = []
        self.input_listbox.delete(0, tk.END)

        messagebox.showinfo("Conversion Completed", "All files have been successfully converted.")
    
    def _convert_csv_to_psv(self, input_file, output_file):
        with open(input_file, 'r', newline='') as csv_file:
            csv_reader = csv.reader(csv_file)
            with open(output_file, 'w', newline='') as psv_file:
                psv_writer = csv.writer(psv_file, delimiter='|')
                for row in csv_reader:
                    psv_writer.writerow(row)
    
    def _convert_excel_to_psv(self, input_file, output_file):
        df = pd.read_excel(input_file)
        df.to_csv(output_file, sep='|', index=False)


def main():
    root = tk.Tk()
    app = ExcelPsvConverter(root)
    root.mainloop()


if __name__ == "__main__":
    main()