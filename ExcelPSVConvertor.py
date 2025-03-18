import os
import tkinter as tk
from tkinter import filedialog, scrolledtext
import openpyxl
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
        
        window_width = 600
        window_height = 280
        
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        
        center_x = int((screen_width - window_width) / 2)
        center_y = int((screen_height - window_height) / 2)
        
        self.root.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')
        
        try:
            self.root.wm_iconphoto(False, tk.PhotoImage(file='assets/icon.png'))
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
        
        # Create input listbox with scrollbar in a frame
        self.input_frame = tk.Frame(self.root)
        self.input_scrollbar = tk.Scrollbar(self.input_frame, orient=tk.VERTICAL)
        self.input_listbox = tk.Listbox(self.input_frame, height=10, width=40, 
                                        yscrollcommand=self.input_scrollbar.set)
        self.input_scrollbar.config(command=self.input_listbox.yview)
        self.input_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.input_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Create output listbox with scrollbar in a frame
        self.output_frame = tk.Frame(self.root)
        self.output_scrollbar = tk.Scrollbar(self.output_frame, orient=tk.VERTICAL)
        self.output_listbox = tk.Listbox(self.output_frame, height=10, width=40, fg="green", 
                                         yscrollcommand=self.output_scrollbar.set)
        self.output_scrollbar.config(command=self.output_listbox.yview)
        self.output_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.output_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.status_frame = tk.Frame(self.root, relief=tk.SUNKEN, bd=1)
        self.status_bar = tk.Label(self.status_frame, text="Ready", anchor=tk.W, padx=5, pady=2)
        self.status_bar.pack(fill=tk.X)
        
    def _layout_widgets(self):
        self.input_label.grid(row=0, column=0, sticky="w")
        self.input_path.grid(row=0, column=1, padx=(0, 5), sticky="e")
        self.output_label.grid(row=0, column=2, padx=(5, 0), sticky="w")
        self.output_path.grid(row=0, column=3, padx=(0, 0), sticky="e")
        
        # Grid the frames containing listboxes instead of the listboxes directly
        self.input_frame.grid(row=1, column=0, columnspan=2, padx=(0, 5), pady=5, sticky="ew")
        self.output_frame.grid(row=1, column=2, columnspan=2, padx=(5, 0), pady=5, sticky="ew")
        
        self.browse_button.grid(row=0, column=0, padx=(0, 5), pady=(0, 0), sticky="ew")
        self.delete_button.grid(row=0, column=1, padx=(5, 0), pady=(0, 0), sticky="ew")
        
        self.convert_button.grid(row=0, column=0, pady=(0, 0), sticky="ew")
        
        self.browse_delete_frame.columnconfigure(0, weight=1)
        self.browse_delete_frame.columnconfigure(1, weight=1)
        self.convert_frame.columnconfigure(0, weight=1)
        
        self.browse_delete_frame.grid(row=2, column=0, columnspan=2, pady=(5, 5), padx=(0, 5), sticky="ew")
        self.convert_frame.grid(row=2, column=2, columnspan=2, pady=(5, 5), padx=(5, 0), sticky="ew")

        self.status_frame.grid(row=3, column=0, columnspan=4, pady=(5, 0), padx=(0, 0), sticky="ew")
        
        self.root.columnconfigure(0, weight=1)
        self.root.columnconfigure(1, weight=1)
        self.root.columnconfigure(2, weight=1)
        self.root.columnconfigure(3, weight=1)
    
    def update_status(self, message, color="black"):
        current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.status_bar.config(text=f"{current_time}: {message}", fg=color)
    
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
                
            self.update_status(f"{len(file_paths)} files selected", "blue")
            
        except Exception as e:
            self.update_status(f"Error: {str(e)}", "red")

    def delete_files(self):
        selected_indices = self.input_listbox.curselection()
        if not selected_indices:
            self.update_status("No files selected to delete", "red")
            return
            
        for index in sorted(selected_indices, reverse=True):
            if index < len(self.file_path_list):
                file_name = os.path.basename(self.file_path_list[index])
                del self.file_path_list[index]
                self.input_listbox.delete(index)
                self.update_status(f"'{file_name}' removed from selection", "red")


    def convert_files(self):
        if not self.file_path_list:
            self.update_status("No files selected", "red")
            return
            
        os.makedirs(self.output_folder, exist_ok=True)
        
        file_count = 0
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
                file_count += 1
                self.update_status(f"Converting: {file_name}", "blue")
                
            except Exception as e:
                self.update_status(f"Error converting '{os.path.basename(input_file)}'", "red")

        self.update_status(f"Conversion completed: {file_count} files converted", "green")
        self.file_path_list = []
        self.input_listbox.delete(0, tk.END)

        messagebox.showinfo("Conversion Completed", "All files have been successfully converted.")
    
    def _convert_csv_to_psv(self, input_file, output_file):
        with open(input_file, 'r', newline='', encoding='utf-8') as csv_file:
            csv_reader = csv.reader(csv_file)
            with open(output_file, 'w', newline='', encoding='utf-8') as psv_file:
                psv_writer = csv.writer(psv_file, delimiter='|')
                for row in csv_reader:
                    psv_writer.writerow(row)
    
    def _convert_excel_to_psv(self, input_file, output_file):
        try:
            workbook = openpyxl.load_workbook(input_file, data_only=True)
            
            sheet = workbook.active
            
            with open(output_file, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f, delimiter='|', quoting=csv.QUOTE_MINIMAL)
                
                for row in sheet.rows:
                    row_values = []
                    for cell in row:
                        if cell.value is None:
                            row_values.append('')
                        else:
                            row_values.append(str(cell.value))
                    writer.writerow(row_values)
            
            workbook.close()
            return True
        except Exception as e:
            self.update_status(f"Excel conversion error: {str(e)}", "red")
            return False


def main():
    root = tk.Tk()
    app = ExcelPsvConverter(root)
    root.mainloop()


if __name__ == "__main__":
    main()