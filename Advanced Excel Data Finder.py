import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import re
import os
from concurrent.futures import ThreadPoolExecutor, as_completed

class ExcelDataFinder:
    def __init__(self, root):
        self.root = root
        self.root.title("Advanced Excel Data Finder")
        self.root.geometry("500x400")
        
        self.setup_ui()

    def setup_ui(self):
        # Search frame
        search_frame = ttk.LabelFrame(self.root, text="Search")
        search_frame.pack(padx=10, pady=10, fill='x')

        ttk.Label(search_frame, text="Search term:").grid(row=0, column=0, padx=5, pady=5, sticky='w')
        self.search_entry = ttk.Entry(search_frame, width=40)
        self.search_entry.grid(row=0, column=1, padx=5, pady=5)

        # Options frame
        options_frame = ttk.LabelFrame(self.root, text="Options")
        options_frame.pack(padx=10, pady=10, fill='x')

        self.case_sensitive_var = tk.BooleanVar()
        ttk.Checkbutton(options_frame, text="Case sensitive", variable=self.case_sensitive_var).grid(row=0, column=0, padx=5, pady=5, sticky='w')

        self.whole_word_var = tk.BooleanVar()
        ttk.Checkbutton(options_frame, text="Whole word only", variable=self.whole_word_var).grid(row=1, column=0, padx=5, pady=5, sticky='w')

        ttk.Label(options_frame, text="Search in:").grid(row=2, column=0, padx=5, pady=5, sticky='w')
        self.search_in_var = tk.StringVar(value="all")
        ttk.Radiobutton(options_frame, text="All columns", variable=self.search_in_var, value="all").grid(row=2, column=1, padx=5, pady=5, sticky='w')
        ttk.Radiobutton(options_frame, text="Specific columns", variable=self.search_in_var, value="specific").grid(row=3, column=1, padx=5, pady=5, sticky='w')
        
        self.columns_entry = ttk.Entry(options_frame, width=30)
        self.columns_entry.grid(row=3, column=2, padx=5, pady=5)
        ttk.Label(options_frame, text="(comma-separated)").grid(row=3, column=3, padx=5, pady=5, sticky='w')

        # Buttons
        button_frame = ttk.Frame(self.root)
        button_frame.pack(pady=10)

        ttk.Button(button_frame, text="Select Files", command=self.select_files).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Search", command=self.search_files).pack(side=tk.LEFT, padx=5)

        # Results
        self.result_tree = ttk.Treeview(self.root, columns=("File", "Sheet", "Row", "Column", "Value"), show="headings")
        self.result_tree.heading("File", text="File")
        self.result_tree.heading("Sheet", text="Sheet")
        self.result_tree.heading("Row", text="Row")
        self.result_tree.heading("Column", text="Column")
        self.result_tree.heading("Value", text="Value")
        self.result_tree.pack(padx=10, pady=10, fill='both', expand=True)

        # Scrollbar for results
        scrollbar = ttk.Scrollbar(self.root, orient='vertical', command=self.result_tree.yview)
        scrollbar.pack(side='right', fill='y')
        self.result_tree.configure(yscrollcommand=scrollbar.set)

        self.status_var = tk.StringVar()
        ttk.Label(self.root, textvariable=self.status_var).pack(pady=5)

        self.file_paths = []

    def select_files(self):
        self.file_paths = filedialog.askopenfilenames(filetypes=[("Excel files", "*.xlsx *.xls")])
        if self.file_paths:
            self.status_var.set(f"{len(self.file_paths)} file(s) selected")
        else:
            self.status_var.set("No files selected")

    def search_files(self):
        search_term = self.search_entry.get()
        if not search_term:
            messagebox.showwarning("Warning", "Please enter a search term.")
            return
        if not self.file_paths:
            messagebox.showwarning("Warning", "No files selected.")
            return

        # Clear previous results
        for i in self.result_tree.get_children():
            self.result_tree.delete(i)

        case_sensitive = self.case_sensitive_var.get()
        whole_word = self.whole_word_var.get()
        search_in = self.search_in_var.get()
        specific_columns = [col.strip() for col in self.columns_entry.get().split(',')] if search_in == 'specific' else None

        if not case_sensitive:
            search_term = search_term.lower()

        if whole_word:
            search_term = fr'\b{re.escape(search_term)}\b'

        self.status_var.set("Searching...")
        self.root.update()

        with ThreadPoolExecutor() as executor:
            futures = []
            for file_path in self.file_paths:
                futures.append(executor.submit(self.search_file, file_path, search_term, case_sensitive, specific_columns))

            for future in as_completed(futures):
                results = future.result()
                for result in results:
                    self.result_tree.insert('', 'end', values=result)

        self.status_var.set("Search completed")

    def search_file(self, file_path, search_term, case_sensitive, specific_columns):
        results = []
        try:
            xl = pd.ExcelFile(file_path)
            for sheet_name in xl.sheet_names:
                df = xl.parse(sheet_name)
                columns_to_search = specific_columns if specific_columns else df.columns
                for col in columns_to_search:
                    if col in df.columns:
                        for index, value in df[col].items():
                            cell_value = str(value)
                            if not case_sensitive:
                                cell_value = cell_value.lower()
                            if re.search(search_term, cell_value):
                                results.append((os.path.basename(file_path), sheet_name, index + 2, col, value))
        except Exception as e:
            messagebox.showerror("Error", f"Error reading file {file_path}: {str(e)}")
        return results

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelDataFinder(root)
    root.mainloop()