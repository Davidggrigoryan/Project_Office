import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd


class ExcelDiffApp:
    def __init__(self, master):
        self.master = master
        master.title("Excel Diff App")

        self.file1_path = tk.StringVar()
        self.file2_path = tk.StringVar()
        self.output_path = tk.StringVar()

        self.df1 = None
        self.df2 = None

        self.column_options1 = []
        self.column_options2 = []

        self._build_ui()

    def _build_ui(self):
        frame = ttk.Frame(self.master, padding=10)
        frame.grid(row=0, column=0, sticky="nsew")

        # File selectors
        ttk.Label(frame, text="Report 1:").grid(row=0, column=0, sticky="w")
        ttk.Entry(frame, textvariable=self.file1_path, width=40).grid(row=0, column=1)
        ttk.Button(frame, text="Browse", command=self.load_file1).grid(row=0, column=2)

        ttk.Label(frame, text="Report 2:").grid(row=1, column=0, sticky="w")
        ttk.Entry(frame, textvariable=self.file2_path, width=40).grid(row=1, column=1)
        ttk.Button(frame, text="Browse", command=self.load_file2).grid(row=1, column=2)

        # Column mapping widgets
        ttk.Label(frame, text="Project column (Report 1):").grid(row=2, column=0, sticky="w")
        self.project_col1 = tk.StringVar()
        self.project_menu1 = ttk.Combobox(frame, textvariable=self.project_col1, values=self.column_options1, state="readonly")
        self.project_menu1.grid(row=2, column=1)

        ttk.Label(frame, text="Project column (Report 2):").grid(row=3, column=0, sticky="w")
        self.project_col2 = tk.StringVar()
        self.project_menu2 = ttk.Combobox(frame, textvariable=self.project_col2, values=self.column_options2, state="readonly")
        self.project_menu2.grid(row=3, column=1)

        ttk.Label(frame, text="Value column (Report 1):").grid(row=4, column=0, sticky="w")
        self.value_col1 = tk.StringVar()
        self.value_menu1 = ttk.Combobox(frame, textvariable=self.value_col1, values=self.column_options1, state="readonly")
        self.value_menu1.grid(row=4, column=1)

        ttk.Label(frame, text="Value column (Report 2):").grid(row=5, column=0, sticky="w")
        self.value_col2 = tk.StringVar()
        self.value_menu2 = ttk.Combobox(frame, textvariable=self.value_col2, values=self.column_options2, state="readonly")
        self.value_menu2.grid(row=5, column=1)

        # Output file selector
        ttk.Label(frame, text="Output file:").grid(row=6, column=0, sticky="w")
        ttk.Entry(frame, textvariable=self.output_path, width=40).grid(row=6, column=1)
        ttk.Button(frame, text="Save As", command=self.choose_output).grid(row=6, column=2)

        # Process button
        ttk.Button(frame, text="Compute Difference", command=self.process).grid(row=7, column=1, pady=10)

    def load_file1(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if not path:
            return
        try:
            self.df1 = pd.read_excel(path)
        except Exception as exc:
            messagebox.showerror("Error", f"Failed to load file: {exc}")
            return
        self.file1_path.set(path)
        self.column_options1 = list(self.df1.columns)
        self.project_menu1['values'] = self.column_options1
        self.value_menu1['values'] = self.column_options1

    def load_file2(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if not path:
            return
        try:
            self.df2 = pd.read_excel(path)
        except Exception as exc:
            messagebox.showerror("Error", f"Failed to load file: {exc}")
            return
        self.file2_path.set(path)
        self.column_options2 = list(self.df2.columns)
        self.project_menu2['values'] = self.column_options2
        self.value_menu2['values'] = self.column_options2

    def choose_output(self):
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if path:
            self.output_path.set(path)

    def process(self):
        if not all([self.df1 is not None, self.df2 is not None,
                    self.project_col1.get(), self.project_col2.get(),
                    self.value_col1.get(), self.value_col2.get(),
                    self.output_path.get()]):
            messagebox.showwarning("Incomplete", "Please load files and select all columns and output file.")
            return

        try:
            merged = self.df1.merge(
                self.df2,
                left_on=self.project_col1.get(),
                right_on=self.project_col2.get(),
                suffixes=("_1", "_2")
            )
            merged['Difference'] = merged[self.value_col1.get()] - merged[self.value_col2.get()]
            result = merged[[self.project_col1.get(), 'Difference']]
            result.to_excel(self.output_path.get(), index=False)
        except Exception as exc:
            messagebox.showerror("Error", f"Processing failed: {exc}")
            return

        self.show_result(result)

    def show_result(self, df):
        win = tk.Toplevel(self.master)
        win.title("Result")

        tree = ttk.Treeview(win, columns=list(df.columns), show='headings')
        for col in df.columns:
            tree.heading(col, text=col)
            tree.column(col, width=100)
        for _, row in df.iterrows():
            tree.insert('', tk.END, values=list(row))
        tree.pack(fill=tk.BOTH, expand=True)


if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelDiffApp(root)
    root.mainloop()
