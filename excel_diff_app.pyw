import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd


class ExcelDiffApp:
    def __init__(self, master):
        self.master = master
        master.title("Сравнение Excel")

        self.file1_path = tk.StringVar()
        self.file2_path = tk.StringVar()
        self.output_path = tk.StringVar()

        self.df1 = None
        self.df2 = None

        self.column_options1 = []
        self.column_options2 = []
        self.column_pairs = []

        self.list1 = None
        self.list2 = None
        self.pair_list = None

        self._build_ui()

    def _build_ui(self):
        frame = ttk.Frame(self.master, padding=10)
        frame.grid(row=0, column=0, sticky="nsew")

        # File selectors
        ttk.Label(frame, text="Отчёт 1:").grid(row=0, column=0, sticky="w")
        ttk.Entry(frame, textvariable=self.file1_path, width=60).grid(row=0, column=1)
        ttk.Button(frame, text="Обзор", command=self.load_file1).grid(row=0, column=2)

        ttk.Label(frame, text="Отчёт 2:").grid(row=1, column=0, sticky="w")
        ttk.Entry(frame, textvariable=self.file2_path, width=60).grid(row=1, column=1)
        ttk.Button(frame, text="Обзор", command=self.load_file2).grid(row=1, column=2)

        # Column pair widgets
        pair_frame = ttk.Frame(frame)
        pair_frame.grid(row=2, column=0, columnspan=3, pady=5)

        ttk.Label(pair_frame, text="Столбцы отчёта 1").grid(row=0, column=0)
        ttk.Label(pair_frame, text="Столбцы отчёта 2").grid(row=0, column=1)

        self.list1 = tk.Listbox(pair_frame, exportselection=False, height=15, width=30)
        self.list1.grid(row=1, column=0, padx=5)

        self.list2 = tk.Listbox(pair_frame, exportselection=False, height=15, width=30)
        self.list2.grid(row=1, column=1, padx=5)

        ttk.Button(pair_frame, text="Добавить пару", command=self.add_pair).grid(row=1, column=2, padx=5)

        self.pair_list = tk.Listbox(pair_frame, height=10, width=60)
        self.pair_list.grid(row=2, column=0, columnspan=2, pady=5)

        # Output file selector
        ttk.Label(frame, text="Выходной файл:").grid(row=3, column=0, sticky="w")
        ttk.Entry(frame, textvariable=self.output_path, width=60).grid(row=3, column=1)
        ttk.Button(frame, text="Сохранить как", command=self.choose_output).grid(row=3, column=2)

        # Process button
        ttk.Button(frame, text="Вычислить разницу", command=self.process).grid(row=4, column=1, pady=10)

    def load_file1(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if not path:
            return
        try:
            self.df1 = pd.read_excel(path)
        except Exception as exc:
            messagebox.showerror("Ошибка", f"Не удалось загрузить файл: {exc}")
            return
        self.file1_path.set(path)
        self.column_options1 = [c for c in self.df1.columns if not str(c).lower().startswith('unnamed')]
        self.list1.delete(0, tk.END)
        for col in self.column_options1:
            self.list1.insert(tk.END, col)

    def load_file2(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if not path:
            return
        try:
            self.df2 = pd.read_excel(path)
        except Exception as exc:
            messagebox.showerror("Ошибка", f"Не удалось загрузить файл: {exc}")
            return
        self.file2_path.set(path)
        self.column_options2 = [c for c in self.df2.columns if not str(c).lower().startswith('unnamed')]
        self.list2.delete(0, tk.END)
        for col in self.column_options2:
            self.list2.insert(tk.END, col)

    def add_pair(self):
        sel1 = self.list1.curselection()
        sel2 = self.list2.curselection()
        if not sel1 or not sel2:
            return
        c1 = self.list1.get(sel1[0])
        c2 = self.list2.get(sel2[0])
        self.column_pairs.append((c1, c2))
        self.pair_list.insert(tk.END, f"{c1} <-> {c2}")

    def choose_output(self):
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if path:
            self.output_path.set(path)

    def process(self):
        if not all([self.df1 is not None, self.df2 is not None,
                    self.column_pairs, self.output_path.get()]):
            messagebox.showwarning("Недостаточно данных", "Загрузите отчёты, выберите столбцы и выходной файл.")
            return

        try:
            proj_col1, proj_col2 = self.column_pairs[0]
            merged = self.df1.merge(
                self.df2,
                left_on=proj_col1,
                right_on=proj_col2,
                suffixes=("_1", "_2")
            )
            proj_col = proj_col1 + ("_1" if proj_col1 in self.df2.columns else "")
            result = merged[[proj_col]].copy()
            result.rename(columns={proj_col: proj_col1}, inplace=True)
            for c1, c2 in self.column_pairs[1:]:
                c1_name = c1 + ("_1" if c1 in self.df2.columns else "")
                c2_name = c2 + ("_2" if c2 in self.df1.columns else "")
                result[f"{c1} - {c2}"] = merged[c1_name] - merged[c2_name]
            result.to_excel(self.output_path.get(), index=False)
        except Exception as exc:
            messagebox.showerror("Ошибка", f"Не удалось обработать: {exc}")
            return

        self.show_result(result)

    def show_result(self, df):
        win = tk.Toplevel(self.master)
        win.title("Результат")

        tree = ttk.Treeview(win, columns=list(df.columns), show='headings')
        for col in df.columns:
            tree.heading(col, text=col)
            tree.column(col, width=150)
        for _, row in df.iterrows():
            tree.insert('', tk.END, values=list(row))
        tree.pack(fill=tk.BOTH, expand=True)


if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelDiffApp(root)
    root.mainloop()
