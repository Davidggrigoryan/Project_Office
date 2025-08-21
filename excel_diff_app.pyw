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

        # Column mapping widgets
        ttk.Label(frame, text="Столбец проекта (Отчёт 1):").grid(row=2, column=0, sticky="w")
        self.project_col1 = tk.StringVar()
        self.project_menu1 = ttk.Combobox(frame, textvariable=self.project_col1, values=self.column_options1, state="readonly", width=57)
        self.project_menu1.grid(row=2, column=1)

        ttk.Label(frame, text="Столбец проекта (Отчёт 2):").grid(row=3, column=0, sticky="w")
        self.project_col2 = tk.StringVar()
        self.project_menu2 = ttk.Combobox(frame, textvariable=self.project_col2, values=self.column_options2, state="readonly", width=57)
        self.project_menu2.grid(row=3, column=1)

        ttk.Button(frame, text="Пары столбцов...", command=self.open_pair_window).grid(row=4, column=1, pady=5)

        # Output file selector
        ttk.Label(frame, text="Выходной файл:").grid(row=5, column=0, sticky="w")
        ttk.Entry(frame, textvariable=self.output_path, width=60).grid(row=5, column=1)
        ttk.Button(frame, text="Сохранить как", command=self.choose_output).grid(row=5, column=2)

        # Process button
        ttk.Button(frame, text="Вычислить разницу", command=self.process).grid(row=6, column=1, pady=10)

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
        self.column_options1 = list(self.df1.columns)
        self.project_menu1['values'] = self.column_options1
        self.column_pairs.clear()

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
        self.column_options2 = list(self.df2.columns)
        self.project_menu2['values'] = self.column_options2
        self.column_pairs.clear()

    def choose_output(self):
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if path:
            self.output_path.set(path)

    def process(self):
        if not all([self.df1 is not None, self.df2 is not None,
                    self.project_col1.get(), self.project_col2.get(),
                    self.column_pairs, self.output_path.get()]):
            messagebox.showwarning("Недостаточно данных", "Загрузите отчёты, выберите столбцы и выходной файл.")
            return

        try:
            merged = self.df1.merge(
                self.df2,
                left_on=self.project_col1.get(),
                right_on=self.project_col2.get(),
                suffixes=("_1", "_2")
            )
            proj_name = self.project_col1.get()
            proj_col = proj_name + ("_1" if proj_name in self.df2.columns else "")
            result = merged[[proj_col]].copy()
            result.rename(columns={proj_col: proj_name}, inplace=True)
            for c1, c2 in self.column_pairs:
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

    def open_pair_window(self):
        if not (self.column_options1 and self.column_options2):
            messagebox.showwarning("Нет данных", "Сначала загрузите оба отчёта.")
            return

        win = tk.Toplevel(self.master)
        win.title("Пары столбцов")

        list1 = tk.Listbox(win, exportselection=False, height=15, width=30)
        for col in self.column_options1:
            list1.insert(tk.END, col)
        list1.grid(row=0, column=0, padx=5, pady=5)

        list2 = tk.Listbox(win, exportselection=False, height=15, width=30)
        for col in self.column_options2:
            list2.insert(tk.END, col)
        list2.grid(row=0, column=1, padx=5, pady=5)

        pair_list = tk.Listbox(win, height=10, width=60)
        pair_list.grid(row=1, column=0, columnspan=2, padx=5, pady=5)

        def add_pair():
            sel1 = list1.curselection()
            sel2 = list2.curselection()
            if not sel1 or not sel2:
                return
            c1 = list1.get(sel1[0])
            c2 = list2.get(sel2[0])
            self.column_pairs.append((c1, c2))
            pair_list.insert(tk.END, f"{c1} <-> {c2}")

        ttk.Button(win, text="Добавить пару", command=add_pair).grid(row=0, column=2, padx=5)
        ttk.Button(win, text="Закрыть", command=win.destroy).grid(row=1, column=2, padx=5, pady=5, sticky="s")


if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelDiffApp(root)
    root.mainloop()
