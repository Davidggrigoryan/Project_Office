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

        self._build_ui()

    def _build_ui(self):
        frame = ttk.Frame(self.master, padding=10)
        frame.grid(row=0, column=0, sticky="nsew")

        # File selectors
        ttk.Label(frame, text="Отчёт 1:").grid(row=0, column=0, sticky="w")
        ttk.Entry(frame, textvariable=self.file1_path, width=60).grid(row=0, column=1)
        ttk.Button(frame, text="Выбрать", command=self.load_file1).grid(row=0, column=2)

        ttk.Label(frame, text="Отчёт 2:").grid(row=1, column=0, sticky="w")
        ttk.Entry(frame, textvariable=self.file2_path, width=60).grid(row=1, column=1)
        ttk.Button(frame, text="Выбрать", command=self.load_file2).grid(row=1, column=2)

        # Column mapping widgets
        ttk.Label(frame, text="Столбец проекта (Отчёт 1):").grid(row=2, column=0, sticky="w")
        self.project_col1 = tk.StringVar()
        self.project_menu1 = ttk.Combobox(frame, textvariable=self.project_col1,
                                          values=self.column_options1, state="readonly", width=40)
        self.project_menu1.grid(row=2, column=1)

        ttk.Label(frame, text="Столбец проекта (Отчёт 2):").grid(row=3, column=0, sticky="w")
        self.project_col2 = tk.StringVar()
        self.project_menu2 = ttk.Combobox(frame, textvariable=self.project_col2,
                                          values=self.column_options2, state="readonly", width=40)
        self.project_menu2.grid(row=3, column=1)

        ttk.Label(frame, text="Столбец значения (Отчёт 1):").grid(row=4, column=0, sticky="w")
        self.value_col1 = tk.StringVar()
        self.value_menu1 = ttk.Combobox(frame, textvariable=self.value_col1,
                                        values=self.column_options1, state="readonly", width=40)
        self.value_menu1.grid(row=4, column=1)

        ttk.Label(frame, text="Столбец значения (Отчёт 2):").grid(row=5, column=0, sticky="w")
        self.value_col2 = tk.StringVar()
        self.value_menu2 = ttk.Combobox(frame, textvariable=self.value_col2,
                                        values=self.column_options2, state="readonly", width=40)
        self.value_menu2.grid(row=5, column=1)

        # Output file selector
        ttk.Label(frame, text="Файл результата:").grid(row=6, column=0, sticky="w")
        ttk.Entry(frame, textvariable=self.output_path, width=60).grid(row=6, column=1)
        ttk.Button(frame, text="Сохранить как", command=self.choose_output).grid(row=6, column=2)

        # Process button
        ttk.Button(frame, text="Вычислить разницу", command=self.process).grid(row=7, column=1, pady=10)

    def load_file1(self):
        path = filedialog.askopenfilename(filetypes=[("Excel-файлы", "*.xlsx *.xls")])
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
        self.value_menu1['values'] = self.column_options1

    def load_file2(self):
        path = filedialog.askopenfilename(filetypes=[("Excel-файлы", "*.xlsx *.xls")])
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
        self.value_menu2['values'] = self.column_options2

    def choose_output(self):
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel-файлы", "*.xlsx")])
        if path:
            self.output_path.set(path)

    def process(self):
        if not all([self.df1 is not None, self.df2 is not None,
                    self.project_col1.get(), self.project_col2.get(),
                    self.value_col1.get(), self.value_col2.get(),
                    self.output_path.get()]):
            messagebox.showwarning("Не все поля заполнены", "Пожалуйста, загрузите файлы, выберите все столбцы и файл для сохранения.")
            return

        try:
            merged = self.df1.merge(
                self.df2,
                left_on=self.project_col1.get(),
                right_on=self.project_col2.get(),
                suffixes=("_1", "_2")
            )

            proj_col = self.project_col1.get()
            val1 = self.value_col1.get()
            val2 = self.value_col2.get()

            val1_col = val1 if val1 != val2 else f"{val1}_1"
            val2_col = val2 if val1 != val2 else f"{val2}_2"

            result = merged[[proj_col, val1_col, val2_col]].copy()
            result.rename(columns={
                proj_col: "Проект",
                val1_col: val1,
                val2_col: val2,
            }, inplace=True)
            result["Разница"] = result[val1] - result[val2]
            result.to_excel(self.output_path.get(), index=False)
        except Exception as exc:
            messagebox.showerror("Ошибка", f"Не удалось обработать данные: {exc}")
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


def main():
    root = tk.Tk()
    app = ExcelDiffApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
