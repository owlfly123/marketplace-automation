import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import numpy as np

# --- helpers ---

def parse_int_list(text: str):
    # "65,63,53" -> [65,63,53]
    parts = [p.strip() for p in text.replace(";", ",").split(",") if p.strip()]
    cols = []
    for p in parts:
        if not p.isdigit():
            raise ValueError(f"Некорректный номер колонки: '{p}'")
        n = int(p)
        if n <= 0:
            raise ValueError(f"Номер колонки должен быть > 0: {n}")
        cols.append(n)
    if not cols:
        raise ValueError("Список колонок для суммирования пуст.")
    # remove duplicates keeping order
    seen = set()
    out = []
    for c in cols:
        if c not in seen:
            out.append(c)
            seen.add(c)
    return out

def to_number_series(s: pd.Series) -> pd.Series:
    """
    Устойчивый парсер чисел:
    - убирает пробелы/неразрывные пробелы
    - поддерживает "123 456,78"
    - поддерживает "123,456.78" (запятые как разделитель тысяч)
    """
    s = s.astype("string")

    # normalize spaces
    s = s.str.replace("\u00A0", "", regex=False).str.replace(" ", "", regex=False)

    # keep digits, minus, comma, dot
    s = s.str.replace(r"[^0-9,\.\-]", "", regex=True)

    has_comma = s.str.contains(",", na=False)
    has_dot = s.str.contains(r"\.", na=False)

    # if both comma and dot -> assume comma thousands separators
    s = s.mask(has_comma & has_dot, s.str.replace(",", "", regex=False))
    # if comma only -> comma is decimal separator
    s = s.mask(has_comma & ~has_dot, s.str.replace(",", ".", regex=False))

    out = pd.to_numeric(s, errors="coerce").fillna(0.0)
    return out

def format_results(total_sums: dict[int, float], files_count: int) -> str:
    lines = []
    lines.append(f"Файлов обработано: {files_count}")
    lines.append("")
    lines.append("ИТОГО (сумма по всем файлам):")
    for col_num in sorted(total_sums.keys()):
        val = total_sums[col_num]
        # формат: с 2 знаками после запятой, без научной нотации
        lines.append(f"Колонка {col_num}: {val:,.2f}".replace(",", " "))
    return "\n".join(lines)

# --- app ---

class WBSummarizerApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("WB Summarizer (по номерам колонок)")

        self.files = []

        frm = tk.Frame(root, padx=10, pady=10)
        frm.pack(fill="both", expand=True)

        # files controls
        top = tk.Frame(frm)
        top.pack(fill="x")

        tk.Button(top, text="Выбрать файлы Excel…", command=self.pick_files).pack(side="left")
        tk.Button(top, text="Очистить список", command=self.clear_files).pack(side="left", padx=6)

        self.lbl_files = tk.Label(top, text="Файлы не выбраны")
        self.lbl_files.pack(side="left", padx=10)

        # parameters
        params = tk.LabelFrame(frm, text="Параметры", padx=10, pady=10)
        params.pack(fill="x", pady=10)

        tk.Label(params, text="Колонка с типом документа (1-based):").grid(row=0, column=0, sticky="w")
        self.ent_doc_col = tk.Entry(params, width=10)
        self.ent_doc_col.insert(0, "18")
        self.ent_doc_col.grid(row=0, column=1, sticky="w", padx=8)

        tk.Label(params, text="Значение для возврата:").grid(row=0, column=2, sticky="w")
        self.ent_return_value = tk.Entry(params, width=18)
        self.ent_return_value.insert(0, "Возврат")
        self.ent_return_value.grid(row=0, column=3, sticky="w", padx=8)

        tk.Label(params, text="Колонки для суммирования (1-based, через запятую):").grid(row=1, column=0, sticky="w", pady=(8, 0))
        self.ent_sum_cols = tk.Entry(params, width=40)
        self.ent_sum_cols.insert(0, "65,63,53")
        self.ent_sum_cols.grid(row=1, column=1, columnspan=3, sticky="we", padx=8, pady=(8, 0))

        params.grid_columnconfigure(3, weight=1)

        # actions
        actions = tk.Frame(frm)
        actions.pack(fill="x", pady=(0, 10))

        tk.Button(actions, text="Посчитать суммы", command=self.run).pack(side="left")
        tk.Button(actions, text="Скопировать результат", command=self.copy_result).pack(side="left", padx=6)

        # output
        outbox = tk.LabelFrame(frm, text="Результат", padx=10, pady=10)
        outbox.pack(fill="both", expand=True)

        self.txt = tk.Text(outbox, height=14, wrap="word")
        self.txt.pack(fill="both", expand=True)

    def pick_files(self):
        paths = filedialog.askopenfilenames(
            title="Выберите Excel-файлы",
            filetypes=[("Excel files", "*.xlsx *.xlsm *.xls"), ("All files", "*.*")]
        )
        if not paths:
            return
        self.files = list(paths)
        if len(self.files) == 1:
            self.lbl_files.config(text="Выбран 1 файл")
        else:
            self.lbl_files.config(text=f"Выбрано файлов: {len(self.files)}")

        self.txt.delete("1.0", tk.END)
        self.txt.insert(tk.END, "Список файлов:\n")
        for p in self.files:
            self.txt.insert(tk.END, f" - {p}\n")

    def clear_files(self):
        self.files = []
        self.lbl_files.config(text="Файлы не выбраны")
        self.txt.delete("1.0", tk.END)

    def copy_result(self):
        text = self.txt.get("1.0", tk.END).strip()
        if not text:
            return
        self.root.clipboard_clear()
        self.root.clipboard_append(text)
        self.root.update()  # keep clipboard after exit
        messagebox.showinfo("Скопировано", "Результат скопирован в буфер обмена.")

    def run(self):
        if not self.files:
            messagebox.showerror("Ошибка", "Сначала выберите файлы.")
            return

        # validate inputs
        try:
            doc_col = int(self.ent_doc_col.get().strip())
            if doc_col <= 0:
                raise ValueError
        except Exception:
            messagebox.showerror("Ошибка", "Колонка с типом документа должна быть целым числом > 0.")
            return

        return_value = self.ent_return_value.get().strip()
        if not return_value:
            messagebox.showerror("Ошибка", "Значение для возврата не может быть пустым.")
            return

        try:
            sum_cols = parse_int_list(self.ent_sum_cols.get())
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))
            return

        needed_cols = []
        for c in [doc_col] + sum_cols:
            if c not in needed_cols:
                needed_cols.append(c)

        total_sums = {c: 0.0 for c in sum_cols}
        per_file_lines = []

        for fpath in self.files:
            try:
                # read only needed columns (0-based for pandas)
                usecols = [c - 1 for c in needed_cols]
                df = pd.read_excel(
                    fpath,
                    engine="openpyxl",
                    usecols=usecols,
                    dtype="string"
                )

                # pandas возвращает колонки в порядке файла (по возрастанию),
                # а не в порядке usecols — поэтому позицию ищем в отсортированном списке
                order = sorted(needed_cols)
                doc_pos = order.index(doc_col)
                doc_series = df.iloc[:, doc_pos].astype("string").fillna("")

                is_return = doc_series.eq(return_value)

                for c in sum_cols:
                    pos = order.index(c)
                    raw = df.iloc[:, pos]
                    num = to_number_series(raw)

                    # make returns negative (force minus of abs)
                    num = np.where(is_return.to_numpy(), -np.abs(num.to_numpy()), num.to_numpy())
                    s = float(np.sum(num))
                    total_sums[c] += s

                per_file_lines.append(f"OK: {fpath}")

            except Exception as e:
                per_file_lines.append(f"ОШИБКА: {fpath}\n  {type(e).__name__}: {e}")

        # output
        self.txt.delete("1.0", tk.END)
        self.txt.insert(tk.END, "\n".join(per_file_lines) + "\n\n")
        self.txt.insert(tk.END, format_results(total_sums, len(self.files)))

def main():
    root = tk.Tk()
    app = WBSummarizerApp(root)
    root.minsize(820, 420)
    root.mainloop()

if __name__ == "__main__":
    main()
