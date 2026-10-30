# flip_builder.py
# ASCII-only source code.
# Build one B XLSX file per input A file.
#
# Rules:
# - Any number of columns in A is allowed.
# - Uses only:
#     A column 1 -> B column 3
#     A column 3 -> B column 4
# - B column 1 = FIXED1 (counterparty BIN) as text.
# - B column 2 = empty.
# - Skip ONLY first row in A (headers).
# - Output name:
#     one input file:  FLIP.KZ (OPT)_DD.MM.YYYY.xlsx
#     many files:     FLIP.KZ (OPT)_DD.MM.YYYY_1.xlsx, _2.xlsx, ...
#   Real filename uses Cyrillic OPT through Unicode escapes.
# - Output folder: "__OUT" next to source file.
# - Supports XLSX and CSV.

import csv
import re
import threading
from pathlib import Path
from datetime import datetime
from tkinter import Tk, Button, Label, Text, END, filedialog, messagebox, StringVar
from tkinter.ttk import Progressbar

try:
    from openpyxl import Workbook, load_workbook
except Exception as exc:
    raise RuntimeError("Missing dependency: openpyxl") from exc


FIXED1 = "000000000000"  # counterparty BIN (demo value)
FIXED2 = ""
SKIP_ROWS = 1
BASE_NAME = "FLIP.KZ (\u041e\u041f\u0422)"


def now_time():
    return datetime.now().strftime("%H:%M:%S")


def detect_delimiter(sample):
    counts = {
        ";": sample.count(";"),
        ",": sample.count(","),
        "\t": sample.count("\t"),
    }
    return max(counts, key=counts.get)


def read_csv(path):
    last_error = None

    for enc in ["utf-8-sig", "cp1251", "utf-16"]:
        try:
            with open(path, "r", encoding=enc, newline="") as f:
                sample = f.read(5000)

            delimiter = detect_delimiter(sample)

            rows = []
            with open(path, "r", encoding=enc, newline="") as f:
                reader = csv.reader(f, delimiter=delimiter)
                for row in reader:
                    rows.append([str(x) if x is not None else "" for x in row])

            return rows

        except Exception as exc:
            last_error = exc

    raise RuntimeError(f"CSV read error: {last_error}")


def read_xlsx(path):
    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb.worksheets[0]

    rows = []
    for row in ws.iter_rows(values_only=True):
        rows.append([str(x) if x is not None else "" for x in row])

    wb.close()
    return rows


def read_rows(path):
    ext = path.suffix.lower()

    if ext == ".csv":
        return read_csv(path)

    if ext == ".xlsx":
        return read_xlsx(path)

    raise RuntimeError("Input must be XLSX or CSV")


def ensure_unique(path):
    if not path.exists():
        return path

    idx = 1
    while True:
        candidate = path.with_name(f"{path.stem} ({idx}){path.suffix}")
        if not candidate.exists():
            return candidate
        idx += 1


def to_number_or_text(value):
    s = str(value).strip()

    if re.fullmatch(r"\d+", s):
        return int(s)

    return str(value)


def build_b(rows):
    out = []

    for idx in range(SKIP_ROWS, len(rows)):
        row = rows[idx]

        a1 = row[0] if len(row) >= 1 else ""
        a3 = row[2] if len(row) >= 3 else ""

        out.append([FIXED1, FIXED2, a1, a3])

    return out


def save_xlsx(rows, path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    for r_idx, row in enumerate(rows, start=1):
        for c_idx, val in enumerate(row, start=1):
            cell = ws.cell(r_idx, c_idx)

            if c_idx == 1:
                cell.number_format = "@"
                cell.value = str(val)

            elif c_idx == 2:
                cell.value = ""

            else:
                converted = to_number_or_text(val)

                if isinstance(converted, int):
                    cell.number_format = "0"
                    cell.value = converted
                else:
                    cell.number_format = "@"
                    cell.value = str(converted)

    wb.save(path)
    wb.close()


def output_path_for(source_path, total_files, file_index):
    date_str = datetime.now().strftime("%d.%m.%Y")
    out_dir = source_path.parent / "__OUT"
    out_dir.mkdir(exist_ok=True)

    if total_files > 1:
        name = f"{BASE_NAME}_{date_str}_{file_index}.xlsx"
    else:
        name = f"{BASE_NAME}_{date_str}.xlsx"

    return ensure_unique(out_dir / name)


def process_file(file_path, total_files, file_index, log_func):
    p = Path(file_path)

    log_func(f"Reading: {p.name}")

    rows = read_rows(p)

    log_func(f"Rows read: {len(rows)}")
    log_func(f"Skip first rows: {SKIP_ROWS}")

    out_rows = build_b(rows)

    if len(out_rows) == 0:
        raise RuntimeError("No rows to write after skipping header.")

    out_path = output_path_for(p, total_files, file_index)
    save_xlsx(out_rows, out_path)

    log_func(f"Saved: {out_path}")
    return out_path


class App:
    def __init__(self, root):
        self.root = root
        self.root.title("FLIP Builder")
        self.root.geometry("760x520")

        self.status = StringVar()
        self.status.set("Select files or folder")

        Label(root, text="FLIP Builder").pack(pady=8)
        Label(root, textvariable=self.status).pack(pady=4)

        self.btn_files = Button(root, text="Select files", command=self.select_files)
        self.btn_files.pack(pady=4)

        self.btn_folder = Button(root, text="Select folder", command=self.select_folder)
        self.btn_folder.pack(pady=4)

        self.progress = Progressbar(
            root,
            orient="horizontal",
            length=700,
            mode="determinate",
        )
        self.progress.pack(pady=8)

        self.log = Text(root, width=95, height=24)
        self.log.pack(padx=10, pady=10)

    def write_log(self, msg):
        line = f"[{now_time()}] {msg}"
        self.log.insert(END, line + "\n")
        self.log.see(END)
        self.root.update_idletasks()

    def set_progress(self, cur, total):
        if total <= 0:
            self.progress["value"] = 0
        else:
            self.progress["value"] = int(cur * 100 / total)

        self.root.update_idletasks()

    def disable_buttons(self):
        self.btn_files.config(state="disabled")
        self.btn_folder.config(state="disabled")

    def enable_buttons(self):
        self.btn_files.config(state="normal")
        self.btn_folder.config(state="normal")

    def select_files(self):
        files = filedialog.askopenfilenames(
            title="Select A files",
            filetypes=[("Excel/CSV", "*.xlsx *.csv"), ("All files", "*.*")],
        )

        if files:
            self.run(list(files))

    def select_folder(self):
        folder = filedialog.askdirectory(title="Select folder")

        if not folder:
            return

        p = Path(folder)
        files = sorted([str(x) for x in p.glob("*.xlsx")])
        files += sorted([str(x) for x in p.glob("*.csv")])

        if not files:
            messagebox.showwarning("No files", "No XLSX or CSV files found.")
            return

        self.run(files)

    def run(self, files):
        self.disable_buttons()

        self.status.set("Processing...")
        self.progress["value"] = 0

        def worker():
            created = 0
            failed = 0
            total = len(files)

            for idx, file_path in enumerate(files, start=1):
                self.set_progress(idx - 1, total)

                self.write_log("-" * 60)
                self.write_log(f"File {idx} of {total}")

                try:
                    process_file(file_path, total, idx, self.write_log)
                    created += 1

                except Exception as exc:
                    failed += 1
                    self.write_log(f"ERROR: {exc}")

            self.set_progress(total, total)
            self.status.set("Done")

            messagebox.showinfo(
                "Done",
                f"Processed: {total}\n"
                f"Created: {created}\n"
                f"Failed: {failed}\n\n"
                "Results are in __OUT folders.",
            )

            self.enable_buttons()

        threading.Thread(target=worker, daemon=True).start()


def main():
    root = Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
