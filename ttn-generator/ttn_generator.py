#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ТТН Generator - Создание транспортных накладных
"""

import tkinter as tk
from tkinter import ttk, messagebox
import openpyxl
from openpyxl.styles import Font, Alignment
import os
import sys
import re
from datetime import date, timedelta

# ─────────────────────────────────────────────────────────────────────────────
# Шаблон ТТН лежит рядом со скриптом (template_ttn.xlsx).
# При сборке в exe он упаковывается внутрь (см. build_exe.bat).
# ─────────────────────────────────────────────────────────────────────────────

# ─────────────────────────────────────────────────────────────────────────────
# Initial databases (pre-loaded on first run)
# ─────────────────────────────────────────────────────────────────────────────
# Демо-данные. Реальные базы хранятся в Грузополучатели.xlsx и Водители.xlsx
# рядом с программой и пополняются через интерфейс.
INITIAL_RECIPIENTS = [
    (
        "ООО «Маркетплейс Пример», ИНН/КПП: 0000000000 / 000000000, 100000, Московская область, "
        "г. Примерск, ул. Логистическая, д. 1, Р/С: 00000000000000000000, БИК 000000000",
        "Московская область, г. Примерск, ул. Логистическая, д. 1",
        "Склад Север",
    ),
    (
        "ООО «Маркетплейс Пример», ИНН/КПП: 0000000000 / 000000000, 100001, Московская область, "
        "г. Образцово, Промзона, стр. 2, Р/С: 00000000000000000000, БИК 000000000",
        "Московская область, г. Образцово, Промзона, стр. 2",
        "Склад Юг",
    ),
]

INITIAL_DRIVERS = [
    ("Газель", "а000аа00", "Иванов Иван Иванович"),
    ("Форд",   "в111вв11", "Петров Пётр Петрович"),
]


# ─────────────────────────────────────────────────────────────────────────────
# Paths
# ─────────────────────────────────────────────────────────────────────────────
def app_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def template_path():
    # В exe (PyInstaller) файлы из --add-data распаковываются во временную папку sys._MEIPASS
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, "template_ttn.xlsx")


def recipients_path():
    return os.path.join(app_dir(), "Грузополучатели.xlsx")


def drivers_path():
    return os.path.join(app_dir(), "Водители.xlsx")


# ─────────────────────────────────────────────────────────────────────────────
# Database helpers
# ─────────────────────────────────────────────────────────────────────────────
def save_recipients(records):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Грузополучатели"
    ws.append(["Грузополучатель (Реквизиты)", "Адрес доставки", "Краткое наименование"])
    for r in records:
        ws.append([r["full"], r["address"], r["short"]])
    wb.save(recipients_path())
    wb.close()


def load_recipients():
    path = recipients_path()
    if not os.path.exists(path):
        recs = [{"full": f, "address": a, "short": s} for f, a, s in INITIAL_RECIPIENTS]
        save_recipients(recs)
    wb = openpyxl.load_workbook(path)
    ws = wb.active
    result = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if any(v for v in row):
            result.append({
                "full":    str(row[0] or "").strip(),
                "address": str(row[1] or "").strip(),
                "short":   str(row[2] or "").strip(),
            })
    wb.close()
    return sorted(result, key=lambda x: x["short"].lower())


def save_drivers(records):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Водители"
    ws.append(["Марка ТС", "Номер ТС", "ФИО водителя"])
    for r in records:
        ws.append([r["marka"], r["nomer"], r["fio"]])
    wb.save(drivers_path())
    wb.close()


def load_drivers():
    path = drivers_path()
    if not os.path.exists(path):
        drv = [{"marka": m, "nomer": n, "fio": f} for m, n, f in INITIAL_DRIVERS]
        save_drivers(drv)
    wb = openpyxl.load_workbook(path)
    ws = wb.active
    result = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if any(v for v in row):
            result.append({
                "marka": str(row[0] or "").strip(),
                "nomer": str(row[1] or "").strip(),
                "fio":   str(row[2] or "").strip(),
            })
    wb.close()
    return result


# ─────────────────────────────────────────────────────────────────────────────
# Utilities
# ─────────────────────────────────────────────────────────────────────────────
def strip_phone(text):
    """Remove phone numbers from driver FIO string (for BR17)."""
    return re.sub(r'\s*[\+\d][\d\s\-\(\)]{7,}', '', text).strip()


def parse_numbers(text):
    """Extract 13-digit numbers starting with 2000."""
    return re.findall(r'\b2000\d{9}\b', text)


def format_numbers_cl9(nums):
    """Format numbers for CL9: 2 per line, comma-separated."""
    lines = []
    for i in range(0, len(nums), 2):
        lines.append(", ".join(nums[i:i + 2]))
    return "\n".join(lines)


def normalize_date(s):
    """Try to parse a date string and return DD.MM.YYYY, or None."""
    s = s.strip()
    for fmt in ("%d.%m.%Y", "%d.%m.%y", "%d/%m/%Y", "%Y-%m-%d"):
        try:
            from datetime import datetime
            return datetime.strptime(s, fmt).strftime("%d.%m.%Y")
        except ValueError:
            pass
    return None


# ─────────────────────────────────────────────────────────────────────────────
# TTN Generation
# ─────────────────────────────────────────────────────────────────────────────
def generate_ttn(form):
    """
    form keys: date_str, numbers (list), recipient (dict|None),
               places, weight, amount, driver_fio, driver_marka, driver_nomer
    Returns: (output_path, warnings_list)
    """
    warnings = []

    wb = openpyxl.load_workbook(template_path())
    ws1 = wb["стр.1"]
    ws2 = wb["стр.2"]

    # ── BU9: Date ──────────────────────────────────────────────────────────
    date_str = form.get("date_str", "").strip()
    norm = normalize_date(date_str) if date_str else None
    if not norm:
        tomorrow = date.today() + timedelta(days=1)
        norm = tomorrow.strftime("%d.%m.%Y")
        if date_str:
            warnings.append(f"⚠️  Дата '{date_str}' не распознана, использована дата завтра: {norm}")
        else:
            warnings.append(f"⚠️  Дата не указана, использована дата завтра: {norm}")
    ws1["BU9"] = norm
    display_date = norm  # for filename

    # ── CL9: Order numbers ─────────────────────────────────────────────────
    numbers = form.get("numbers", [])
    if not numbers:
        warnings.append("⚠️  Номера поставок не указаны — ячейка CL9 пустая.")
        ws1["CL9"] = ""
    else:
        cell_text = format_numbers_cl9(numbers)
        ws1["CL9"] = cell_text
        ws1["CL9"].font = Font(
            name="Arial",
            size=6,
            bold=False,
        )
        ws1["CL9"].alignment = Alignment(
            wrap_text=True, vertical="top", horizontal="center"
        )
        ws1.row_dimensions[9].height = 100

    # ── B20, B22: Recipient ────────────────────────────────────────────────
    rec = form.get("recipient")
    if not rec:
        warnings.append("⚠️  Грузополучатель не выбран — ячейки B20, B22 пустые.")
        ws1["B20"] = ""
        ws1["B22"] = ""
        ws2["B26"] = ""
    else:
        ws1["B20"] = rec["full"]
        ws1["B22"] = rec["address"]
        ws2["B26"] = rec["address"]

    # ── BF25, B13, BF30: Places ────────────────────────────────────────────
    places = form.get("places", "").strip()
    if not places:
        warnings.append("⚠️  Количество мест не указано — ячейки BF25, B13, BF30 пустые.")
    ws1["BF25"] = places
    ws2["B13"]  = places
    ws2["BF30"] = places

    # ── B27, B11, B32: Weight ──────────────────────────────────────────────
    weight = form.get("weight", "").strip()
    if not weight:
        warnings.append("⚠️  Вес не указан — ячейки B27, B11, B32 пустые.")
    ws1["B27"] = weight
    ws2["B11"] = weight
    ws2["B32"] = weight

    # ── BF29: Amount ───────────────────────────────────────────────────────
    amount = form.get("amount", "").strip()
    if not amount:
        warnings.append("⚠️  Сумма не указана — ячейка BF29 пустая.")
    ws1["BF29"] = amount

    # ── BF44, BR17: Driver FIO ─────────────────────────────────────────────
    fio_full = form.get("driver_fio", "").strip()   # may include phone
    fio_clean = strip_phone(fio_full)               # BR17 — no phone
    if not fio_clean:
        warnings.append("⚠️  ФИО водителя не указано — ячейки BF44, BR17 пустые.")
    ws1["BF44"] = fio_full
    ws2["BR17"] = fio_clean

    # ── B47, BF47: Vehicle ─────────────────────────────────────────────────
    marka = form.get("driver_marka", "").strip()
    nomer = form.get("driver_nomer", "").strip()
    if not marka:
        warnings.append("⚠️  Марка ТС не указана — ячейка B47 пустая.")
    if not nomer:
        warnings.append("⚠️  Номер ТС не указан — ячейка BF47 пустая.")
    ws1["B47"]  = marka
    ws1["BF47"] = nomer

    # ── Build output filename ──────────────────────────────────────────────
    surname = fio_clean.split()[0] if fio_clean else "ТТН"
    short   = rec["short"] if rec else "Грузополучатель"
    date_fs = display_date.replace(".", "_")
    out_name = f"{surname}_{short}_{date_fs}.xlsx"
    out_name = re.sub(r'[<>:"/\\|?*]', "_", out_name)
    out_path = os.path.join(app_dir(), out_name)

    wb.save(out_path)
    wb.close()
    return out_path, warnings


# ─────────────────────────────────────────────────────────────────────────────
# Dialog: Add Recipient
# ─────────────────────────────────────────────────────────────────────────────
class AddRecipientDialog(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Новый грузополучатель")
        self.configure(bg=BG)
        self.resizable(False, False)
        self.result = None
        self._build()
        self.transient(parent)
        self.grab_set()
        self.wait_window()

    def _build(self):
        p = dict(padx=12, pady=4)
        tk.Label(self, text="Краткое наименование  *", bg=BG, font=LBL).grid(row=0, column=0, sticky="w", **p)
        self.e_short = tk.Entry(self, width=42, font=ENT, relief="solid", bd=1)
        self.e_short.grid(row=0, column=1, **p); bind_paste(self.e_short)

        tk.Label(self, text="Грузополучатель (реквизиты)  *", bg=BG, font=LBL).grid(row=1, column=0, sticky="nw", **p)
        self.e_full = tk.Text(self, width=42, height=6, font=ENT, relief="solid", bd=1)
        self.e_full.grid(row=1, column=1, **p); bind_paste(self.e_full, multiline=True)

        tk.Label(self, text="Адрес доставки  *", bg=BG, font=LBL).grid(row=2, column=0, sticky="nw", **p)
        self.e_addr = tk.Text(self, width=42, height=3, font=ENT, relief="solid", bd=1)
        self.e_addr.grid(row=2, column=1, **p); bind_paste(self.e_addr, multiline=True)

        bf = tk.Frame(self, bg=BG)
        bf.grid(row=3, column=0, columnspan=2, pady=10)
        tk.Button(bf, text="Сохранить", command=self._save, bg="#4a86c8", fg="white",
                  font=LBL, width=12, relief="flat").pack(side="left", padx=6)
        tk.Button(bf, text="Отмена", command=self.destroy, bg="#c8c8c8",
                  font=LBL, width=12, relief="flat").pack(side="left", padx=6)

    def _save(self):
        short = self.e_short.get().strip()
        full  = self.e_full.get("1.0", "end").strip()
        addr  = self.e_addr.get("1.0", "end").strip()
        if not short or not full or not addr:
            messagebox.showerror("Ошибка", "Заполните все обязательные поля (*)", parent=self)
            return
        self.result = {"short": short, "full": full, "address": addr}
        self.destroy()


# ─────────────────────────────────────────────────────────────────────────────
# Dialog: Add Driver
# ─────────────────────────────────────────────────────────────────────────────
class AddDriverDialog(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Новый водитель")
        self.configure(bg=BG)
        self.resizable(False, False)
        self.result = None
        self._build()
        self.transient(parent)
        self.grab_set()
        self.wait_window()

    def _build(self):
        p = dict(padx=12, pady=5)
        tk.Label(self, text="ФИО водителя  *", bg=BG, font=LBL).grid(row=0, column=0, sticky="w", **p)
        self.e_fio = tk.Entry(self, width=36, font=ENT, relief="solid", bd=1)
        self.e_fio.grid(row=0, column=1, **p); bind_paste(self.e_fio)

        tk.Label(self, text="Марка ТС  *", bg=BG, font=LBL).grid(row=1, column=0, sticky="w", **p)
        self.e_marka = tk.Entry(self, width=36, font=ENT, relief="solid", bd=1)
        self.e_marka.grid(row=1, column=1, **p); bind_paste(self.e_marka)

        tk.Label(self, text="Номер ТС  *", bg=BG, font=LBL).grid(row=2, column=0, sticky="w", **p)
        self.e_nomer = tk.Entry(self, width=36, font=ENT, relief="solid", bd=1)
        self.e_nomer.grid(row=2, column=1, **p); bind_paste(self.e_nomer)

        bf = tk.Frame(self, bg=BG)
        bf.grid(row=3, column=0, columnspan=2, pady=10)
        tk.Button(bf, text="Сохранить", command=self._save, bg="#4a86c8", fg="white",
                  font=LBL, width=12, relief="flat").pack(side="left", padx=6)
        tk.Button(bf, text="Отмена", command=self.destroy, bg="#c8c8c8",
                  font=LBL, width=12, relief="flat").pack(side="left", padx=6)

    def _save(self):
        fio   = self.e_fio.get().strip()
        marka = self.e_marka.get().strip()
        nomer = self.e_nomer.get().strip()
        if not fio or not marka or not nomer:
            messagebox.showerror("Ошибка", "Заполните все обязательные поля (*)", parent=self)
            return
        self.result = {"fio": fio, "marka": marka, "nomer": nomer}
        self.destroy()


# ─────────────────────────────────────────────────────────────────────────────
# Style constants
# ─────────────────────────────────────────────────────────────────────────────
BG  = "#f0f0f0"
BG2 = "#e4e4e4"
LBL = ("Arial", 9)
ENT = ("Arial", 10)
TTL = ("Arial", 13, "bold")
SEC = ("Arial", 10, "bold")

# ─────────────────────────────────────────────────────────────────────────────
# Paste helper (ensures Ctrl+V works on all platforms / all widget types)
# ─────────────────────────────────────────────────────────────────────────────
def bind_paste(widget, multiline=False):
    """Explicitly bind Ctrl+V/Ctrl+Shift+V paste to any Entry or Text widget."""
    def _paste(event):
        try:
            text = widget.clipboard_get()
            if not multiline:
                # Entry: strip newlines so single-line fields stay clean
                text = text.replace("\r\n", " ").replace("\n", " ").replace("\r", " ").strip()
            if isinstance(widget, tk.Text):
                try:
                    widget.delete(tk.SEL_FIRST, tk.SEL_LAST)
                except tk.TclError:
                    pass
                widget.insert(tk.INSERT, text)
            else:
                try:
                    sel_start = widget.index(tk.SEL_FIRST)
                    sel_end   = widget.index(tk.SEL_LAST)
                    widget.delete(sel_start, sel_end)
                except tk.TclError:
                    pass
                widget.insert(tk.INSERT, text)
        except tk.TclError:
            pass
        return "break"

    widget.bind("<Control-v>",       _paste)
    widget.bind("<Control-V>",       _paste)
    widget.bind("<Control-KeyPress-v>", _paste)



# ─────────────────────────────────────────────────────────────────────────────
# Main window
# ─────────────────────────────────────────────────────────────────────────────
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("ТТН – Создание накладных")
        self.configure(bg=BG)
        self.resizable(False, False)

        self.recipients  = []
        self.drivers     = []
        self._sel_rec    = None   # selected recipient dict
        self._all_driver_fios = []  # full list for filter restore

        self._load_data()
        self._build_ui()
        self._bind_paste()

    # ── Data ──────────────────────────────────────────────────────────────────
    def _load_data(self):
        self.recipients = load_recipients()
        self.drivers    = load_drivers()
        self._all_driver_fios = [d["fio"] for d in self.drivers]

    def _rec_short_list(self):
        return [r["short"] for r in self.recipients]

    # ── UI Build ──────────────────────────────────────────────────────────────
    def _build_ui(self):
        # ── Title ──────────────────────────────────────────────────────────
        tk.Label(self, text="Создание Транспортной Накладной",
                 font=TTL, bg=BG).grid(row=0, column=0, columnspan=3,
                                       pady=(16, 8), padx=20)

        sep = ttk.Separator(self, orient="horizontal")
        sep.grid(row=1, column=0, columnspan=3, sticky="ew", padx=16)

        # ── Left panel ─────────────────────────────────────────────────────
        lp = tk.Frame(self, bg=BG)
        lp.grid(row=2, column=0, sticky="nsew", padx=(20, 8), pady=12)

        rp = tk.Frame(self, bg=BG)
        rp.grid(row=2, column=1, sticky="nsew", padx=(8, 20), pady=12)

        vsep = ttk.Separator(self, orient="vertical")
        vsep.grid(row=2, column=2, sticky="ns", padx=0)

        # swap: left=left, right=right
        self._build_left(lp)
        self._build_right(rp)

        sep2 = ttk.Separator(self, orient="horizontal")
        sep2.grid(row=3, column=0, columnspan=3, sticky="ew", padx=16)

        # ── Generate button ────────────────────────────────────────────────
        bf = tk.Frame(self, bg=BG)
        bf.grid(row=4, column=0, columnspan=3, pady=12)
        self.btn_gen = tk.Button(
            bf, text="  ✅  СОЗДАТЬ ТТН  ",
            command=self._on_generate,
            font=("Arial", 12, "bold"),
            bg="#4a86c8", fg="white",
            relief="flat", padx=24, pady=10,
            cursor="hand2",
        )
        self.btn_gen.pack()

        sep3 = ttk.Separator(self, orient="horizontal")
        sep3.grid(row=5, column=0, columnspan=3, sticky="ew", padx=16)

        # ── Log area ───────────────────────────────────────────────────────
        log_frame = tk.LabelFrame(self, text="  Сообщения  ", bg=BG,
                                   font=LBL, labelanchor="nw")
        log_frame.grid(row=6, column=0, columnspan=3,
                       sticky="ew", padx=16, pady=(8, 16))
        self.log = tk.Text(
            log_frame, height=6, font=("Courier", 9),
            bg="white", relief="flat", state="disabled",
            wrap="word", bd=1,
        )
        sb = ttk.Scrollbar(log_frame, command=self.log.yview)
        self.log.configure(yscrollcommand=sb.set)
        self.log.pack(side="left", fill="both", expand=True, padx=4, pady=4)
        sb.pack(side="right", fill="y", pady=4)

    def _section(self, parent, text, row):
        tk.Label(parent, text=text, font=SEC, bg=BG2,
                 anchor="w", width=35).grid(
            row=row, column=0, columnspan=2,
            sticky="ew", pady=(12, 2), ipady=3, padx=2)

    def _lbl(self, parent, text, row):
        tk.Label(parent, text=text, font=LBL, bg=BG, anchor="w").grid(
            row=row, column=0, sticky="w", pady=(4, 0))

    # ── Left panel ────────────────────────────────────────────────────────
    def _build_left(self, p):
        r = 0

        # ─ Delivery ─
        self._section(p, "  Доставка", r); r += 1

        self._lbl(p, "Дата  *", r); r += 1
        tomorrow = (date.today() + timedelta(days=1)).strftime("%d.%m.%Y")
        self.date_var = tk.StringVar(value=tomorrow)
        _e = tk.Entry(p, textvariable=self.date_var, font=ENT, width=16,
                 relief="solid", bd=1)
        _e.grid(row=r, column=0, sticky="w", pady=(0, 4)); bind_paste(_e); r += 1

        self._lbl(p, "Грузополучатель  *", r); r += 1
        self.rec_var = tk.StringVar()
        self._rec_cb = ttk.Combobox(p, textvariable=self.rec_var,
                              values=self._rec_short_list(),
                              width=30, state="readonly")
        self._rec_cb.grid(row=r, column=0, sticky="w", pady=(0, 2)); r += 1
        self._rec_cb.bind("<<ComboboxSelected>>", self._on_rec_selected)
        tk.Button(p, text="+ Добавить нового грузополучателя",
                  command=self._add_recipient, font=("Arial", 8),
                  bg=BG2, relief="flat", cursor="hand2").grid(
            row=r, column=0, sticky="w", pady=(0, 4)); r += 1

        # ─ Cargo ─
        self._section(p, "  Груз", r); r += 1

        self._lbl(p, "Количество мест  *  (8 коробов / 2 паллет)", r); r += 1
        self.places_var = tk.StringVar()
        _e = tk.Entry(p, textvariable=self.places_var, font=ENT, width=22,
                 relief="solid", bd=1)
        _e.grid(row=r, column=0, sticky="w", pady=(0, 4)); bind_paste(_e); r += 1

        self._lbl(p, "Вес  *  (96 кг)", r); r += 1
        self.weight_var = tk.StringVar()
        _e = tk.Entry(p, textvariable=self.weight_var, font=ENT, width=22,
                 relief="solid", bd=1)
        _e.grid(row=r, column=0, sticky="w", pady=(0, 4)); bind_paste(_e); r += 1

        self._lbl(p, "Сумма  *  (176720 руб.)", r); r += 1
        self.amount_var = tk.StringVar()
        _e = tk.Entry(p, textvariable=self.amount_var, font=ENT, width=22,
                 relief="solid", bd=1)
        _e.grid(row=r, column=0, sticky="w", pady=(0, 4)); bind_paste(_e); r += 1

    # ── Right panel ───────────────────────────────────────────────────────
    def _build_right(self, p):
        r = 0

        # ─ Order numbers ─
        self._section(p, "  Номера поставок", r); r += 1

        tk.Label(p, text="Введите номера (через пробел или по одному на строке)  *",
                 font=("Arial", 8), bg=BG, fg="#555", anchor="w").grid(
            row=r, column=0, columnspan=2, sticky="w"); r += 1
        self.numbers_text = tk.Text(p, font=ENT, width=34, height=6,
                                     relief="solid", bd=1)
        self.numbers_text.grid(row=r, column=0, columnspan=2,
                                sticky="ew", pady=(0, 4)); bind_paste(self.numbers_text, multiline=True); r += 1

        # ─ Driver ─
        self._section(p, "  Водитель", r); r += 1

        self._lbl(p, "Поиск по фамилии:", r); r += 1
        self.drv_search_var = tk.StringVar()
        self.drv_search_var.trace("w", self._filter_drivers)
        _e = tk.Entry(p, textvariable=self.drv_search_var, font=ENT, width=34,
                 relief="solid", bd=1)
        _e.grid(row=r, column=0, columnspan=2,
                sticky="ew", pady=(0, 2)); bind_paste(_e); r += 1

        self.drv_var = tk.StringVar()
        self.drv_cb = ttk.Combobox(p, textvariable=self.drv_var,
                                    values=["+ Добавить нового водителя"] + self._all_driver_fios,
                                    width=32, state="readonly")
        self.drv_cb.grid(row=r, column=0, columnspan=2,
                          sticky="ew", pady=(0, 6)); r += 1
        self.drv_cb.bind("<<ComboboxSelected>>", self._on_drv_selected)

        # Detail fields
        detail = tk.Frame(p, bg=BG)
        detail.grid(row=r, column=0, columnspan=2, sticky="ew"); r += 1

        self._lbl(detail, "ФИО водителя  *  (можно добавить телефон)", 0)
        detail.grid_rowconfigure(0, pad=2)
        self.fio_var = tk.StringVar()
        _e = tk.Entry(detail, textvariable=self.fio_var, font=ENT, width=34,
                 relief="solid", bd=1)
        _e.grid(row=1, column=0, sticky="w", pady=(0, 4)); bind_paste(_e)

        tk.Label(detail, text="Марка ТС  *", font=LBL, bg=BG, anchor="w").grid(
            row=2, column=0, sticky="w", pady=(4, 0))
        self.marka_var = tk.StringVar()
        _e = tk.Entry(detail, textvariable=self.marka_var, font=ENT, width=20,
                 relief="solid", bd=1)
        _e.grid(row=3, column=0, sticky="w", pady=(0, 4)); bind_paste(_e)

        tk.Label(detail, text="Номер ТС  *", font=LBL, bg=BG, anchor="w").grid(
            row=4, column=0, sticky="w", pady=(4, 0))
        self.nomer_var = tk.StringVar()
        _e = tk.Entry(detail, textvariable=self.nomer_var, font=ENT, width=20,
                 relief="solid", bd=1)
        _e.grid(row=5, column=0, sticky="w", pady=(0, 4)); bind_paste(_e)

    # ── Events ────────────────────────────────────────────────────────────────
    def _on_rec_selected(self, _=None):
        val = self.rec_var.get()
        self._sel_rec = next((r for r in self.recipients if r["short"] == val), None)

    def _filter_drivers(self, *_):
        term = self.drv_search_var.get().lower().strip()
        if term:
            filtered = [d["fio"] for d in self.drivers if term in d["fio"].lower()]
        else:
            filtered = self._all_driver_fios
        self.drv_cb["values"] = ["+ Добавить нового водителя"] + filtered

    def _on_drv_selected(self, _=None):
        val = self.drv_var.get()
        if val == "+ Добавить нового водителя":
            self._add_driver()
            return
        drv = next((d for d in self.drivers if d["fio"] == val), None)
        if drv:
            self.fio_var.set(drv["fio"])
            self.marka_var.set(drv["marka"])
            self.nomer_var.set(drv["nomer"])

    def _add_recipient(self):
        dlg = AddRecipientDialog(self)
        if dlg.result:
            self.recipients.append(dlg.result)
            self.recipients = sorted(self.recipients, key=lambda x: x["short"].lower())
            save_recipients(self.recipients)
            self.rec_var.set(dlg.result["short"])
            self._sel_rec = dlg.result
            # Update combobox
            cb = [w for w in self.winfo_children() if isinstance(w, tk.Frame)]
            self._refresh_rec_combo()
            self._rec_cb['values'] = self._rec_short_list()
            self._log(f"✅  Грузополучатель «{dlg.result['short']}» добавлен в базу.")

    def _refresh_rec_combo(self):
        pass  # handled via self._rec_cb

    def _add_driver(self):
        dlg = AddDriverDialog(self)
        if dlg.result:
            self.drivers.append(dlg.result)
            self._all_driver_fios = [d["fio"] for d in self.drivers]
            save_drivers(self.drivers)
            self.drv_cb["values"] = ["+ Добавить нового водителя"] + self._all_driver_fios
            self.drv_var.set(dlg.result["fio"])
            self.fio_var.set(dlg.result["fio"])
            self.marka_var.set(dlg.result["marka"])
            self.nomer_var.set(dlg.result["nomer"])
            self._log(f"✅  Водитель «{dlg.result['fio']}» добавлен в базу.")
        else:
            self.drv_var.set("")

    # ── Log ───────────────────────────────────────────────────────────────────
    def _log(self, msg):
        self.log.configure(state="normal")
        self.log.insert("end", msg + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def _clear_log(self):
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.configure(state="disabled")

    # ── Validate + Normalize ──────────────────────────────────────────────────
    def _prepare_form(self):
        """Returns (form_dict, warnings_list). Form is always returned."""
        warnings = []

        numbers_raw = self.numbers_text.get("1.0", "end").strip()
        numbers = parse_numbers(numbers_raw)
        if numbers_raw and not numbers:
            warnings.append("⚠️  Номера поставок не найдены. Ожидаются 13-значные числа, начинающиеся с 2000.")
        elif not numbers_raw:
            warnings.append("⚠️  Номера поставок не введены.")

        # Places – validate format
        places = self.places_var.get().strip()
        PLACE_WORDS = ["коробок", "коробов", "паллет", "пал", "кор", " к"]
        if places:
            has_num  = bool(re.search(r'\d', places))
            has_word = any(w in places.lower() for w in PLACE_WORDS)
            if not has_num or not has_word:
                warnings.append(
                    "⚠️  Формат поля «Количество мест» не распознан.\n"
                    "    Ожидается: число + тип (коробок / паллет / пал / кор / к). "
                    "    Значение будет вставлено как есть."
                )
        else:
            warnings.append("⚠️  Количество мест не указано.")

        # Weight – normalize
        weight = self.weight_var.get().strip()
        if weight:
            if re.match(r'^[\d\s,\.]+$', weight):
                weight = weight.strip() + " кг"
            elif not re.search(r'\d', weight):
                warnings.append("⚠️  Поле «Вес» не содержит числа. Вставлено как есть.")
        else:
            warnings.append("⚠️  Вес не указан.")

        # Amount – normalize
        amount = self.amount_var.get().strip()
        if amount:
            if re.match(r'^[\d\s,\.]+$', amount):
                amount = amount.strip() + " руб."
            elif not re.search(r'\d', amount):
                warnings.append("⚠️  Поле «Сумма» не содержит числа. Вставлено как есть.")
        else:
            warnings.append("⚠️  Сумма не указана.")

        # Recipient
        if not self._sel_rec:
            warnings.append("⚠️  Грузополучатель не выбран.")

        # Driver
        fio   = self.fio_var.get().strip()
        marka = self.marka_var.get().strip()
        nomer = self.nomer_var.get().strip()
        if not fio:   warnings.append("⚠️  ФИО водителя не указано.")
        if not marka: warnings.append("⚠️  Марка ТС не указана.")
        if not nomer: warnings.append("⚠️  Номер ТС не указан.")

        form = {
            "date_str":     self.date_var.get().strip(),
            "numbers":      numbers,
            "recipient":    self._sel_rec,
            "places":       places,
            "weight":       weight,
            "amount":       amount,
            "driver_fio":   fio,
            "driver_marka": marka,
            "driver_nomer": nomer,
        }
        return form, warnings


    def _bind_paste(self):
        """Enable Ctrl+V paste for all editable Entry and Text widgets.
        Works with both English and Russian keyboard layouts by checking
        the physical keycode (86 = V key on any layout on Windows)."""
        def _do_paste(w):
            try:
                clip = self.clipboard_get()
                if isinstance(w, tk.Text):
                    try:
                        w.delete(tk.SEL_FIRST, tk.SEL_LAST)
                    except tk.TclError:
                        pass
                    w.insert(tk.INSERT, clip)
                else:
                    try:
                        w.delete(tk.SEL_FIRST, tk.SEL_LAST)
                    except tk.TclError:
                        pass
                    w.insert(tk.INSERT, clip)
            except tk.TclError:
                pass

        def _add(widget):
            if isinstance(widget, (tk.Entry, tk.Text)):
                try:
                    state = str(widget.cget('state'))
                except Exception:
                    state = 'normal'
                if state not in ('readonly', 'disabled'):
                    def _handler(event, w=widget):
                        # keycode 86 = physical V key on Windows,
                        # works regardless of EN/RU keyboard layout
                        if event.keycode == 86:
                            _do_paste(w)
                            return "break"
                    # keycode-based: catches both EN (v) and RU (м)
                    widget.bind('<Control-KeyPress>', _handler)
                    # also keep char bindings as fallback for other OS
                    def _paste_char(event, w=widget):
                        _do_paste(w)
                        return "break"
                    widget.bind('<Control-v>', _paste_char)
                    widget.bind('<Control-V>', _paste_char)
            for child in widget.winfo_children():
                _add(child)
        _add(self)

    # ── Generate ──────────────────────────────────────────────────────────────
    def _on_generate(self):
        self._clear_log()
        form, warnings = self._prepare_form()

        if warnings:
            self._log("─── Предупреждения ──────────────────")
            for w in warnings:
                self._log(w)
            self._log("─────────────────────────────────────")

        self.btn_gen.configure(state="disabled", text="  ⏳  Создаётся...")
        self.update_idletasks()

        try:
            out_path, gen_warn = generate_ttn(form)
            for w in gen_warn:
                self._log(w)
            self._log(f"✅  Файл создан: {os.path.basename(out_path)}")
            self._log(f"📁  Папка: {os.path.dirname(out_path)}")
        except Exception as exc:
            self._log(f"❌  Ошибка создания файла: {exc}")
            import traceback
            traceback.print_exc()
        finally:
            self.btn_gen.configure(state="normal", text="  ✅  СОЗДАТЬ ТТН  ")


# ─────────────────────────────────────────────────────────────────────────────
# Entry point
# ─────────────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    app = App()
    app.mainloop()
