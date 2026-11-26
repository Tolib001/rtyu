"""
oak.uz "Himoya e'lonlari" -> XLSX (F.I.O) for a selected month.

Usage:
    python oak_fio.py              # opens a window to pick year/month
    python oak_fio.py 2026 11      # command line: year month
"""
import sys
import threading
from datetime import date, datetime

import requests
from openpyxl import Workbook
from openpyxl.styles import Font

API = "https://oak.uz/api/elonlar"
LIST_ID = 8  # "Himoya haqidagi e'lonlar"
SITE = "https://oak.uz"

MONTHS = ["Yanvar", "Fevral", "Mart", "Aprel", "May", "Iyun",
          "Iyul", "Avgust", "Sentabr", "Oktabr", "Noyabr", "Dekabr"]

session = requests.Session()
session.headers["User-Agent"] = "Mozilla/5.0"


def fetch_page(page):
    r = session.get(API, params={"list": LIST_ID, "page": page}, timeout=30)
    r.raise_for_status()
    return r.json()


def parse_date(s):
    return datetime.strptime(s, "%d.%m.%Y").date()


def month_key(d):
    return (d.year, d.month)


def find_start_page(target, last_page, log):
    """Binary search for the first page that contains items of `target` month or older.
    Items are sorted newest -> oldest."""
    lo, hi = 1, last_page
    while lo < hi:
        mid = (lo + hi) // 2
        items = fetch_page(mid)["items"]
        log(f"Qidirilmoqda... sahifa {mid}")
        oldest = parse_date(items[-1]["date"]) if items else date.min
        if month_key(oldest) <= target:
            hi = mid
        else:
            lo = mid + 1
    return lo


def collect(year, month, log=print):
    target = (year, month)
    first = fetch_page(1)
    last_page = first["lastPage"]
    page = find_start_page(target, last_page, log)

    rows, seen = [], set()
    while page <= last_page:
        log(f"Yuklanmoqda: sahifa {page}  (topildi: {len(rows)})")
        items = fetch_page(page)["items"]
        if not items:
            break
        stop = False
        for it in items:
            d = parse_date(it["date"])
            if month_key(d) > target:
                continue
            if month_key(d) < target:
                stop = True
                break
            if it["id"] in seen:
                continue
            seen.add(it["id"])
            rows.append((it.get("person", "").strip(), it.get("degree", ""),
                         it["date"], SITE + it.get("href", "")))
        if stop:
            break
        page += 1
    return rows


def save_xlsx(rows, year, month, path=None):
    path = path or f"OAK_FIO_{year}_{month:02d}.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = f"{MONTHS[month - 1]} {year}"
    ws.append(["№", "F.I.O", "Daraja", "Sana", "Havola"])
    for c in ws[1]:
        c.font = Font(bold=True)
    for i, (fio, degree, d, link) in enumerate(rows, 1):
        ws.append([i, fio, degree, d, link])
    for col, w in zip("ABCDE", (6, 45, 8, 12, 45)):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A2"
    wb.save(path)
    return path


def run_gui():
    import tkinter as tk
    from tkinter import ttk, messagebox, filedialog

    root = tk.Tk()
    root.title("OAK himoya e'lonlari -> XLSX")
    root.resizable(False, False)
    today = date.today()

    frm = ttk.Frame(root, padding=16)
    frm.grid()
    ttk.Label(frm, text="Yil:").grid(row=0, column=0, sticky="w")
    year_var = tk.StringVar(value=str(today.year))
    ttk.Spinbox(frm, from_=2010, to=today.year, textvariable=year_var, width=8).grid(row=0, column=1, sticky="w")
    ttk.Label(frm, text="Oy:").grid(row=1, column=0, sticky="w", pady=6)
    month_cb = ttk.Combobox(frm, values=MONTHS, state="readonly", width=12)
    month_cb.current(today.month - 1)
    month_cb.grid(row=1, column=1, sticky="w")
    status = tk.StringVar(value="Oyni tanlang va 'Yaratish' tugmasini bosing")
    ttk.Label(frm, textvariable=status, width=48).grid(row=3, column=0, columnspan=2, pady=(10, 0))
    btn = ttk.Button(frm, text="Yaratish")
    btn.grid(row=2, column=0, columnspan=2, pady=(8, 0))

    def log(msg):
        root.after(0, status.set, msg)

    def work(year, month, path):
        try:
            rows = collect(year, month, log)
            if not rows:
                root.after(0, messagebox.showinfo, "Natija", "Bu oy uchun e'lonlar topilmadi.")
            else:
                save_xlsx(rows, year, month, path)
                root.after(0, messagebox.showinfo, "Tayyor", f"{len(rows)} ta F.I.O saqlandi:\n{path}")
            log("Tayyor")
        except Exception as e:
            root.after(0, messagebox.showerror, "Xato", str(e))
            log("Xato")
        finally:
            root.after(0, lambda: btn.config(state="normal"))

    def start():
        try:
            year = int(year_var.get())
        except ValueError:
            messagebox.showerror("Xato", "Yil noto'g'ri")
            return
        month = month_cb.current() + 1
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")],
            initialfile=f"OAK_FIO_{year}_{month:02d}.xlsx")
        if not path:
            return
        btn.config(state="disabled")
        threading.Thread(target=work, args=(year, month, path), daemon=True).start()

    btn.config(command=start)
    root.mainloop()


if __name__ == "__main__":
    if len(sys.argv) == 3:
        y, m = int(sys.argv[1]), int(sys.argv[2])
        data = collect(y, m)
        print(f"{len(data)} ta yozuv -> {save_xlsx(data, y, m)}")
    else:
        run_gui()
