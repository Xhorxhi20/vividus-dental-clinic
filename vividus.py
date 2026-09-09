import customtkinter as ctk
from tkinter import messagebox
import sqlite3
import os
import shutil
import re
from datetime import datetime
import openpyxl


# ================= BACKUP DATABASE =================

# Databaza ruhet gjithmonë në Dokumentet e përdoruesit, jo në folderin e programit
import sys


def get_app_folder():
    """Gjen një vend të sigurt për databazën, që punon në çdo kompjuter."""
    candidates = [
        os.environ.get("LOCALAPPDATA"),
        os.path.join(os.path.expanduser("~"), "AppData", "Local"),
        os.path.dirname(sys.executable if getattr(sys, "frozen", False) else __file__),
    ]

    for base in candidates:
        if not base:
            continue
        folder = os.path.join(base, "VividusClinic")
        try:
            os.makedirs(folder, exist_ok=True)
            # provë reale shkrimi
            test = os.path.join(folder, ".test")
            with open(test, "w") as f:
                f.write("ok")
            os.remove(test)
            return folder
        except Exception:
            continue

    return os.getcwd()


APP_FOLDER = get_app_folder()
DB_FILE = os.path.join(APP_FOLDER, "vividus_pro.db")
BACKUP_DIR = os.path.join(APP_FOLDER, "backups")

def backup_database():
    if not os.path.exists(DB_FILE):
        return

    if not os.path.exists(BACKUP_DIR):
        os.makedirs(BACKUP_DIR)

    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    shutil.copy(DB_FILE, os.path.join(BACKUP_DIR, f"vividus_backup_{stamp}.db"))

    backups = sorted(f for f in os.listdir(BACKUP_DIR) if f.startswith("vividus_backup_"))
    while len(backups) > 10:
        os.remove(os.path.join(BACKUP_DIR, backups.pop(0)))


backup_database()


# ================= DATABASE =================

conn = sqlite3.connect(DB_FILE)
cursor = conn.cursor()

cursor.execute("""
CREATE TABLE IF NOT EXISTS patients(
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT,
    surname TEXT,
    phone TEXT,
    email TEXT,
    address TEXT,
    notes TEXT
)
""")

cursor.execute("""
CREATE TABLE IF NOT EXISTS appointments(
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    patient_id INTEGER,
    date TEXT,
    time TEXT,
    treatment TEXT,
    price REAL,
    status TEXT,
    notes TEXT
)
""")

conn.commit()


# ================= VALIDATION =================

def is_valid_date(value):
    try:
        datetime.strptime(value.strip(), "%d/%m/%Y")
        return True
    except ValueError:
        return False


def is_valid_time(value):
    return bool(re.match(r"^([01]\d|2[0-3]):[0-5]\d$", value.strip()))


def is_valid_price(value):
    try:
        return float(value.replace(",", ".")) >= 0
    except ValueError:
        return False


# ================= PALETTE =================

ctk.set_appearance_mode("light")

CREAM = "#FBF6EE"
CARD = "#F3E7D6"
CARD_HOVER = "#F7EFE3"
BROWN = "#6F4E37"
DARK_BROWN = "#3B2A20"
GOLD = "#C9A227"
GOLD_LIGHT = "#E4C767"
SUCCESS = "#7C8C4B"
DANGER = "#8B3A3A"
NEUTRAL = "#9C8776"
TEXT_DARK = "#2E2118"
TEXT_MUTED = "#8A7767"
LINE = "#E8DCCB"

STATUS_COLORS = {
    "Planifikuar": GOLD,
    "Përfunduar": SUCCESS,
    "Anulluar": DANGER,
}


# ================= MAIN WINDOW =================

app = ctk.CTk()
app.title("Vividus Dental Clinic")
app.geometry("1340x820")
app.configure(fg_color=CREAM)


# ================= HELPERS =================

def bind_recursive(widget, sequence, callback):
    """Bind an event to a widget and all of its children."""
    widget.bind(sequence, callback)
    for child in widget.winfo_children():
        bind_recursive(child, sequence, callback)


def initials(name, surname):
    a = name.strip()[:1].upper() if name.strip() else "?"
    b = surname.strip()[:1].upper() if surname.strip() else ""
    return a + b


def pill(parent, text, color):
    """Small rounded colored badge."""
    p = ctk.CTkFrame(parent, fg_color=color, corner_radius=12, height=26)
    ctk.CTkLabel(
        p, text=text, font=("Segoe UI", 11, "bold"),
        text_color="white", fg_color="transparent",
    ).pack(padx=14, pady=3)
    return p


def avatar(parent, text, color=BROWN, size=44):
    a = ctk.CTkFrame(parent, width=size, height=size, corner_radius=size // 2, fg_color=color)
    a.pack_propagate(False)
    ctk.CTkLabel(
        a, text=text, font=("Segoe UI", 15, "bold"),
        text_color="white", fg_color="transparent",
    ).place(relx=0.5, rely=0.5, anchor="center")
    return a


# ================= HEADER =================

header = ctk.CTkFrame(app, height=92, fg_color=DARK_BROWN, corner_radius=0)
header.pack(fill="x")

ctk.CTkLabel(
    header, text="🦷  VIVIDUS  DENTAL  CLINIC",
    font=("Georgia", 30, "bold"), text_color=GOLD_LIGHT,
).pack(pady=(18, 0))

ctk.CTkLabel(
    header, text="Kujdes dentar me përsosmëri",
    font=("Segoe UI", 12, "italic"), text_color="#D9C9A8",
).pack(pady=(0, 12))

gold_line = ctk.CTkFrame(app, height=3, fg_color=GOLD, corner_radius=0)
gold_line.pack(fill="x")


# ================= LAYOUT =================

main = ctk.CTkFrame(app, fg_color=CREAM)
main.pack(fill="both", expand=True)

menu = ctk.CTkFrame(main, width=230, fg_color=DARK_BROWN, corner_radius=0)
menu.pack(side="left", fill="y")
menu.pack_propagate(False)

content = ctk.CTkFrame(main, fg_color="white", corner_radius=16)
content.pack(side="right", fill="both", expand=True, padx=14, pady=14)


def clear_content():
    for widget in content.winfo_children():
        widget.destroy()


ctk.CTkLabel(menu, text="✦", font=("Georgia", 38), text_color=GOLD).pack(pady=(28, 6))
ctk.CTkLabel(menu, text="M E N U", font=("Segoe UI", 11, "bold"), text_color="#8C7660").pack(pady=(0, 14))


def menu_button(text, command):
    btn = ctk.CTkButton(
        menu, text=text, width=196, height=46, anchor="w",
        font=("Segoe UI", 14, "bold"), corner_radius=10,
        fg_color=BROWN, hover_color=GOLD, text_color="white",
        command=command,
    )
    btn.pack(pady=8, padx=17)
    return btn


def section_title(text, subtitle=None):
    wrap = ctk.CTkFrame(content, fg_color="white")
    wrap.pack(fill="x", padx=28, pady=(20, 6))
    ctk.CTkLabel(wrap, text=text, font=("Georgia", 25, "bold"), text_color=DARK_BROWN).pack(anchor="w")
    if subtitle:
        ctk.CTkLabel(wrap, text=subtitle, font=("Segoe UI", 12), text_color=TEXT_MUTED).pack(anchor="w", pady=(2, 0))
    ctk.CTkFrame(wrap, height=2, fg_color=GOLD_LIGHT).pack(fill="x", pady=(10, 0))
    return wrap


def empty_state(parent, icon, text):
    box = ctk.CTkFrame(parent, fg_color="white")
    box.pack(pady=60)
    ctk.CTkLabel(box, text=icon, font=("Segoe UI", 42), text_color=LINE).pack()
    ctk.CTkLabel(box, text=text, font=("Segoe UI", 14), text_color=TEXT_MUTED).pack(pady=8)


# ================= PATIENTS PAGE =================

def patients_page():
    clear_content()
    state = {"editing": None, "selected": None, "cards": {}}

    section_title("👤  Pacientët", "Regjistro, kërko dhe menaxho kartelat e pacientëve")

    # ---------- FORM ----------
    form = ctk.CTkFrame(content, fg_color=CARD, corner_radius=14, border_width=1, border_color=GOLD_LIGHT)
    form.pack(fill="x", padx=28, pady=(10, 6))

    entries = {}
    fields = [
        ("name", "Emri", 0, 0), ("surname", "Mbiemri", 0, 1),
        ("phone", "Telefon", 0, 2), ("email", "Email", 0, 3),
        ("address", "Adresa", 1, 0), ("notes", "Shënime", 1, 1),
    ]
    for key, label, r, c in fields:
        e = ctk.CTkEntry(
            form, placeholder_text=label, height=38, corner_radius=9,
            width=230 if key != "notes" else 480,
            border_color=LINE, fg_color="white",
        )
        e.grid(row=r, column=c, padx=10, pady=12,
               columnspan=2 if key == "notes" else 1, sticky="w")
        entries[key] = e

    def clear_form():
        for e in entries.values():
            e.delete(0, "end")
        state["editing"] = None
        save_btn.configure(text="💾   Ruaj Pacient", fg_color=SUCCESS)

    # ---------- SEARCH BAR ----------
    bar = ctk.CTkFrame(content, fg_color="white")
    bar.pack(fill="x", padx=28, pady=(8, 0))

    search_entry = ctk.CTkEntry(
        bar, placeholder_text="🔍   Kërko pacient (emër, mbiemër, telefon)...",
        width=420, height=38, corner_radius=19, border_color=LINE, fg_color=CREAM,
    )
    search_entry.pack(side="left")

    count_label = ctk.CTkLabel(bar, text="", font=("Segoe UI", 12), text_color=TEXT_MUTED)
    count_label.pack(side="right", padx=6)

    # ---------- LIST HEADER ----------
    head = ctk.CTkFrame(content, fg_color=DARK_BROWN, corner_radius=10, height=38)
    head.pack(fill="x", padx=28, pady=(12, 4))
    head.pack_propagate(False)

    def head_col(text, width, anchor="w"):
        ctk.CTkLabel(
            head, text=text, width=width, anchor=anchor,
            font=("Segoe UI", 11, "bold"), text_color=GOLD_LIGHT,
        ).pack(side="left", padx=(14, 0))

    head_col("PACIENTI", 260)
    head_col("TELEFON", 150)
    head_col("EMAIL", 210)
    head_col("ADRESA", 180)
    head_col("SHËNIME", 160)

    # ---------- SCROLLABLE CARD LIST ----------
    listbox = ctk.CTkScrollableFrame(content, fg_color="white", corner_radius=0)
    listbox.pack(fill="both", expand=True, padx=22, pady=(0, 4))

    def select_card(pid):
        state["selected"] = pid
        for other_id, card in state["cards"].items():
            if other_id == pid:
                card.configure(fg_color=CARD, border_color=GOLD, border_width=2)
            else:
                card.configure(fg_color="white", border_color=LINE, border_width=1)

    def build_card(row):
        pid, name, surname, phone, email, address, notes = row

        card = ctk.CTkFrame(
            listbox, fg_color="white", corner_radius=12,
            border_width=1, border_color=LINE, height=66,
        )
        card.pack(fill="x", pady=4, padx=4)
        card.pack_propagate(False)

        left = ctk.CTkFrame(card, fg_color="transparent")
        left.pack(side="left", fill="y", padx=(12, 0))

        avatar(left, initials(name, surname)).pack(side="left", pady=11)

        info = ctk.CTkFrame(left, fg_color="transparent")
        info.pack(side="left", padx=12)
        ctk.CTkLabel(
            info, text=f"{name} {surname}", font=("Segoe UI", 14, "bold"),
            text_color=TEXT_DARK, anchor="w", width=190,
        ).pack(anchor="w", pady=(12, 0))
        ctk.CTkLabel(
            info, text=f"ID #{pid}", font=("Segoe UI", 10),
            text_color=TEXT_MUTED, anchor="w", width=190,
        ).pack(anchor="w")

        def col(text, width, muted=False):
            ctk.CTkLabel(
                card, text=text if text else "—", width=width, anchor="w",
                font=("Segoe UI", 12), text_color=TEXT_MUTED if muted else TEXT_DARK,
            ).pack(side="left", padx=(14, 0))

        col(phone, 150)
        col(email, 210, muted=True)
        col(address, 180, muted=True)
        col(notes, 160, muted=True)

        def on_click(_=None):
            select_card(pid)

        def on_enter(_=None):
            if state["selected"] != pid:
                card.configure(fg_color=CARD_HOVER)

        def on_leave(_=None):
            if state["selected"] != pid:
                card.configure(fg_color="white")

        bind_recursive(card, "<Button-1>", on_click)
        bind_recursive(card, "<Double-Button-1>", lambda e: (select_card(pid), edit_patient()))
        bind_recursive(card, "<Enter>", on_enter)
        bind_recursive(card, "<Leave>", on_leave)

        state["cards"][pid] = card

    def load_data(filter_text=""):
        for w in listbox.winfo_children():
            w.destroy()
        state["cards"] = {}

        if filter_text:
            like = f"%{filter_text}%"
            cursor.execute("""
                SELECT id, name, surname, phone, email, address, notes FROM patients
                WHERE name LIKE ? OR surname LIKE ? OR phone LIKE ?
                ORDER BY id DESC
            """, (like, like, like))
        else:
            cursor.execute("""
                SELECT id, name, surname, phone, email, address, notes
                FROM patients ORDER BY id DESC
            """)

        rows = cursor.fetchall()
        count_label.configure(text=f"{len(rows)} pacientë")

        if not rows:
            empty_state(listbox, "🗂", "Nuk u gjet asnjë pacient")
            return

        for row in rows:
            build_card(row)

        if state["selected"] in state["cards"]:
            select_card(state["selected"])

    load_data()
    search_entry.bind("<KeyRelease>", lambda e: load_data(search_entry.get().strip()))

    # ---------- ACTIONS ----------
    def save_patient():
        if entries["name"].get().strip() == "" or entries["surname"].get().strip() == "":
            messagebox.showwarning("Gabim", "Emri dhe mbiemri janë të detyrueshme")
            return

        values = tuple(entries[k].get().strip() for k in
                       ("name", "surname", "phone", "email", "address", "notes"))

        if state["editing"] is None:
            cursor.execute(
                "INSERT INTO patients (name, surname, phone, email, address, notes) VALUES (?,?,?,?,?,?)",
                values)
            conn.commit()
            messagebox.showinfo("Sukses", "Pacienti u ruajt")
        else:
            cursor.execute(
                "UPDATE patients SET name=?, surname=?, phone=?, email=?, address=?, notes=? WHERE id=?",
                values + (state["editing"],))
            conn.commit()
            messagebox.showinfo("Sukses", "Pacienti u përditësua")

        clear_form()
        load_data(search_entry.get().strip())

    save_btn = ctk.CTkButton(
        form, text="💾   Ruaj Pacient", command=save_patient,
        corner_radius=9, height=38, width=210, fg_color=SUCCESS, hover_color=GOLD,
        font=("Segoe UI", 13, "bold"),
    )
    save_btn.grid(row=1, column=3, padx=10, pady=12)

    def require_selection():
        if state["selected"] is None:
            messagebox.showwarning("Gabim", "Kliko mbi një pacient në listë")
            return None
        cursor.execute("SELECT id, name, surname, phone, email, address, notes FROM patients WHERE id=?",
                       (state["selected"],))
        return cursor.fetchone()

    def edit_patient():
        row = require_selection()
        if not row:
            return
        clear_form()
        state["editing"] = row[0]
        for key, val in zip(("name", "surname", "phone", "email", "address", "notes"), row[1:]):
            entries[key].insert(0, val if val else "")
        save_btn.configure(text="✏️   Përditëso", fg_color=GOLD)

    def delete_patient():
        row = require_selection()
        if not row:
            return
        if not messagebox.askyesno("Konfirmo", f"Fshi pacientin {row[1]} {row[2]} dhe të gjitha terminet e tij?"):
            return
        cursor.execute("DELETE FROM patients WHERE id=?", (row[0],))
        cursor.execute("DELETE FROM appointments WHERE patient_id=?", (row[0],))
        conn.commit()
        state["selected"] = None
        clear_form()
        load_data(search_entry.get().strip())

    def show_history():
        row = require_selection()
        if not row:
            return
        pid, full_name = row[0], f"{row[1]} {row[2]}"

        win = ctk.CTkToplevel(app)
        win.title(f"Historiku — {full_name}")
        win.geometry("720x480")
        win.configure(fg_color=CREAM)
        win.after(200, win.lift)

        top = ctk.CTkFrame(win, fg_color=DARK_BROWN, corner_radius=0, height=80)
        top.pack(fill="x")
        top.pack_propagate(False)
        avatar(top, initials(row[1], row[2]), GOLD, 46).pack(side="left", padx=18, pady=17)
        box = ctk.CTkFrame(top, fg_color="transparent")
        box.pack(side="left")
        ctk.CTkLabel(box, text=full_name, font=("Georgia", 18, "bold"), text_color=GOLD_LIGHT).pack(anchor="w", pady=(18, 0))
        ctk.CTkLabel(box, text="Historiku i trajtimeve", font=("Segoe UI", 11), text_color="#D9C9A8").pack(anchor="w")

        body = ctk.CTkScrollableFrame(win, fg_color=CREAM)
        body.pack(fill="both", expand=True, padx=16, pady=14)

        cursor.execute("""
            SELECT date, time, treatment, price, status FROM appointments
            WHERE patient_id=? ORDER BY date DESC
        """, (pid,))
        rows = cursor.fetchall()

        cursor.execute("SELECT COALESCE(SUM(price),0) FROM appointments WHERE patient_id=? AND status='Përfunduar'",
                       (pid,))
        total = cursor.fetchone()[0]

        if not rows:
            empty_state(body, "📭", "Ky pacient nuk ka ende asnjë termin")
        else:
            for d, t, treat, price, status in rows:
                c = ctk.CTkFrame(body, fg_color="white", corner_radius=10, border_width=1, border_color=LINE, height=58)
                c.pack(fill="x", pady=4)
                c.pack_propagate(False)
                ctk.CTkLabel(c, text=f"{d}\n{t}", font=("Segoe UI", 11, "bold"),
                             text_color=BROWN, width=90).pack(side="left", padx=(14, 6))
                ctk.CTkLabel(c, text=treat if treat else "—", font=("Segoe UI", 13),
                             text_color=TEXT_DARK, anchor="w", width=250).pack(side="left")
                pill(c, status, STATUS_COLORS.get(status, NEUTRAL)).pack(side="right", padx=14)
                ctk.CTkLabel(c, text=f"{float(price or 0):.2f} €", font=("Segoe UI", 13, "bold"),
                             text_color=SUCCESS, width=100, anchor="e").pack(side="right")

            footer = ctk.CTkFrame(win, fg_color=CARD, corner_radius=0, height=48)
            footer.pack(fill="x")
            footer.pack_propagate(False)
            ctk.CTkLabel(footer, text=f"Total i paguar (terminet e përfunduara):  {total:.2f} €",
                         font=("Segoe UI", 13, "bold"), text_color=DARK_BROWN).pack(pady=13)

    actions = ctk.CTkFrame(content, fg_color="white")
    actions.pack(pady=(6, 14))

    btns = [
        ("✏️   Modifiko", edit_patient, GOLD, BROWN),
        ("📋   Historiku", show_history, BROWN, GOLD),
        ("🗑   Fshi", delete_patient, DANGER, "#6E2A2A"),
        ("✖   Anulo", clear_form, NEUTRAL, "#7A6A5B"),
    ]
    for i, (txt, cmd, fg, hv) in enumerate(btns):
        ctk.CTkButton(actions, text=txt, command=cmd, corner_radius=9, height=38,
                      width=150, fg_color=fg, hover_color=hv,
                      font=("Segoe UI", 13, "bold")).grid(row=0, column=i, padx=7)


# ================= APPOINTMENTS PAGE =================

def appointments_page():
    clear_content()
    state = {"editing": None, "selected": None, "cards": {}}

    section_title("📅  Terminet", "Planifiko dhe ndiq terminet e klinikës")

    # ---------- FORM ----------
    form = ctk.CTkFrame(content, fg_color=CARD, corner_radius=14, border_width=1, border_color=GOLD_LIGHT)
    form.pack(fill="x", padx=28, pady=(10, 6))

    cursor.execute("SELECT id, name, surname FROM patients ORDER BY name")
    patient_values = [f"{p[0]} - {p[1]} {p[2]}" for p in cursor.fetchall()]

    patient_combo = ctk.CTkComboBox(
        form, values=patient_values or ["(asnjë pacient)"], width=250, height=38,
        corner_radius=9, fg_color="white", border_color=LINE,
        button_color=BROWN, button_hover_color=GOLD, dropdown_hover_color=CARD,
    )
    patient_combo.set("")
    patient_combo.grid(row=0, column=0, padx=10, pady=12)

    date_entry = ctk.CTkEntry(form, placeholder_text="Data  DD/MM/YYYY", height=38,
                              corner_radius=9, width=200, fg_color="white", border_color=LINE)
    date_entry.insert(0, datetime.now().strftime("%d/%m/%Y"))
    date_entry.grid(row=0, column=1, padx=10, pady=12)

    time_entry = ctk.CTkEntry(form, placeholder_text="Ora  HH:MM", height=38,
                              corner_radius=9, width=160, fg_color="white", border_color=LINE)
    time_entry.grid(row=0, column=2, padx=10, pady=12)

    treatment_entry = ctk.CTkEntry(form, placeholder_text="Trajtimi", height=38,
                                   corner_radius=9, width=250, fg_color="white", border_color=LINE)
    treatment_entry.grid(row=1, column=0, padx=10, pady=12)

    price_entry = ctk.CTkEntry(form, placeholder_text="Çmimi €", height=38,
                               corner_radius=9, width=200, fg_color="white", border_color=LINE)
    price_entry.grid(row=1, column=1, padx=10, pady=12)

    status_combo = ctk.CTkComboBox(
        form, values=["Planifikuar", "Përfunduar", "Anulluar"], height=38, width=160,
        corner_radius=9, fg_color="white", border_color=LINE,
        button_color=BROWN, button_hover_color=GOLD, dropdown_hover_color=CARD,
    )
    status_combo.set("Planifikuar")
    status_combo.grid(row=1, column=2, padx=10, pady=12)

    def clear_form():
        patient_combo.set("")
        date_entry.delete(0, "end")
        date_entry.insert(0, datetime.now().strftime("%d/%m/%Y"))
        time_entry.delete(0, "end")
        treatment_entry.delete(0, "end")
        price_entry.delete(0, "end")
        status_combo.set("Planifikuar")
        state["editing"] = None
        save_btn.configure(text="📅   Ruaj Termin", fg_color=SUCCESS)

    # ---------- SEARCH + FILTER ----------
    bar = ctk.CTkFrame(content, fg_color="white")
    bar.pack(fill="x", padx=28, pady=(8, 0))

    search_entry = ctk.CTkEntry(
        bar, placeholder_text="🔍   Kërko sipas pacientit, datës ose trajtimit...",
        width=420, height=38, corner_radius=19, border_color=LINE, fg_color=CREAM,
    )
    search_entry.pack(side="left")

    filter_combo = ctk.CTkComboBox(
        bar, values=["Të gjitha", "Planifikuar", "Përfunduar", "Anulluar"],
        width=160, height=38, corner_radius=19, fg_color=CREAM, border_color=LINE,
        button_color=BROWN, button_hover_color=GOLD, dropdown_hover_color=CARD,
        command=lambda choice: load_appointments(search_entry.get().strip()),
    )
    filter_combo.set("Të gjitha")
    filter_combo.pack(side="left", padx=12)

    count_label = ctk.CTkLabel(bar, text="", font=("Segoe UI", 12), text_color=TEXT_MUTED)
    count_label.pack(side="right", padx=6)

    # ---------- LIST HEADER ----------
    head = ctk.CTkFrame(content, fg_color=DARK_BROWN, corner_radius=10, height=38)
    head.pack(fill="x", padx=28, pady=(12, 4))
    head.pack_propagate(False)

    for text, w in [("DATA / ORA", 110), ("PACIENTI", 250), ("TRAJTIMI", 250)]:
        ctk.CTkLabel(head, text=text, width=w, anchor="w",
                     font=("Segoe UI", 11, "bold"), text_color=GOLD_LIGHT).pack(side="left", padx=(14, 0))
    ctk.CTkLabel(head, text="STATUSI", width=120, anchor="e",
                 font=("Segoe UI", 11, "bold"), text_color=GOLD_LIGHT).pack(side="right", padx=(0, 20))
    ctk.CTkLabel(head, text="ÇMIMI", width=100, anchor="e",
                 font=("Segoe UI", 11, "bold"), text_color=GOLD_LIGHT).pack(side="right")

    # ---------- CARD LIST ----------
    listbox = ctk.CTkScrollableFrame(content, fg_color="white", corner_radius=0)
    listbox.pack(fill="both", expand=True, padx=22, pady=(0, 4))

    def select_card(aid):
        state["selected"] = aid
        for other_id, card in state["cards"].items():
            if other_id == aid:
                card.configure(fg_color=CARD, border_color=GOLD, border_width=2)
            else:
                card.configure(fg_color="white", border_color=LINE, border_width=1)

    def build_card(row):
        aid, pname, d, t, treat, price, status = row
        color = STATUS_COLORS.get(status, NEUTRAL)

        card = ctk.CTkFrame(listbox, fg_color="white", corner_radius=12,
                            border_width=1, border_color=LINE, height=70)
        card.pack(fill="x", pady=4, padx=4)
        card.pack_propagate(False)

        # colored status stripe on the left
        stripe = ctk.CTkFrame(card, width=5, fg_color=color, corner_radius=3)
        stripe.pack(side="left", fill="y", padx=(6, 8), pady=10)

        # date / time
        dt = ctk.CTkFrame(card, fg_color="transparent")
        dt.pack(side="left", padx=(4, 0))
        ctk.CTkLabel(dt, text=d, font=("Segoe UI", 12, "bold"),
                     text_color=TEXT_DARK, width=104, anchor="w").pack(anchor="w", pady=(14, 0))
        ctk.CTkLabel(dt, text=f"🕐 {t}", font=("Segoe UI", 11),
                     text_color=TEXT_MUTED, width=104, anchor="w").pack(anchor="w")

        # patient with avatar
        pat = ctk.CTkFrame(card, fg_color="transparent")
        pat.pack(side="left", padx=(14, 0))
        parts = pname.split(" ", 1)
        avatar(pat, initials(parts[0], parts[1] if len(parts) > 1 else ""), BROWN, 38).pack(side="left", pady=16)
        ctk.CTkLabel(pat, text=pname, font=("Segoe UI", 13, "bold"),
                     text_color=TEXT_DARK, width=190, anchor="w").pack(side="left", padx=10)

        # treatment
        ctk.CTkLabel(card, text=treat if treat else "—", font=("Segoe UI", 13),
                     text_color=TEXT_DARK, width=250, anchor="w").pack(side="left", padx=(14, 0))

        # status pill (right)
        holder = ctk.CTkFrame(card, fg_color="transparent", width=130)
        holder.pack(side="right", padx=(0, 18))
        pill(holder, status, color).pack(pady=21)

        # price
        ctk.CTkLabel(card, text=f"{float(price or 0):.2f} €", font=("Segoe UI", 14, "bold"),
                     text_color=SUCCESS, width=100, anchor="e").pack(side="right")

        def on_enter(_=None):
            if state["selected"] != aid:
                card.configure(fg_color=CARD_HOVER)

        def on_leave(_=None):
            if state["selected"] != aid:
                card.configure(fg_color="white")

        bind_recursive(card, "<Button-1>", lambda e: select_card(aid))
        bind_recursive(card, "<Double-Button-1>", lambda e: (select_card(aid), edit_appointment()))
        bind_recursive(card, "<Enter>", on_enter)
        bind_recursive(card, "<Leave>", on_leave)

        state["cards"][aid] = card

    def load_appointments(filter_text=""):
        for w in listbox.winfo_children():
            w.destroy()
        state["cards"] = {}

        query = """
            SELECT appointments.id, patients.name || ' ' || patients.surname,
                   appointments.date, appointments.time, appointments.treatment,
                   appointments.price, appointments.status
            FROM appointments
            JOIN patients ON appointments.patient_id = patients.id
            WHERE 1=1
        """
        params = []

        if filter_text:
            like = f"%{filter_text}%"
            query += """ AND (patients.name LIKE ? OR patients.surname LIKE ?
                         OR appointments.date LIKE ? OR appointments.treatment LIKE ?)"""
            params += [like, like, like, like]

        chosen = filter_combo.get()
        if chosen != "Të gjitha":
            query += " AND appointments.status = ?"
            params.append(chosen)

        query += " ORDER BY substr(appointments.date,7,4) DESC, substr(appointments.date,4,2) DESC, substr(appointments.date,1,2) DESC, appointments.time"

        cursor.execute(query, params)
        rows = cursor.fetchall()
        count_label.configure(text=f"{len(rows)} termine")

        if not rows:
            empty_state(listbox, "📅", "Nuk u gjet asnjë termin")
            return

        for row in rows:
            build_card(row)

        if state["selected"] in state["cards"]:
            select_card(state["selected"])

    load_appointments()
    search_entry.bind("<KeyRelease>", lambda e: load_appointments(search_entry.get().strip()))

    # ---------- ACTIONS ----------
    def save_appointment():
        selected = patient_combo.get()
        if selected == "" or selected == "(asnjë pacient)":
            messagebox.showwarning("Gabim", "Zgjidh pacientin")
            return
        if not is_valid_date(date_entry.get()):
            messagebox.showwarning("Gabim", "Data duhet të jetë në formatin DD/MM/YYYY")
            return
        if not is_valid_time(time_entry.get()):
            messagebox.showwarning("Gabim", "Ora duhet të jetë në formatin HH:MM")
            return
        if not is_valid_price(price_entry.get()):
            messagebox.showwarning("Gabim", "Çmimi duhet të jetë numër pozitiv")
            return
        if treatment_entry.get().strip() == "":
            messagebox.showwarning("Gabim", "Shkruaj trajtimin")
            return

        patient_id = selected.split("-")[0].strip()
        values = (
            patient_id, date_entry.get().strip(), time_entry.get().strip(),
            treatment_entry.get().strip(), float(price_entry.get().replace(",", ".")),
            status_combo.get(), "",
        )

        if state["editing"] is None:
            cursor.execute("""INSERT INTO appointments
                (patient_id, date, time, treatment, price, status, notes)
                VALUES (?,?,?,?,?,?,?)""", values)
            conn.commit()
            messagebox.showinfo("Sukses", "Termini u ruajt")
        else:
            cursor.execute("""UPDATE appointments SET
                patient_id=?, date=?, time=?, treatment=?, price=?, status=?, notes=?
                WHERE id=?""", values + (state["editing"],))
            conn.commit()
            messagebox.showinfo("Sukses", "Termini u përditësua")

        clear_form()
        load_appointments(search_entry.get().strip())

    save_btn = ctk.CTkButton(
        form, text="📅   Ruaj Termin", command=save_appointment,
        corner_radius=9, height=38, width=200, fg_color=SUCCESS, hover_color=GOLD,
        font=("Segoe UI", 13, "bold"),
    )
    save_btn.grid(row=0, column=3, rowspan=2, padx=14, pady=12)

    def edit_appointment():
        if state["selected"] is None:
            messagebox.showwarning("Gabim", "Kliko mbi një termin në listë")
            return

        cursor.execute("""SELECT patient_id, date, time, treatment, price, status
                          FROM appointments WHERE id=?""", (state["selected"],))
        row = cursor.fetchone()
        if not row:
            return

        state["editing"] = state["selected"]
        cursor.execute("SELECT name, surname FROM patients WHERE id=?", (row[0],))
        p = cursor.fetchone()
        patient_combo.set(f"{row[0]} - {p[0]} {p[1]}")

        date_entry.delete(0, "end"); date_entry.insert(0, row[1])
        time_entry.delete(0, "end"); time_entry.insert(0, row[2])
        treatment_entry.delete(0, "end"); treatment_entry.insert(0, row[3])
        price_entry.delete(0, "end"); price_entry.insert(0, f"{float(row[4] or 0):.2f}")
        status_combo.set(row[5])

        save_btn.configure(text="✏️   Përditëso", fg_color=GOLD)

    def delete_appointment():
        if state["selected"] is None:
            messagebox.showwarning("Gabim", "Kliko mbi një termin në listë")
            return
        if not messagebox.askyesno("Konfirmo", "Fshi këtë termin?"):
            return
        cursor.execute("DELETE FROM appointments WHERE id=?", (state["selected"],))
        conn.commit()
        state["selected"] = None
        clear_form()
        load_appointments(search_entry.get().strip())

    def mark_done():
        if state["selected"] is None:
            messagebox.showwarning("Gabim", "Kliko mbi një termin në listë")
            return
        cursor.execute("UPDATE appointments SET status='Përfunduar' WHERE id=?", (state["selected"],))
        conn.commit()
        load_appointments(search_entry.get().strip())

    actions = ctk.CTkFrame(content, fg_color="white")
    actions.pack(pady=(6, 14))

    btns = [
        ("✔   Përfundo", mark_done, SUCCESS, GOLD),
        ("✏️   Modifiko", edit_appointment, GOLD, BROWN),
        ("🗑   Fshi", delete_appointment, DANGER, "#6E2A2A"),
        ("✖   Anulo", clear_form, NEUTRAL, "#7A6A5B"),
    ]
    for i, (txt, cmd, fg, hv) in enumerate(btns):
        ctk.CTkButton(actions, text=txt, command=cmd, corner_radius=9, height=38,
                      width=150, fg_color=fg, hover_color=hv,
                      font=("Segoe UI", 13, "bold")).grid(row=0, column=i, padx=7)


# ================= DASHBOARD =================

def dashboard_page():
    clear_content()
    section_title("🏠  Dashboard", datetime.now().strftime("Sot është %d/%m/%Y"))

    cursor.execute("SELECT COUNT(*) FROM patients")
    patients_count = cursor.fetchone()[0]
    cursor.execute("SELECT COUNT(*) FROM appointments")
    appointments_count = cursor.fetchone()[0]
    today_str = datetime.now().strftime("%d/%m/%Y")
    cursor.execute("SELECT COUNT(*) FROM appointments WHERE date=?", (today_str,))
    today_count = cursor.fetchone()[0]
    cursor.execute("SELECT COALESCE(SUM(price),0) FROM appointments WHERE status='Përfunduar'")
    revenue = cursor.fetchone()[0]

    card_frame = ctk.CTkFrame(content, fg_color="white")
    card_frame.pack(pady=14)

    def stat_card(col, icon_text, label, value, color):
        card = ctk.CTkFrame(card_frame, width=250, height=150, fg_color=color, corner_radius=18)
        card.grid(row=0, column=col, padx=15, pady=10)
        card.grid_propagate(False)

        badge = ctk.CTkFrame(card, width=52, height=52, corner_radius=26, fg_color="white")
        badge.place(relx=0.5, y=34, anchor="center")
        ctk.CTkLabel(badge, text=icon_text, font=("Segoe UI", 21)).place(relx=0.5, rely=0.5, anchor="center")

        ctk.CTkLabel(card, text=label, font=("Segoe UI", 13, "bold"), text_color="white").place(relx=0.5, y=92, anchor="center")
        ctk.CTkLabel(card, text=str(value), font=("Georgia", 26, "bold"), text_color="white").place(relx=0.5, y=122, anchor="center")

    stat_card(0, "👤", "Pacientë", patients_count, BROWN)
    stat_card(1, "📅", "Termine", appointments_count, GOLD)
    stat_card(2, "⏰", "Termine Sot", today_count, DARK_BROWN)
    stat_card(3, "💶", "Të Ardhura", f"{revenue:,.0f} €", SUCCESS)

    # ---------- STATUS CHART ----------
    cursor.execute("SELECT status, COUNT(*) FROM appointments GROUP BY status")
    status_counts = dict(cursor.fetchall())

    ctk.CTkLabel(content, text="📊  Statuset e Termineve",
                 font=("Georgia", 17, "bold"), text_color=DARK_BROWN).pack(pady=(24, 8))

    chart_card = ctk.CTkFrame(content, fg_color=CARD, corner_radius=16,
                              border_width=1, border_color=GOLD_LIGHT)
    chart_card.pack(fill="x", padx=70, pady=(0, 10))

    max_count = max(status_counts.values()) if status_counts else 1

    for status_name in ["Planifikuar", "Përfunduar", "Anulluar"]:
        count = status_counts.get(status_name, 0)
        row = ctk.CTkFrame(chart_card, fg_color=CARD)
        row.pack(fill="x", padx=22, pady=9)

        ctk.CTkLabel(row, text=status_name, width=110, anchor="w",
                     font=("Segoe UI", 13, "bold"), text_color=TEXT_DARK).pack(side="left")

        bar_bg = ctk.CTkFrame(row, width=340, height=18, fg_color="white", corner_radius=9)
        bar_bg.pack(side="left", padx=10)
        bar_bg.pack_propagate(False)

        bar_width = int(340 * (count / max_count)) if max_count else 0
        ctk.CTkFrame(bar_bg, width=max(bar_width, 6), height=18,
                     fg_color=STATUS_COLORS[status_name], corner_radius=9).place(x=0, y=0)

        ctk.CTkLabel(row, text=str(count), font=("Segoe UI", 13, "bold"),
                     text_color=TEXT_DARK).pack(side="left", padx=8)


# ================= EXPORT EXCEL =================

def export_excel():
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Pacientet"
    sheet.append(["ID", "Emri", "Mbiemri", "Telefon", "Email", "Adresa", "Shënime"])

    cursor.execute("SELECT * FROM patients")
    for row in cursor.fetchall():
        sheet.append(row)

    sheet2 = workbook.create_sheet("Terminet")
    sheet2.append(["ID", "Pacienti", "Data", "Ora", "Trajtimi", "Çmimi", "Status"])

    cursor.execute("""
        SELECT appointments.id, patients.name || ' ' || patients.surname,
               appointments.date, appointments.time, appointments.treatment,
               appointments.price, appointments.status
        FROM appointments
        JOIN patients ON appointments.patient_id = patients.id
    """)
    for row in cursor.fetchall():
        sheet2.append(row)

    filename = f"Vividus_Raport_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    workbook.save(filename)
    messagebox.showinfo("Eksport", f"U krijua {filename}")


# ================= MENU =================

menu_button("  🏠   Dashboard", dashboard_page)
menu_button("  👤   Pacientët", patients_page)
menu_button("  📅   Terminet", appointments_page)
menu_button("  📤   Eksporto Excel", export_excel)


def close_app():
    backup_database()
    conn.close()
    app.destroy()


menu_button("  🚪   Dil", close_app)
app.protocol("WM_DELETE_WINDOW", close_app)


# ================= TODAY'S REMINDER =================

def show_today_reminder():
    today_str = datetime.now().strftime("%d/%m/%Y")
    cursor.execute("""
        SELECT patients.name || ' ' || patients.surname, appointments.time, appointments.treatment
        FROM appointments
        JOIN patients ON appointments.patient_id = patients.id
        WHERE appointments.date=? AND appointments.status='Planifikuar'
        ORDER BY appointments.time
    """, (today_str,))
    rows = cursor.fetchall()

    if rows:
        lines = [f"• {r[1]} — {r[0]} ({r[2]})" for r in rows]
        messagebox.showinfo("Terminet e sotme", f"Ke {len(rows)} termine sot:\n\n" + "\n".join(lines))


# ================= START =================

dashboard_page()
app.after(400, show_today_reminder)
app.mainloop()