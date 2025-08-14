#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Dienstplan Flohkiste - Windows GUI App (Tkinter)
Funktionen:
- Mitarbeiterverwaltung (Name, Wochenarbeitszeit hh:mm, Aktiv, Übertrag/Saldo)
- Wochenplanung (Mo–Fr): Status (Arbeitstag/Urlaub/Krank/Feiertag), Start1/Ende1, Start2/Ende2
- Automatische Pausenregel: bis 9:15 -> 0:30, darüber -> 0:45 (pro Tag, Summe beider Blöcke)
- Berechnung pro Mitarbeiter: Tagesnetto, Wochensumme, Differenz zu Soll, neuer Übertrag
- Speicherung in JSON (flohkiste_data.json) im Arbeitsverzeichnis
- PDF-Export (A4 Querformat) mit ReportLab (wenn installiert)

Erstellt für: Windows (läuft aber auch unter macOS/Linux mit Python 3.9+)
Benötigt: Python 3.9+ und (für PDF) reportlab
"""

import json
import os
import sys
import re
from datetime import datetime, timedelta
from dataclasses import dataclass, field
from typing import Dict, List, Optional

import tkinter as tk
from tkinter import ttk, messagebox, filedialog

DATA_FILE = "flohkiste_data.json"

# ---------- Hilfsfunktionen für Zeit ----------

TIME_RE = re.compile(r"^([0-2]?\d):([0-5]\d)$")

def hhmm_to_minutes(s: str) -> int:
    """Konvertiere 'HH:MM' nach Minuten. Leere Strings -> 0.
       Ungültiges Format wirft ValueError."""
    s = (s or "").strip()
    if not s:
        return 0
    m = TIME_RE.match(s)
    if not m:
        raise ValueError(f"Ungültiges Zeitformat: {s}. Bitte HH:MM eingeben (z. B. 08:30).")
    h = int(m.group(1))
    mm = int(m.group(2))
    return h * 60 + mm

def minutes_to_hhmm(m: int, allow_negative: bool = True) -> str:
    """Konvertiere Minuten nach 'HH:MM'. Bei allow_negative werden negative Zeiten mit '-' dargestellt."""
    if allow_negative and m < 0:
        return "-" + minutes_to_hhmm(-m, allow_negative=False)
    h = m // 60
    mm = m % 60
    return f"{h:02d}:{mm:02d}"

# ---------- Datenklassen ----------

@dataclass
class Employee:
    id: str
    name: str
    weekly_target_min: int = 0
    active: bool = True
    balance_min: int = 0   # kumulierter Übertrag (+/-) nach letzter abgeschlossener Woche

@dataclass
class DayEntry:
    status: str = "Arbeitstag"  # "Arbeitstag", "Urlaub", "Krank", "Feiertag"
    start1: str = ""
    end1: str = ""
    start2: str = ""
    end2: str = ""

    def total_minutes_raw(self) -> int:
        def span(a,b):
            try:
                am = hhmm_to_minutes(a)
                bm = hhmm_to_minutes(b)
                return max(bm - am, 0)
            except ValueError:
                return 0
        return span(self.start1, self.end1) + span(self.start2, self.end2)

    def net_minutes_with_pause(self, pause_short=30, pause_long=45, threshold_min=555) -> int:
        """Pausenregel: bis threshold (inkl.) -> pause_short, darüber -> pause_long. Gilt nur an Arbeitstagen."""
        if self.status != "Arbeitstag":
            return 0
        total = self.total_minutes_raw()
        if total <= 0:
            return 0
        pause = pause_short if total <= threshold_min else pause_long
        return max(total - pause, 0)

@dataclass
class WeekEmployeeData:
    carry_prev_min: int = 0
    days: Dict[str, DayEntry] = field(default_factory=lambda: {
        "Montag": DayEntry(), "Dienstag": DayEntry(), "Mittwoch": DayEntry(),
        "Donnerstag": DayEntry(), "Freitag": DayEntry()
    })

    def weekly_ist(self, pause_short=30, pause_long=45, threshold_min=555) -> int:
        return sum(d.net_minutes_with_pause(pause_short, pause_long, threshold_min) for d in self.days.values())

# ---------- Persistenz ----------

def load_data() -> dict:
    if os.path.exists(DATA_FILE):
        with open(DATA_FILE, "r", encoding="utf-8") as f:
            try:
                return json.load(f)
            except json.JSONDecodeError:
                return {}
    return {}

def save_data(data: dict):
    tmp = DATA_FILE + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    os.replace(tmp, DATA_FILE)

def ensure_structure(data: dict) -> dict:
    data.setdefault("employees", [])
    data.setdefault("weeks", {})  # key: "YYYY-KW" (ISO Kalenderwoche)
    data.setdefault("settings", {
        "pause_threshold_min": 9*60+15,  # 9:15
        "pause_short_min": 30,
        "pause_long_min": 45,
        "count_vacation_as_work": False,   # Urlaub/Krank/Feiertag als Soll? (Default False)
    })
    return data
# ---------- GUI ----------

class FlohkisteApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Dienstplan Flohkiste")
        self.geometry("1200x700")
        self.minsize(1100, 600)

        self.data = ensure_structure(load_data())

        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True)

        self.page_employees = EmployeesPage(self.notebook, self)
        self.page_week = WeekPage(self.notebook, self)
        self.page_settings = SettingsPage(self.notebook, self)
        self.page_export = ExportPage(self.notebook, self)

        self.notebook.add(self.page_employees, text="Mitarbeiter")
        self.notebook.add(self.page_week, text="Woche")
        self.notebook.add(self.page_settings, text="Einstellungen")
        self.notebook.add(self.page_export, text="Export / PDF")

        self.protocol("WM_DELETE_WINDOW", self.on_close)

    def on_close(self):
        # Beim Schließen automatisch speichern
        save_data(self.data)
        self.destroy()

    # --- helpers ---
    def employees(self) -> List[Employee]:
        res = []
        for e in self.data["employees"]:
            res.append(Employee(
                id=e.get("id",""),
                name=e.get("name",""),
                weekly_target_min=int(e.get("weekly_target_min",0)),
                active=bool(e.get("active", True)),
                balance_min=int(e.get("balance_min",0)),
            ))
        return res

    def set_employees(self, employees: List[Employee]):
        self.data["employees"] = [{
            "id": e.id, "name": e.name, "weekly_target_min": e.weekly_target_min,
            "active": e.active, "balance_min": e.balance_min
        } for e in employees]
        save_data(self.data)

    def get_week_key(self, year: int, kw: int) -> str:
        return f"{year}-{kw:02d}"

    def get_week_data(self, emp_id: str, year: int, kw: int) -> 'WeekEmployeeData':
        key = self.get_week_key(year, kw)
        self.data["weeks"].setdefault(key, {"employees":{}})
        empstore = self.data["weeks"][key]["employees"].setdefault(emp_id, {})
        # build object
        wed = WeekEmployeeData()
        wed.carry_prev_min = int(empstore.get("carry_prev_min", 0))
        # days
        wed.days = {}
        for day in ["Montag","Dienstag","Mittwoch","Donnerstag","Freitag"]:
            di = empstore.get("days", {}).get(day, {})
            wed.days[day] = DayEntry(
                status=di.get("status", "Arbeitstag"),
                start1=di.get("start1",""), end1=di.get("end1",""),
                start2=di.get("start2",""), end2=di.get("end2","")
            )
        return wed

    def set_week_data(self, emp_id: str, year: int, kw: int, wed: 'WeekEmployeeData'):
        key = self.get_week_key(year, kw)
        self.data["weeks"].setdefault(key, {"employees":{}})
        empstore = self.data["weeks"][key]["employees"].setdefault(emp_id, {})
        empstore["carry_prev_min"] = wed.carry_prev_min
        empstore["days"] = {day: {
            "status": de.status, "start1": de.start1, "end1": de.end1,
            "start2": de.start2, "end2": de.end2
        } for day, de in wed.days.items()}
        save_data(self.data)

# ---------- Employees Page ----------

class EmployeesPage(ttk.Frame):
    def __init__(self, parent, app: FlohkisteApp):
        super().__init__(parent)
        self.app = app()

        # left: list, right: form
        self.columnconfigure(0, weight=1)
        self.columnconfigure(1, weight=1)
        self.rowconfigure(0, weight=1)

        # Treeview
        cols = ("name","target","active","balance")
        self.tree = ttk.Treeview(self, columns=cols, show="headings", height=18)
        self.tree.heading("name", text="Name")
        self.tree.heading("target", text="Wochenarbeitszeit (hh:mm)")
        self.tree.heading("active", text="Aktiv")
        self.tree.heading("balance", text="Übertrag (±hh:mm)")
        self.tree.column("name", width=220)
        self.tree.column("target", width=180, anchor="center")
        self.tree.column("active", width=70, anchor="center")
        self.tree.column("balance", width=130, anchor="center")
        self.tree.grid(row=0, column=0, sticky="nsew", padx=8, pady=8)

        self.scroll = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=self.scroll.set)
        self.scroll.grid(row=0, column=0, sticky="nse", padx=(0,8), pady=8)

        # Form
        frm = ttk.LabelFrame(self, text="Mitarbeiterdaten")
        frm.grid(row=0, column=1, sticky="nsew", padx=8, pady=8)
        frm.columnconfigure(1, weight=1)

        ttk.Label(frm, text="Name:").grid(row=0, column=0, sticky="w", padx=6, pady=6)
        self.var_name = tk.StringVar()
        ttk.Entry(frm, textvariable=self.var_name).grid(row=0, column=1, sticky="ew", padx=6, pady=6)

        ttk.Label(frm, text="Wochenarbeitszeit (hh:mm):").grid(row=1, column=0, sticky="w", padx=6, pady=6)
        self.var_target = tk.StringVar(value="38:30")
        ttk.Entry(frm, textvariable=self.var_target).grid(row=1, column=1, sticky="ew", padx=6, pady=6)

        self.var_active = tk.BooleanVar(value=True)
        ttk.Checkbutton(frm, text="Aktiv", variable=self.var_active).grid(row=2, column=1, sticky="w", padx=6, pady=6)

        ttk.Label(frm, text="Übertrag (±hh:mm):").grid(row=3, column=0, sticky="w", padx=6, pady=6)
        self.var_balance = tk.StringVar(value="00:00")
        ttk.Entry(frm, textvariable=self.var_balance).grid(row=3, column=1, sticky="ew", padx=6, pady=6)

        # Buttons
        btns = ttk.Frame(frm)
        btns.grid(row=4, column=0, columnspan=2, sticky="ew", padx=6, pady=8)
        for i in range(4): btns.columnconfigure(i, weight=1)

        ttk.Button(btns, text="Neu", command=self.on_new).grid(row=0, column=0, sticky="ew", padx=4)
        ttk.Button(btns, text="Speichern", command=self.on_save).grid(row=0, column=1, sticky="ew", padx=4)
        ttk.Button(btns, text="Löschen", command=self.on_delete).grid(row=0, column=2, sticky="ew", padx=4)
        ttk.Button(btns, text="Aktualisieren", command=self.reload).grid(row=0, column=3, sticky="ew", padx=4)

        self.selected_id: Optional[str] = None
        self.reload()

        self.tree.bind("<<TreeviewSelect>>", self.on_select)

    def reload(self):
        for i in self.tree.get_children():
            self.tree.delete(i)
        for e in self.app.employees():
            self.tree.insert("", "end", iid=e.id, values=(
                e.name, minutes_to_hhmm(e.weekly_target_min, False), "Ja" if e.active else "Nein",
                minutes_to_hhmm(e.balance_min, True)
            ))
        self.selected_id = None
        self.var_name.set("")
        self.var_target.set("38:30")
        self.var_active.set(True)
        self.var_balance.set("00:00")

    def on_new(self):
        # simple id
        new_id = f"emp_{int(datetime.now().timestamp()*1000)}"
        try:
            target_min = hhmm_to_minutes(self.var_target.get())
            balance_min = hhmm_to_minutes(self.var_balance.get())
        except ValueError as e:
            messagebox.showerror("Eingabefehler", str(e))
            return
        if not self.var_name.get().strip():
            messagebox.showerror("Eingabefehler", "Bitte einen Namen eingeben.")
            return
        emps = self.app.employees()
        emps.append(Employee(id=new_id, name=self.var_name.get().strip(),
                             weekly_target_min=target_min, active=self.var_active.get(),
                             balance_min=balance_min))
        self.app.set_employees(emps)
        self.reload()

    def on_save(self):
        if self.selected_id is None:
            messagebox.showinfo("Hinweis", "Bitte zuerst einen Mitarbeiter in der Liste auswählen oder 'Neu' anlegen.")
            return
        try:
            target_min = hhmm_to_minutes(self.var_target.get())
            balance_min = hhmm_to_minutes(self.var_balance.get())
        except ValueError as e:
            messagebox.showerror("Eingabefehler", str(e)); return
        emps = self.app.employees()
        for i, e in enumerate(emps):
            if e.id == self.selected_id:
                e.name = self.var_name.get().strip()
                e.weekly_target_min = target_min
                e.active = self.var_active.get()
                e.balance_min = balance_min
                emps[i] = e
                break
        self.app.set_employees(emps)
        self.reload()

    def on_delete(self):
        if self.selected_id is None:
            return
        if not messagebox.askyesno("Löschen bestätigen", "Mitarbeiter wirklich löschen?"):
            return
        emps = [e for e in self.app.employees() if e.id != self.selected_id]
        self.app.set_employees(emps)
        self.reload()

    def on_select(self, event):
        sel = self.tree.selection()
        if not sel:
            return
        iid = sel[0]
        self.selected_id = iid
        e = None
        for emp in self.app.employees():
            if emp.id == iid:
                e = emp; break
        if e:
            self.var_name.set(e.name)
            self.var_target.set(minutes_to_hhmm(e.weekly_target_min, False))
            self.var_active.set(e.active)
            self.var_balance.set(minutes_to_hhmm(e.balance_min, True))
# ---------- Week Page ----------

class WeekPage(ttk.Frame):
    STATUSES = ("Arbeitstag","Urlaub","Krank","Feiertag")
    DAYS = ("Montag","Dienstag","Mittwoch","Donnerstag","Freitag")

    def __init__(self, parent, app: FlohkisteApp):
        super().__init__(parent)
        self.app = app

        # Header with year/week
        header = ttk.Frame(self)
        header.pack(fill="x", padx=8, pady=6)

        ttk.Label(header, text="Jahr:").pack(side="left")
        self.var_year = tk.IntVar(value=datetime.now().year)
        ttk.Entry(header, textvariable=self.var_year, width=6).pack(side="left", padx=(4,12))

        ttk.Label(header, text="Kalenderwoche (KW):").pack(side="left")
        iso_kw = datetime.now().isocalendar().week
        self.var_kw = tk.IntVar(value=iso_kw)
        ttk.Entry(header, textvariable=self.var_kw, width=4).pack(side="left", padx=(4,12))

        ttk.Button(header, text="Speichern", command=self.save_all).pack(side="right", padx=4)
        ttk.Button(header, text="Neu berechnen", command=self.refresh_view).pack(side="right")

        # Scrollable canvas for employee day grids
        outer = ttk.Frame(self)
        outer.pack(fill="both", expand=True, padx=8, pady=8)

        self.canvas = tk.Canvas(outer)
        self.scrollbar = ttk.Scrollbar(outer, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        self.scrollbar.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)

        self.inner = ttk.Frame(self.canvas)
        self.canvas.create_window((0,0), window=self.inner, anchor="nw")

        self.inner.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)

        self.rows: List['EmployeeWeekRow'] = []
        self.build_rows()

    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")

    def build_rows(self):
        # clear
        for r in self.rows:
            r.frame.destroy()
        self.rows.clear()

        # headers row
        hdr = ttk.Frame(self.inner)
        hdr.grid(row=0, column=0, sticky="ew")
        for i in range(1, 12):
            hdr.columnconfigure(i, weight=1)
        ttk.Label(hdr, text="Mitarbeiter", width=22).grid(row=0, column=0, sticky="w", padx=4)

        col = 1
        for day in self.DAYS:
            ttk.Label(hdr, text=f"{day}\nStatus").grid(row=0, column=col, sticky="n", padx=2); col+=1
            ttk.Label(hdr, text=f"{day}\nStart1").grid(row=0, column=col, sticky="n", padx=2); col+=1
            ttk.Label(hdr, text=f"{day}\nEnde1").grid(row=0, column=col, sticky="n", padx=2); col+=1
            ttk.Label(hdr, text=f"{day}\nStart2").grid(row=0, column=col, sticky="n", padx=2); col+=1
            ttk.Label(hdr, text=f"{day}\nEnde2").grid(row=0, column=col, sticky="n", padx=2); col+=1
        ttk.Label(hdr, text="Übertrag\nVorw.").grid(row=0, column=col, sticky="n"); col+=1
        ttk.Label(hdr, text="Woche\nIst").grid(row=0, column=col, sticky="n"); col+=1
        ttk.Label(hdr, text="Woche\nSoll").grid(row=0, column=col, sticky="n"); col+=1
        ttk.Label(hdr, text="Diff\nIst-Soll").grid(row=0, column=col, sticky="n"); col+=1
        ttk.Label(hdr, text="Übertrag\nneu").grid(row=0, column=col, sticky="n")

        # employee rows
        r = 1
        for emp in self.app.employees():
            if not emp.active:
                continue
            row = EmployeeWeekRow(self.inner, self.app, emp, self.STATUSES, self.DAYS,
                                  self.var_year, self.var_kw)
            row.frame.grid(row=r, column=0, sticky="ew", pady=2)
            self.rows.append(row)
            r += 1

    def refresh_view(self):
        # rebuild rows so that recalculation happens from stored values
        self.build_rows()

    def save_all(self):
        year = self.var_year.get()
        kw = self.var_kw.get()
        # persist rows
        for row in self.rows:
            wed = row.collect_week_data()
            self.app.set_week_data(row.emp.id, year, kw, wed)
            # update employee balance with new carry
            # recompute using current settings
            pause_t = self.app.data["settings"]["pause_threshold_min"]
            pause_s = self.app.data["settings"]["pause_short_min"]
            pause_l = self.app.data["settings"]["pause_long_min"]
            weekly_ist = wed.weekly_ist(pause_s, pause_l, pause_t)
            emp = row.emp
            weekly_soll = emp.weekly_target_min
            diff = weekly_ist - weekly_soll
            new_carry = wed.carry_prev_min + diff
            # write back
            emps = self.app.employees()
            for i, e in enumerate(emps):
                if e.id == emp.id:
                    e.balance_min = new_carry
                    emps[i] = e
                    break
            self.app.set_employees(emps)

        messagebox.showinfo("Gespeichert", "Wochenplanung und Mitarbeiter-Überträge wurden gespeichert.")

class EmployeeWeekRow:
    def __init__(self, parent, app: FlohkisteApp, emp: Employee, statuses, days, var_year, var_kw):
        self.app = app
        self.emp = emp
        self.statuses = statuses
        self.days = days
        self.var_year = var_year
        self.var_kw = var_kw

        self.frame = ttk.Frame(parent)
        self.vars = {"carry": tk.StringVar(value=minutes_to_hhmm(emp.balance_min, True))}
        self.day_vars = {}
        ttk.Label(self.frame, text=emp.name, width=22).grid(row=0, column=0, sticky="w", padx=4)

        col = 1
        wed = app.get_week_data(emp.id, var_year.get(), var_kw.get())
        # If no explicit carry stored for this week, prefill with employee's current balance
        if wed.carry_prev_min == 0 and emp.balance_min:
            self.vars["carry"].set(minutes_to_hhmm(emp.balance_min, True))
        else:
            self.vars["carry"].set(minutes_to_hhmm(wed.carry_prev_min, True))

        # day cells
        for day in self.days:
            dentry = wed.days.get(day, DayEntry())
            # Status
            vs = tk.StringVar(value=dentry.status)
            cb = ttk.Combobox(self.frame, values=self.statuses, width=10, textvariable=vs, state="readonly")
            cb.grid(row=0, column=col, padx=1); col+=1
            # Times
            v_s1 = tk.StringVar(value=dentry.start1); e_s1 = ttk.Entry(self.frame, width=6, textvariable=v_s1); e_s1.grid(row=0, column=col, padx=1); col+=1
            v_e1 = tk.StringVar(value=dentry.end1); e_e1 = ttk.Entry(self.frame, width=6, textvariable=v_e1); e_e1.grid(row=0, column=col, padx=1); col+=1
            v_s2 = tk.StringVar(value=dentry.start2); e_s2 = ttk.Entry(self.frame, width=6, textvariable=v_s2); e_s2.grid(row=0, column=col, padx=1); col+=1
            v_e2 = tk.StringVar(value=dentry.end2); e_e2 = ttk.Entry(self.frame, width=6, textvariable=v_e2); e_e2.grid(row=0, column=col, padx=1); col+=1

            def on_status_change(var=vs, widgets=(e_s1,e_e1,e_s2,e_e2)):
                st = var.get()
                state = "normal" if st=="Arbeitstag" else "disabled"
                for w in widgets:
                    w.configure(state=state)

            cb.bind("<<ComboboxSelected>>", lambda e, v=vs, w=(e_s1,e_e1,e_s2,e_e2): on_status_change(v,w))
            on_status_change(vs, (e_s1,e_e1,e_s2,e_e2))

            self.day_vars[day] = {
                "status": vs, "start1": v_s1, "end1": v_e1, "start2": v_s2, "end2": v_e2
            }

        # Summary columns
        self.lbl_ist = ttk.Label(self.frame, text="00:00", width=8, anchor="center"); self.lbl_ist.grid(row=0, column=col); col+=1
        self.lbl_soll = ttk.Label(self.frame, text=minutes_to_hhmm(emp.weekly_target_min, False), width=8, anchor="center"); self.lbl_soll.grid(row=0, column=col); col+=1
        self.lbl_diff = ttk.Label(self.frame, text="00:00", width=8, anchor="center"); self.lbl_diff.grid(row=0, column=col); col+=1

        ttk.Entry(self.frame, width=8, textvariable=self.vars["carry"]).grid(row=0, column=col); col+=1

        # initial calc
        self.recalculate_labels()

        # update on edit
        for day in self.days:
            for key in ("status","start1","end1","start2","end2"):
                var = self.day_vars[day][key]
                var.trace_add("write", lambda *a: self.recalculate_labels())

    def collect_week_data(self) -> WeekEmployeeData:
        wed = WeekEmployeeData()
        try:
            wed.carry_prev_min = hhmm_to_minutes(self.vars["carry"].get())
        except ValueError:
            wed.carry_prev_min = 0
        wed.days = {}
        for day in self.days:
            dv = self.day_vars[day]
            wed.days[day] = DayEntry(
                status=dv["status"].get(),
                start1=dv["start1"].get(), end1=dv["end1"].get(),
                start2=dv["start2"].get(), end2=dv["end2"].get(),
            )
        return wed

    def recalculate_labels(self):
        settings = self.app.data["settings"]
        wed = self.collect_week_data()
        ist = wed.weekly_ist(settings["pause_short_min"], settings["pause_long_min"], settings["pause_threshold_min"])
        soll = self.emp.weekly_target_min
        diff = ist - soll
        self.lbl_ist.configure(text=minutes_to_hhmm(ist, False))
        self.lbl_diff.configure(text=minutes_to_hhmm(diff, True))
# ---------- Settings Page ----------

class SettingsPage(ttk.Frame):
    def __init__(self, parent, app: FlohkisteApp):
        super().__init__(parent)
        self.app = app

        frm = ttk.LabelFrame(self, text="Pausenregel & Optionen")
        frm.pack(fill="x", padx=10, pady=10)

        s = self.app.data["settings"]
        self.var_threshold = tk.IntVar(value=int(s.get("pause_threshold_min", 555)))  # 9:15 = 555 Min.
        self.var_short = tk.IntVar(value=int(s.get("pause_short_min", 30)))
        self.var_long = tk.IntVar(value=int(s.get("pause_long_min", 45)))
        self.var_vaca = tk.BooleanVar(value=bool(s.get("count_vacation_as_work", False)))

        row = 0
        ttk.Label(frm, text="Pausenschwelle (Minuten, Standard 555 = 9:15):").grid(row=row, column=0, sticky="w", padx=6, pady=6)
        ttk.Entry(frm, textvariable=self.var_threshold, width=8).grid(row=row, column=1, sticky="w", padx=6, pady=6); row+=1

        ttk.Label(frm, text="Pause kurz (Minuten, bis Schwelle inkl.):").grid(row=row, column=0, sticky="w", padx=6, pady=6)
        ttk.Entry(frm, textvariable=self.var_short, width=8).grid(row=row, column=1, sticky="w", padx=6, pady=6); row+=1

        ttk.Label(frm, text="Pause lang (Minuten, über Schwelle):").grid(row=row, column=0, sticky="w", padx=6, pady=6)
        ttk.Entry(frm, textvariable=self.var_long, width=8).grid(row=row, column=1, sticky="w", padx=6, pady=6); row+=1

        ttk.Checkbutton(frm, text="Urlaub/Krank/Feiertag als Arbeitszeit-Soll zählen (optional)", variable=self.var_vaca).grid(row=row, column=0, columnspan=2, sticky="w", padx=6, pady=6); row+=1

        ttk.Button(self, text="Einstellungen speichern", command=self.save).pack(anchor="e", padx=10, pady=(0,10))

        hint = ttk.Label(self, text="Hinweis: Änderungen wirken sich auf neue Berechnungen aus. Bereits gespeicherte Wochen werden beim erneuten Öffnen neu berechnet.", foreground="#555")
        hint.pack(anchor="w", padx=12)

    def save(self):
        s = self.app.data["settings"]
        s["pause_threshold_min"] = max(0, int(self.var_threshold.get()))
        s["pause_short_min"] = max(0, int(self.var_short.get()))
        s["pause_long_min"] = max(0, int(self.var_long.get()))
        s["count_vacation_as_work"] = bool(self.var_vaca.get())
        save_data(self.app.data)
        messagebox.showinfo("Gespeichert", "Einstellungen gespeichert.")

# ---------- Export Page ----------

class ExportPage(ttk.Frame):
    def __init__(self, parent, app: FlohkisteApp):
        super().__init__(parent)
        self.app = app

        frm = ttk.LabelFrame(self, text="PDF-Export (A4 Querformat)")
        frm.pack(fill="x", padx=10, pady=10)

        self.var_year = tk.IntVar(value=datetime.now().year)
        self.var_kw = tk.IntVar(value=datetime.now().isocalendar().week)

        row=0
        ttk.Label(frm, text="Jahr:").grid(row=row, column=0, sticky="w", padx=6, pady=6)
        ttk.Entry(frm, textvariable=self.var_year, width=6).grid(row=row, column=1, sticky="w", padx=6, pady=6); row+=1

        ttk.Label(frm, text="Kalenderwoche (KW):").grid(row=row, column=0, sticky="w", padx=6, pady=6)
        ttk.Entry(frm, textvariable=self.var_kw, width=6).grid(row=row, column=1, sticky="w", padx=6, pady=6); row+=1

        ttk.Button(frm, text="PDF erzeugen…", command=self.export_pdf).grid(row=row, column=0, columnspan=2, sticky="ew", padx=6, pady=8)

        self.hint = ttk.Label(self, text="Tipp: Speichere zuerst die Woche auf dem 'Woche'-Tab, bevor du exportierst.", foreground="#555")
        self.hint.pack(anchor="w", padx=12, pady=(0,8))

    def export_pdf(self):
        try:
            from reportlab.lib.pagesizes import A4, landscape
            from reportlab.pdfgen import canvas
            from reportlab.lib.units import mm
        except Exception:
            messagebox.showerror("ReportLab fehlt", "Für den PDF-Export wird das Paket 'reportlab' benötigt.\nInstalliere es mit:\n\npip install reportlab")
            return

        year = int(self.var_year.get())
        kw = int(self.var_kw.get())

        # Datei auswählen
        default_name = f"Dienstplan_Flohkiste_KW{kw}_{year}.pdf"
        path = filedialog.asksaveasfilename(title="PDF speichern", defaultextension=".pdf",
                                            filetypes=[("PDF","*.pdf")], initialfile=default_name)
        if not path:
            return

        employees = [e for e in self.app.employees() if e.active]
        if not employees:
            messagebox.showwarning("Keine aktiven Mitarbeiter", "Es sind keine aktiven Mitarbeiter vorhanden.")
            return

        s = self.app.data["settings"]
        headers = ["Mitarbeiter","Montag","Dienstag","Mittwoch","Donnerstag","Freitag","Woche Ist","Woche Soll","Diff","Übertrag neu"]
        rows = []

        for emp in employees:
            wed = self.app.get_week_data(emp.id, year, kw)
            ist = wed.weekly_ist(s["pause_short_min"], s["pause_long_min"], s["pause_threshold_min"])
            soll = emp.weekly_target_min
            diff = ist - soll
            new_carry = wed.carry_prev_min + diff
            day_values = []
            for day in ["Montag","Dienstag","Mittwoch","Donnerstag","Freitag"]:
                dm = wed.days[day].net_minutes_with_pause(s["pause_short_min"], s["pause_long_min"], s["pause_threshold_min"])
                day_values.append(minutes_to_hhmm(dm, False))
            row = [emp.name] + day_values + [
                minutes_to_hhmm(ist, False),
                minutes_to_hhmm(soll, False),
                minutes_to_hhmm(diff, True),
                minutes_to_hhmm(new_carry, True),
            ]
            rows.append(row)

        # PDF zeichnen
        pagesize = landscape(A4)
        c = canvas.Canvas(path, pagesize=pagesize)
        width, height = pagesize

        title = f"Dienstplan Flohkiste – KW {kw}, {year}"
        c.setFont("Helvetica-Bold", 16)
        c.drawString(20*mm, height-15*mm, title)

        left = 10*mm
        top = height - 25*mm
        col_widths = [40*mm, 25*mm,25*mm,25*mm,25*mm,25*mm, 25*mm,25*mm,25*mm,30*mm]

        # Überschriften
        c.setFont("Helvetica-Bold", 9)
        x = left
        for i, h in enumerate(headers):
            c.drawString(x+2, top, h)
            x += col_widths[i]

        # Zeilen
        c.setFont("Helvetica", 9)
        y = top - 6*mm
        row_height = 8*mm

        for row in rows:
            x = left
            for i, val in enumerate(row):
                c.drawString(x+2, y, str(val))
                x += col_widths[i]
            y -= row_height
            if y < 15*mm:
                c.showPage()
                c.setFont("Helvetica-Bold", 16)
                c.drawString(20*mm, height-15*mm, title)
                c.setFont("Helvetica-Bold", 9)
                x = left
                for i, h in enumerate(headers):
                    c.drawString(x+2, height-25*mm, h)
                c.setFont("Helvetica", 9)
                y = height - 31*mm

        c.save()
        messagebox.showinfo("PDF erstellt", f"PDF gespeichert:\n{path}")

# ---------- main ----------

def main():
    app = FlohkisteApp()
    app.mainloop()

if __name__ == "__main__":
    main()
