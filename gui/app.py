"""gui/app.py — Main AttendTrack GUI window built with ttkbootstrap."""

import threading
import tkinter as tk
from tkinter import filedialog, messagebox
from pathlib import Path
from datetime import datetime
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
try:
    from ttkbootstrap.scrolled import ScrolledFrame
    _HAS_SCROLLED = True
except ImportError:
    _HAS_SCROLLED = False

from core import config, reader, calculator, reporter
from utils import logger as log_mod

# ── MACROS ────────────────────────────────────────────────────────────────────
SIDEBAR_W       = 220
CMD_WIN_H       = 220
CARD_PAD        = 12
ROW_HEIGHT      = 26
KPI_CARD_W      = 180
FONT_TITLE      = ("Segoe UI", 22, "bold")
FONT_SUBTITLE   = ("Segoe UI", 10)
FONT_LABEL      = ("Segoe UI", 9)
FONT_CARD_NUM   = ("Segoe UI", 28, "bold")
FONT_CARD_LBL   = ("Segoe UI", 8, "bold")
FONT_CMD        = ("Consolas", 9)
FONT_BTN        = ("Segoe UI", 9, "bold")
# ──────────────────────────────────────────────────────────────────────────────


class AttendTrackApp:
    def __init__(self):
        self.settings   = config.load()
        self.log        = log_mod.setup(
            self.settings["app"]["name"],
            Path("attendtrack.log")
        )
        self._df_raw    = None
        self._df_comp   = None
        self._df_emp    = None
        self._df_dept   = None
        self._kpis      = {}
        self._file_path = self.settings["app"].get("last_file", "")

        theme = self.settings["app"].get("theme", "darkly")
        self.root = ttk.Window(themename=theme)
        self._build_window()
        self._build_sidebar()
        self._build_main()
        self._build_command_window()
        self.log.add_callback(self._cmd_log)
        self.log.info(f"{self.settings['app']['name']} ready.")

        if self._file_path and Path(self._file_path).exists():
            self._load_file(self._file_path)

    # ── Window ────────────────────────────────────────────────────────────────

    def _build_window(self):
        s = self.settings["app"]
        self.root.title(f"  {s['name']}  —  Attendance Tracker  v{s['version']}")
        w, h = s["window_width"], s["window_height"]
        self.root.geometry(f"{w}x{h}")
        self.root.minsize(900, 600)

    # ── Sidebar ───────────────────────────────────────────────────────────────

    def _build_sidebar(self):
        sb = ttk.Frame(self.root, width=SIDEBAR_W, bootstyle="dark")
        sb.pack(side=LEFT, fill=Y)
        sb.pack_propagate(False)

        # Logo / title area
        logo_frame = ttk.Frame(sb, bootstyle="dark")
        logo_frame.pack(fill=X, pady=(24, 0))

        ttk.Label(
            logo_frame, text="📋", font=("Segoe UI", 28),
            bootstyle="inverse-dark"
        ).pack()
        ttk.Label(
            logo_frame,
            text=self.settings["app"]["name"],
            font=("Segoe UI", 14, "bold"),
            bootstyle="inverse-dark"
        ).pack(pady=(4, 2))
        ttk.Label(
            logo_frame, text="Attendance Tracker",
            font=FONT_SUBTITLE, bootstyle="inverse-dark"
        ).pack()

        ttk.Separator(sb, bootstyle="secondary").pack(fill=X, pady=20, padx=16)

        # Nav buttons
        nav_items = [
            ("🏠  Dashboard",      self._show_dashboard),
            ("👥  Employee View",  self._show_employees),
            ("🏢  Departments",    self._show_departments),
            ("📄  Detail Records", self._show_detail),
        ]
        self._nav_btns = []
        for label, cmd in nav_items:
            btn = ttk.Button(
                sb, text=label, command=cmd,
                bootstyle="secondary-link",
                width=22,
            )
            btn.pack(fill=X, padx=10, pady=2)
            self._nav_btns.append(btn)

        ttk.Separator(sb, bootstyle="secondary").pack(fill=X, pady=16, padx=16)

        # File section
        ttk.Label(sb, text="FILE", font=("Segoe UI", 7, "bold"),
                  bootstyle="secondary-inverse").pack(padx=16, anchor=W)

        ttk.Button(
            sb, text="📂  Open Excel File",
            command=self._browse_file,
            bootstyle="info-outline",
            width=22,
        ).pack(fill=X, padx=10, pady=(4, 2))

        ttk.Button(
            sb, text="📊  Generate Report",
            command=self._generate_report,
            bootstyle="success-outline",
            width=22,
        ).pack(fill=X, padx=10, pady=2)

        ttk.Button(
            sb, text="📝  Download Template",
            command=self._download_template,
            bootstyle="warning-outline",
            width=22,
        ).pack(fill=X, padx=10, pady=2)

        ttk.Separator(sb, bootstyle="secondary").pack(fill=X, pady=16, padx=16)

        ttk.Button(
            sb, text="⚙️  Settings",
            command=self._open_settings,
            bootstyle="secondary-link",
            width=22,
        ).pack(fill=X, padx=10, pady=2)

        ttk.Button(
            sb, text="⌨️  Command Window",
            command=self._toggle_cmd,
            bootstyle="secondary-link",
            width=22,
        ).pack(fill=X, padx=10, pady=2)

        # File label at bottom
        self._file_lbl_var = tk.StringVar(value="No file loaded")
        ttk.Label(
            sb, textvariable=self._file_lbl_var,
            font=("Segoe UI", 7), bootstyle="secondary-inverse",
            wraplength=SIDEBAR_W - 20
        ).pack(side=BOTTOM, padx=10, pady=10)

    # ── Main content area ─────────────────────────────────────────────────────

    def _build_main(self):
        self._main = ttk.Frame(self.root)
        self._main.pack(side=LEFT, fill=BOTH, expand=True)

        # Top bar
        topbar = ttk.Frame(self._main, height=52)
        topbar.pack(fill=X)
        topbar.pack_propagate(False)

        self._page_title_var = tk.StringVar(value="Dashboard")
        ttk.Label(
            topbar, textvariable=self._page_title_var,
            font=FONT_TITLE
        ).pack(side=LEFT, padx=24, pady=8)

        self._status_var = tk.StringVar(value="Ready")
        self._status_lbl = ttk.Label(
            topbar, textvariable=self._status_var,
            font=FONT_LABEL, bootstyle="secondary"
        )
        self._status_lbl.pack(side=RIGHT, padx=24)

        ttk.Separator(self._main).pack(fill=X)

        # Content notebook (pages as frames, only one visible at a time)
        self._content = ttk.Frame(self._main)
        self._content.pack(fill=BOTH, expand=True)

        self._pages = {}
        for name in ("dashboard", "employees", "departments", "detail"):
            frame = ttk.Frame(self._content)
            frame.place(relx=0, rely=0, relwidth=1, relheight=1)
            self._pages[name] = frame

        self._build_dashboard_page()
        self._build_employees_page()
        self._build_departments_page()
        self._build_detail_page()
        self._show_dashboard()

    # ── Dashboard Page ────────────────────────────────────────────────────────

    def _build_dashboard_page(self):
        p = self._pages["dashboard"]

        # KPI cards row
        self._kpi_frame = ttk.Frame(p)
        self._kpi_frame.pack(fill=X, padx=20, pady=(20, 10))

        self._kpi_vars = {}
        kpi_defs = [
            ("attendance_rate",   "Attendance",   "primary"),
            ("punctuality_rate",  "Punctuality",  "success"),
            ("total_employees",   "Employees",    "info"),
            ("total_absent",      "Absences",     "danger"),
            ("total_late",        "Late Arrivals","warning"),
            ("avg_hours",         "Avg Hrs/Day",  "secondary"),
            ("total_overtime",    "Overtime Hrs", "dark"),
        ]
        for key, label, style in kpi_defs:
            card = ttk.Frame(self._kpi_frame, bootstyle=style)
            card.pack(side=LEFT, padx=5, pady=4, ipadx=CARD_PAD, ipady=CARD_PAD)

            var = tk.StringVar(value="—")
            self._kpi_vars[key] = var
            ttk.Label(card, text=label.upper(), font=FONT_CARD_LBL,
                      bootstyle=f"inverse-{style}").pack()
            ttk.Label(card, textvariable=var, font=FONT_CARD_NUM,
                      bootstyle=f"inverse-{style}").pack()

        # Charts row + summary table
        bottom = ttk.Frame(p)
        bottom.pack(fill=BOTH, expand=True, padx=20, pady=10)

        # Attendance by department table (left)
        left_panel = ttk.LabelFrame(bottom, text="Department Overview")
        left_panel.pack(side=LEFT, fill=BOTH, expand=True, padx=(0, 8))

        cols = ("Department", "Employees", "Avg Hours", "Absences", "Late")
        self._dept_tree = ttk.Treeview(
            left_panel, columns=cols, show="headings",
            height=8, bootstyle="primary"
        )
        for col in cols:
            self._dept_tree.heading(col, text=col)
            self._dept_tree.column(col, width=90, anchor=CENTER)
        self._dept_tree.pack(fill=BOTH, expand=True)
        ttk.Scrollbar(left_panel, orient=VERTICAL,
                      command=self._dept_tree.yview).pack(side=RIGHT, fill=Y)

        # Top lates / absentees (right)
        right_panel = ttk.LabelFrame(bottom, text="Notable Records")
        right_panel.pack(side=LEFT, fill=BOTH, expand=True, padx=(8, 0))

        self._notable_text = tk.Text(
            right_panel, wrap=WORD, state=DISABLED,
            font=FONT_CMD, relief="flat", height=10
        )
        self._notable_text.pack(fill=BOTH, expand=True)

    # ── Employees Page ────────────────────────────────────────────────────────

    def _build_employees_page(self):
        p = self._pages["employees"]

        # Filter bar
        fb = ttk.Frame(p)
        fb.pack(fill=X, padx=20, pady=12)
        ttk.Label(fb, text="Filter:", font=FONT_LABEL).pack(side=LEFT)
        self._emp_filter = tk.StringVar()
        self._emp_filter.trace_add("write", lambda *_: self._refresh_emp_table())
        ttk.Entry(fb, textvariable=self._emp_filter, width=30).pack(side=LEFT, padx=8)
        ttk.Button(fb, text="Clear", command=lambda: self._emp_filter.set(""),
                   bootstyle="secondary-outline").pack(side=LEFT, padx=4, pady=3)

        # Table
        frame = ttk.Frame(p)
        frame.pack(fill=BOTH, expand=True, padx=20, pady=(0, 16))

        emp_cols = ("Employee", "Department", "Present", "Absent",
                    "Leave", "Late", "Total Hrs", "Avg Hrs", "Overtime", "Avg Late (min)")
        self._emp_tree = ttk.Treeview(
            frame, columns=emp_cols, show="headings",
            height=20, bootstyle="primary"
        )
        for col in emp_cols:
            self._emp_tree.heading(col, text=col,
                                   command=lambda c=col: self._sort_emp(c))
            self._emp_tree.column(col, width=100, anchor=CENTER)
        self._emp_tree.column("Employee", width=160, anchor=W)

        vsb = ttk.Scrollbar(frame, orient=VERTICAL, command=self._emp_tree.yview)
        hsb = ttk.Scrollbar(frame, orient=HORIZONTAL, command=self._emp_tree.xview)
        self._emp_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        vsb.pack(side=RIGHT, fill=Y)
        hsb.pack(side=BOTTOM, fill=X)
        self._emp_tree.pack(fill=BOTH, expand=True)

        self._emp_sort_col  = None
        self._emp_sort_rev  = False

    # ── Departments Page ──────────────────────────────────────────────────────

    def _build_departments_page(self):
        p = self._pages["departments"]

        frame = ttk.Frame(p)
        frame.pack(fill=BOTH, expand=True, padx=20, pady=16)

        dept_cols = ("Department", "Headcount", "Avg Hours/Day",
                     "Total Absences", "Total Late", "Total Overtime")
        self._dept_detail_tree = ttk.Treeview(
            frame, columns=dept_cols, show="headings",
            height=20, bootstyle="info"
        )
        for col in dept_cols:
            self._dept_detail_tree.heading(col, text=col)
            self._dept_detail_tree.column(col, width=140, anchor=CENTER)
        self._dept_detail_tree.column("Department", width=180, anchor=W)

        vsb = ttk.Scrollbar(frame, orient=VERTICAL,
                             command=self._dept_detail_tree.yview)
        vsb.pack(side=RIGHT, fill=Y)
        self._dept_detail_tree.pack(fill=BOTH, expand=True)

    # ── Detail Page ───────────────────────────────────────────────────────────

    def _build_detail_page(self):
        p = self._pages["detail"]

        # Filter bar
        fb = ttk.Frame(p)
        fb.pack(fill=X, padx=20, pady=12)
        ttk.Label(fb, text="Search:", font=FONT_LABEL).pack(side=LEFT)
        self._detail_filter = tk.StringVar()
        self._detail_filter.trace_add("write", lambda *_: self._refresh_detail_table())
        ttk.Entry(fb, textvariable=self._detail_filter, width=30).pack(side=LEFT, padx=8)

        ttk.Label(fb, text="Show only:", font=FONT_LABEL).pack(side=LEFT, padx=(20, 4))
        self._detail_show = tk.StringVar(value="All")
        cb = ttk.Combobox(
            fb, textvariable=self._detail_show, width=12,
            values=["All", "Late", "Absent", "Leave", "Present"]
        )
        cb.pack(side=LEFT)
        cb.bind("<<ComboboxSelected>>", lambda _: self._refresh_detail_table())

        # Table
        frame = ttk.Frame(p)
        frame.pack(fill=BOTH, expand=True, padx=20, pady=(0, 16))

        det_cols = ("Employee", "Department", "Date", "Time In", "Time Out",
                    "Status", "Hours", "Late?", "Late (min)", "OT Hrs")
        self._detail_tree = ttk.Treeview(
            frame, columns=det_cols, show="headings",
            height=20, bootstyle="primary"
        )
        for col in det_cols:
            self._detail_tree.heading(col, text=col)
            self._detail_tree.column(col, width=95, anchor=CENTER)
        self._detail_tree.column("Employee", width=150, anchor=W)
        self._detail_tree.tag_configure("late",   background="#3a2020")
        self._detail_tree.tag_configure("absent", background="#3a2e20")
        self._detail_tree.tag_configure("leave",  background="#1e2e3a")

        vsb = ttk.Scrollbar(frame, orient=VERTICAL, command=self._detail_tree.yview)
        hsb = ttk.Scrollbar(frame, orient=HORIZONTAL, command=self._detail_tree.xview)
        self._detail_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        vsb.pack(side=RIGHT, fill=Y)
        hsb.pack(side=BOTTOM, fill=X)
        self._detail_tree.pack(fill=BOTH, expand=True)

    # ── Command Window ────────────────────────────────────────────────────────

    def _build_command_window(self):
        self._cmd_win = ttk.Frame(self._main)

        ttk.Separator(self._cmd_win).pack(fill=X)

        topbar = ttk.Frame(self._cmd_win, height=28)
        topbar.pack(fill=X)
        topbar.pack_propagate(False)
        ttk.Label(topbar, text="⌨️  Command Window",
                  font=("Segoe UI", 8, "bold"),
                  bootstyle="secondary").pack(side=LEFT, padx=10)
        ttk.Button(topbar, text="✕", command=self._toggle_cmd,
                   bootstyle="secondary-link").pack(side=RIGHT, padx=6)
        ttk.Button(topbar, text="Clear",
                   command=self._cmd_clear,
                   bootstyle="secondary-link").pack(side=RIGHT)

        # Output log
        log_frame = ttk.Frame(self._cmd_win)
        log_frame.pack(fill=BOTH, expand=True, padx=8, pady=(4, 0))

        self._cmd_output = tk.Text(
            log_frame, height=7, state=DISABLED,
            font=FONT_CMD, relief="flat",
            wrap=WORD, bg="#1a1a2e", fg="#00ff88",
            insertbackground="#00ff88"
        )
        vsb = ttk.Scrollbar(log_frame, command=self._cmd_output.yview)
        self._cmd_output.configure(yscrollcommand=vsb.set)
        vsb.pack(side=RIGHT, fill=Y)
        self._cmd_output.pack(fill=BOTH, expand=True)

        # Input bar
        input_frame = ttk.Frame(self._cmd_win)
        input_frame.pack(fill=X, padx=8, pady=6)

        ttk.Label(input_frame, text=">", font=FONT_CMD,
                  foreground="#00ff88").pack(side=LEFT)
        self._cmd_entry = ttk.Entry(input_frame, font=FONT_CMD)
        self._cmd_entry.pack(side=LEFT, fill=X, expand=True, padx=(4, 4))
        self._cmd_entry.bind("<Return>", self._cmd_execute)

        ttk.Button(input_frame, text="Run",
                   command=self._cmd_execute,
                   bootstyle="success-outline").pack(side=LEFT, padx=(4, 0))

        self._cmd_visible = False

    # ── Page switcher ─────────────────────────────────────────────────────────

    def _show_page(self, name: str, title: str):
        for p in self._pages.values():
            p.lower()
        self._pages[name].lift()
        self._page_title_var.set(title)

    def _show_dashboard(self):   self._show_page("dashboard",    "Dashboard")
    def _show_employees(self):   self._show_page("employees",    "Employee View")
    def _show_departments(self): self._show_page("departments",  "Departments")
    def _show_detail(self):      self._show_page("detail",       "Attendance Detail")

    # ── File operations ───────────────────────────────────────────────────────

    def _browse_file(self):
        path = filedialog.askopenfilename(
            title="Open Attendance Excel File",
            filetypes=[("Excel Files", "*.xlsx *.xls"), ("All Files", "*.*")]
        )
        if path:
            self._load_file(path)

    def _load_file(self, path: str):
        self._status_var.set("Loading…")
        self.log.info(f"Loading file: {path}")

        def worker():
            try:
                df_raw, warnings = reader.load_file(path, self.settings)
                for w in warnings:
                    self.log.warning(w)
                issues = reader.validate_dataframe(df_raw)
                for i in issues:
                    self.log.warning(i)

                df_comp  = calculator.compute_records(df_raw, self.settings)
                df_emp   = calculator.summarize_by_employee(df_comp)
                df_dept  = calculator.summarize_by_department(df_comp)
                kpis     = calculator.overall_kpis(df_comp)

                self.root.after(0, lambda: self._on_load_success(
                    path, df_raw, df_comp, df_emp, df_dept, kpis
                ))
            except Exception as e:
                self.root.after(0, lambda: self._on_load_error(str(e)))

        threading.Thread(target=worker, daemon=True).start()

    def _on_load_success(self, path, df_raw, df_comp, df_emp, df_dept, kpis):
        self._file_path = path
        self._df_raw    = df_raw
        self._df_comp   = df_comp
        self._df_emp    = df_emp
        self._df_dept   = df_dept
        self._kpis      = kpis

        fname = Path(path).name
        self._file_lbl_var.set(fname)
        self._status_var.set(f"Loaded: {len(df_raw)} records")
        self.log.success(f"File loaded — {len(df_raw)} records, "
                         f"{kpis.get('total_employees', 0)} employees.")

        # Save last file
        if self.settings["app"].get("remember_last_file", True):
            self.settings["app"]["last_file"] = path
            config.save(self.settings)

        self._refresh_all()

    def _on_load_error(self, msg: str):
        self._status_var.set("Error loading file.")
        self.log.error(msg)
        messagebox.showerror("Load Error", msg)

    # ── Data refresh ──────────────────────────────────────────────────────────

    def _refresh_all(self):
        self._refresh_kpis()
        self._refresh_emp_table()
        self._refresh_dept_trees()
        self._refresh_detail_table()
        self._refresh_notable()

    def _refresh_kpis(self):
        for key, var in self._kpi_vars.items():
            val = self._kpis.get(key, "—")
            if isinstance(val, float):
                var.set(f"{val:.1f}")
            else:
                var.set(str(val))

    def _refresh_emp_table(self):
        self._emp_tree.delete(*self._emp_tree.get_children())
        if self._df_emp is None or self._df_emp.empty:
            return
        filt = self._emp_filter.get().lower()
        for _, row in self._df_emp.iterrows():
            vals = (
                row.get("Employee", ""),
                row.get("Department", ""),
                row.get("Days Present", ""),
                row.get("Days Absent", ""),
                row.get("Days Leave", ""),
                row.get("Days Late", ""),
                row.get("Total Hours", ""),
                row.get("Avg Hours/Day", ""),
                row.get("Overtime Hours", ""),
                row.get("Avg Late (min)", ""),
            )
            name = str(vals[0]).lower()
            dept = str(vals[1]).lower()
            if filt and filt not in name and filt not in dept:
                continue
            self._emp_tree.insert("", END, values=vals)

    def _sort_emp(self, col: str):
        mapping = {
            "Employee": "Employee", "Department": "Department",
            "Present": "Days Present", "Absent": "Days Absent",
            "Leave": "Days Leave", "Late": "Days Late",
            "Total Hrs": "Total Hours", "Avg Hrs": "Avg Hours/Day",
            "Overtime": "Overtime Hours", "Avg Late (min)": "Avg Late (min)",
        }
        df_col = mapping.get(col, col)
        if self._df_emp is None or df_col not in self._df_emp.columns:
            return
        self._emp_sort_rev = not self._emp_sort_rev if self._emp_sort_col == col else False
        self._emp_sort_col = col
        self._df_emp = self._df_emp.sort_values(
            df_col, ascending=not self._emp_sort_rev
        ).reset_index(drop=True)
        self._refresh_emp_table()

    def _refresh_dept_trees(self):
        # Dashboard dept tree
        for tree in (self._dept_tree, self._dept_detail_tree):
            tree.delete(*tree.get_children())

        if self._df_dept is None or self._df_dept.empty:
            return

        for _, row in self._df_dept.iterrows():
            self._dept_tree.insert("", END, values=(
                row.get("Department", ""),
                row.get("Headcount", ""),
                row.get("Avg Hours/Day", ""),
                row.get("Total Absences", ""),
                row.get("Total Late", ""),
            ))
            self._dept_detail_tree.insert("", END, values=(
                row.get("Department", ""),
                row.get("Headcount", ""),
                row.get("Avg Hours/Day", ""),
                row.get("Total Absences", ""),
                row.get("Total Late", ""),
                row.get("Total Overtime", ""),
            ))

    def _refresh_detail_table(self):
        self._detail_tree.delete(*self._detail_tree.get_children())
        if self._df_comp is None or self._df_comp.empty:
            return

        filt  = self._detail_filter.get().lower()
        show  = self._detail_show.get()

        for _, row in self._df_comp.iterrows():
            status  = str(row.get("status_normalized", "")).lower()
            is_late = row.get("is_late", False)
            emp     = str(row.get("employee_name", ""))

            # Filter by search
            if filt and filt not in emp.lower():
                continue

            # Filter by dropdown
            if show == "Late"    and not is_late:               continue
            if show == "Absent"  and status != "absent":        continue
            if show == "Leave"   and status != "leave":         continue
            if show == "Present" and status not in ("present",): continue

            # Format values
            ti = row.get("time_in")
            to = row.get("time_out")
            dt = row.get("date")
            try:  ti_s = ti.strftime("%H:%M") if ti else "—"
            except: ti_s = str(ti)
            try:  to_s = to.strftime("%H:%M") if to else "—"
            except: to_s = str(to)
            try:  dt_s = dt.strftime("%Y-%m-%d") if dt else "—"
            except: dt_s = str(dt)

            tag = "late" if is_late else ("absent" if status == "absent" else
                  ("leave" if status == "leave" else ""))

            self._detail_tree.insert("", END, tags=(tag,), values=(
                emp,
                row.get("department", "—"),
                dt_s, ti_s, to_s,
                status.title(),
                f"{row.get('hours_worked', 0):.1f}",
                "Yes" if is_late else "No",
                f"{row.get('late_minutes', 0):.0f}",
                f"{row.get('overtime_hours', 0):.1f}",
            ))

    def _refresh_notable(self):
        self._notable_text.configure(state=NORMAL)
        self._notable_text.delete("1.0", END)
        if self._df_comp is None or self._df_comp.empty:
            self._notable_text.insert(END, "Load a file to see records.")
            self._notable_text.configure(state=DISABLED)
            return

        lines = []
        # Top 5 most late employees
        if not self._df_emp.empty and "Days Late" in self._df_emp.columns:
            top_late = self._df_emp.nlargest(5, "Days Late")[
                ["Employee", "Days Late"]
            ]
            lines.append("── Most Late Arrivals ──────────────")
            for _, r in top_late.iterrows():
                lines.append(f"  {r['Employee']:<22} {int(r['Days Late'])} day(s)")

        if not self._df_emp.empty and "Days Absent" in self._df_emp.columns:
            top_absent = self._df_emp.nlargest(5, "Days Absent")[
                ["Employee", "Days Absent"]
            ]
            lines.append("\n── Most Absences ───────────────────")
            for _, r in top_absent.iterrows():
                lines.append(f"  {r['Employee']:<22} {int(r['Days Absent'])} day(s)")

        if not self._df_emp.empty and "Overtime Hours" in self._df_emp.columns:
            top_ot = self._df_emp.nlargest(5, "Overtime Hours")[
                ["Employee", "Overtime Hours"]
            ]
            lines.append("\n── Most Overtime ───────────────────")
            for _, r in top_ot.iterrows():
                lines.append(f"  {r['Employee']:<22} {r['Overtime Hours']:.1f} hrs")

        self._notable_text.insert(END, "\n".join(lines))
        self._notable_text.configure(state=DISABLED)

    # ── Report generation ─────────────────────────────────────────────────────

    def _generate_report(self):
        if self._df_comp is None:
            messagebox.showwarning("No Data", "Please load an attendance file first.")
            return

        folder = self.settings["report"].get("output_folder", "reports")
        Path(folder).mkdir(parents=True, exist_ok=True)
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        out = Path(folder) / f"AttendTrack_Report_{ts}.xlsx"

        self._status_var.set("Generating report…")
        self.log.info(f"Generating report → {out}")

        def worker():
            try:
                reporter.generate_report(
                    self._df_raw, self._df_comp,
                    self._df_emp, self._df_dept,
                    self._kpis, self.settings, out
                )
                self.root.after(0, lambda: self._on_report_done(out))
            except Exception as e:
                self.root.after(0, lambda: self._on_report_error(str(e)))

        threading.Thread(target=worker, daemon=True).start()

    def _on_report_done(self, path: Path):
        self._status_var.set(f"Report saved: {path.name}")
        self.log.success(f"Report saved: {path}")
        messagebox.showinfo("Report Ready",
                            f"Report saved to:\n{path}")

    def _on_report_error(self, msg: str):
        self._status_var.set("Report failed.")
        self.log.error(f"Report error: {msg}")
        messagebox.showerror("Report Error", msg)

    def _download_template(self):
        path = filedialog.asksaveasfilename(
            title="Save Sample Template",
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")],
            initialfile="attendance_template.xlsx"
        )
        if path:
            try:
                reporter.generate_sample_template(path, self.settings)
                self.log.success(f"Template saved: {path}")
                messagebox.showinfo("Template Saved", f"Template saved to:\n{path}")
            except Exception as e:
                messagebox.showerror("Error", str(e))

    # ── Settings window ───────────────────────────────────────────────────────

    def _open_settings(self):
        win = ttk.Toplevel(self.root)
        win.title("Settings")
        win.geometry("500x460")
        win.resizable(False, False)

        ttk.Label(win, text="Settings", font=("Segoe UI", 14, "bold")).pack(pady=16)
        ttk.Separator(win).pack(fill=X, padx=20)

        if _HAS_SCROLLED:
            sf = ScrolledFrame(win, autohide=True)
            sf.pack(fill=BOTH, expand=True, padx=20, pady=10)
        else:
            sf = ttk.Frame(win)
            sf.pack(fill=BOTH, expand=True, padx=20, pady=10)

        att = self.settings["attendance"]
        fields = [
            ("Work Start Time",          "work_start_time"),
            ("Work End Time",            "work_end_time"),
            ("Late Threshold (min)",     "late_threshold_minutes"),
            ("Half Day Hours",           "half_day_hours"),
            ("Full Day Hours",           "full_day_hours"),
            ("Overtime Threshold Hrs",   "overtime_threshold_hours"),
            ("Break Duration (min)",     "break_duration_minutes"),
        ]
        vars_ = {}
        for label, key in fields:
            row = ttk.Frame(sf)
            row.pack(fill=X, pady=4)
            ttk.Label(row, text=label, width=26, anchor=W,
                      font=FONT_LABEL).pack(side=LEFT)
            v = tk.StringVar(value=str(att.get(key, "")))
            vars_[key] = v
            ttk.Entry(row, textvariable=v, width=18).pack(side=LEFT)

        # Company name
        ttk.Separator(sf).pack(fill=X, pady=10)
        row = ttk.Frame(sf)
        row.pack(fill=X, pady=4)
        ttk.Label(row, text="Company Name", width=26, anchor=W,
                  font=FONT_LABEL).pack(side=LEFT)
        co_var = tk.StringVar(value=self.settings["report"].get("company_name", ""))
        ttk.Entry(row, textvariable=co_var, width=24).pack(side=LEFT)

        def save_settings():
            for key, var in vars_.items():
                try:
                    val = var.get()
                    orig = att[key]
                    att[key] = type(orig)(val)
                except Exception:
                    pass
            self.settings["report"]["company_name"] = co_var.get()
            config.save(self.settings)
            self.log.success("Settings saved.")
            win.destroy()

        ttk.Button(win, text="Save Settings",
                   command=save_settings,
                   bootstyle="success").pack(pady=12)

    # ── Command window ────────────────────────────────────────────────────────

    def _toggle_cmd(self):
        if self._cmd_visible:
            self._cmd_win.pack_forget()
            self._cmd_visible = False
        else:
            self._cmd_win.pack(side=BOTTOM, fill=X, in_=self._main)
            self._cmd_visible = True

    def _cmd_log(self, level: str, msg: str):
        """Callback from AppLogger to write to cmd window."""
        colors = {
            "OK":    "#00ff88",
            "INFO":  "#88ccff",
            "WARN":  "#ffcc44",
            "ERROR": "#ff5555",
            "DEBUG": "#888888",
        }
        color = colors.get(level, "#cccccc")
        ts = datetime.now().strftime("%H:%M:%S")

        self._cmd_output.configure(state=NORMAL)
        self._cmd_output.insert(END, f"[{ts}] ", "ts")
        self._cmd_output.insert(END, f"[{level}] ", level)
        self._cmd_output.insert(END, msg + "\n")
        self._cmd_output.tag_configure("ts",    foreground="#555577")
        self._cmd_output.tag_configure(level,   foreground=color)
        self._cmd_output.see(END)
        self._cmd_output.configure(state=DISABLED)

    def _cmd_clear(self):
        self._cmd_output.configure(state=NORMAL)
        self._cmd_output.delete("1.0", END)
        self._cmd_output.configure(state=DISABLED)

    def _cmd_execute(self, event=None):
        raw = self._cmd_entry.get().strip()
        self._cmd_entry.delete(0, END)
        if not raw:
            return

        self.log.info(f"> {raw}")
        cmd = raw.lower().split()
        verb = cmd[0] if cmd else ""

        if verb in ("help", "?"):
            cmds = [
                "help               — show this list",
                "load <path>        — load an Excel file",
                "report             — generate Excel report",
                "template <path>    — save sample template",
                "kpis               — print current KPIs",
                "clear              — clear this window",
                "status             — show loaded file info",
                "reload             — reload current file",
            ]
            for c in cmds:
                self.log.info(c)

        elif verb == "load":
            path = " ".join(cmd[1:])
            if not path:
                self.log.warning("Usage: load <path>")
            else:
                self._load_file(path)

        elif verb == "report":
            self._generate_report()

        elif verb == "template":
            path = " ".join(cmd[1:]) or "attendance_template.xlsx"
            try:
                reporter.generate_sample_template(path, self.settings)
                self.log.success(f"Template saved: {path}")
            except Exception as e:
                self.log.error(str(e))

        elif verb == "kpis":
            for k, v in self._kpis.items():
                self.log.info(f"  {k}: {v}")

        elif verb == "clear":
            self._cmd_clear()

        elif verb == "status":
            if self._file_path:
                self.log.info(f"File: {self._file_path}")
                self.log.info(f"Records: {self._kpis.get('total_records', 0)}")
                self.log.info(f"Employees: {self._kpis.get('total_employees', 0)}")
            else:
                self.log.warning("No file loaded.")

        elif verb == "reload":
            if self._file_path:
                self._load_file(self._file_path)
            else:
                self.log.warning("No file loaded.")

        else:
            self.log.warning(f"Unknown command: '{verb}' — type 'help' for commands.")

    # ── Run ───────────────────────────────────────────────────────────────────

    def run(self):
        self.root.mainloop()