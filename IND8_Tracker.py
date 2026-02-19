import os
import sys
import json
import sqlite3
import datetime
import csv
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from PIL import Image, ImageTk

# ----------------- APP META / VERSIONING -----------------

APP_NAME = "IND8 Tracker"
APP_VERSION = "1.0.0"

# ----------------- CONFIG / PATHS -----------------

CONFIG_FILE_NAME = "ind8_config.json"

def get_base_dir():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    else:
        return os.path.dirname(os.path.abspath(__file__))

def load_config():
    base_dir = get_base_dir()
    config_path = os.path.join(base_dir, CONFIG_FILE_NAME)
    default_config = {"cloud_sync": False}
    if not os.path.exists(config_path):
        save_config(default_config)
        return default_config
    try:
        with open(config_path, "r", encoding="utf-8") as f:
            data = json.load(f)
        for k, v in default_config.items():
            if k not in data:
                data[k] = v
        return data
    except Exception:
        return default_config

def save_config(config):
    base_dir = get_base_dir()
    config_path = os.path.join(base_dir, CONFIG_FILE_NAME)
    try:
        with open(config_path, "w", encoding="utf-8") as f:
            json.dump(config, f, indent=2)
    except Exception:
        pass

def get_db_path(config):
    cloud_sync = config.get("cloud_sync", False)
    if cloud_sync:
        one_drive = os.getenv("OneDrive")
        if one_drive:
            cloud_dir = os.path.join(one_drive, "IND8Tracker")
            os.makedirs(cloud_dir, exist_ok=True)
            return os.path.join(cloud_dir, "indirect_tracking.db")
    local_appdata = os.getenv("LOCALAPPDATA")
    if local_appdata:
        app_dir = os.path.join(local_appdata, "IND8Tracker")
        os.makedirs(app_dir, exist_ok=True)
        return os.path.join(app_dir, "indirect_tracking.db")
    return os.path.join(get_base_dir(), "indirect_tracking.db")

CONFIG = load_config()
DB_FILE = get_db_path(CONFIG)

WARNING_THRESHOLD_HOURS = 5
INDIRECT_LIMIT_HOURS = 6

# ----------------- DB SETUP -----------------

def init_db():
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("""
        CREATE TABLE IF NOT EXISTS associates (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            badge_id TEXT UNIQUE NOT NULL
        )
    """)
    c.execute("""
        CREATE TABLE IF NOT EXISTS sessions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            associate_id INTEGER NOT NULL,
            start_time TEXT NOT NULL,
            end_time TEXT,
            work_type TEXT NOT NULL,
            area TEXT NOT NULL,
            role TEXT NOT NULL,
            FOREIGN KEY (associate_id) REFERENCES associates(id)
        )
    """)
    conn.commit()
    conn.close()

def get_conn():
    return sqlite3.connect(DB_FILE)

def get_or_create_associate(badge_id, name=None):
    conn = get_conn()
    c = conn.cursor()
    c.execute("SELECT id, name FROM associates WHERE badge_id = ?", (badge_id,))
    row = c.fetchone()
    if row:
        conn.close()
        return row[0]
    if not name:
        name = badge_id
    c.execute("INSERT INTO associates (name, badge_id) VALUES (?, ?)", (name, badge_id))
    conn.commit()
    assoc_id = c.lastrowid
    conn.close()
    return assoc_id

def end_active_session(associate_id):
    conn = get_conn()
    c = conn.cursor()
    c.execute("SELECT id FROM sessions WHERE associate_id = ? AND end_time IS NULL", (associate_id,))
    row = c.fetchone()
    if row:
        now = datetime.datetime.now().isoformat()
        c.execute("UPDATE sessions SET end_time = ? WHERE id = ?", (now, row[0]))
        conn.commit()
    conn.close()

def start_session(associate_id, work_type, area, role):
    end_active_session(associate_id)
    conn = get_conn()
    c = conn.cursor()
    now = datetime.datetime.now().isoformat()
    c.execute("""
        INSERT INTO sessions (associate_id, start_time, work_type, area, role)
        VALUES (?, ?, ?, ?, ?)
    """, (associate_id, now, work_type, area, role))
    conn.commit()
    conn.close()

def get_today_sessions(associate_id):
    conn = get_conn()
    c = conn.cursor()
    today = datetime.date.today().isoformat()
    c.execute("""
        SELECT id, start_time, end_time, work_type, area, role
        FROM sessions
        WHERE associate_id = ?
          AND date(start_time) = ?
        ORDER BY start_time
    """, (associate_id, today))
    rows = c.fetchall()
    conn.close()
    return rows

def compute_indirect_hours_today(associate_id):
    sessions = get_today_sessions(associate_id)
    total_seconds = 0
    now = datetime.datetime.now()
    for _, start, end, work_type, _, _ in sessions:
        if work_type != "INDIRECT":
            continue
        start_dt = datetime.datetime.fromisoformat(start)
        end_dt = datetime.datetime.fromisoformat(end) if end else now
        total_seconds += (end_dt - start_dt).total_seconds()
    return total_seconds / 3600.0

def get_all_associates():
    conn = get_conn()
    c = conn.cursor()
    c.execute("SELECT id, name, badge_id FROM associates ORDER BY name")
    rows = c.fetchall()
    conn.close()
    return rows

# ----------------- SPLASH SCREEN -----------------

def show_splash(root):
    splash = tk.Toplevel(root)
    splash.overrideredirect(True)
    splash.configure(bg="#1a1a1a")
    try:
        img = Image.open("splash.png")
        splash_img = ImageTk.PhotoImage(img)
        label = tk.Label(splash, image=splash_img, bg="#1a1a1a")
        label.image = splash_img
        label.pack()
    except Exception:
        label = tk.Label(
            splash,
            text=f"{APP_NAME}\nv{APP_VERSION}",
            bg="#1a1a1a",
            fg="#ff9900",
            font=("Segoe UI", 20, "bold"),
            padx=40,
            pady=30
        )
        label.pack()
    splash.update_idletasks()
    w = splash.winfo_width()
    h = splash.winfo_height()
    x = (splash.winfo_screenwidth() - w) // 2
    y = (splash.winfo_screenheight() - h) // 2
    splash.geometry(f"{w}x{h}+{x}+{y}")
    splash.after(2000, splash.destroy)
    return splash

# ----------------- APP -----------------

class App:
    def __init__(self, root, config):
        self.root = root
        self.config = config
        self.root.title(f"{APP_NAME} v{APP_VERSION}")
        self.root.geometry("1100x650")
        self.root.configure(bg="#1a1a1a")

        self.current_associate_id = None

        self.badge_var = tk.StringVar()
        self.name_var = tk.StringVar()
        self.area_var = tk.StringVar(value="CRET")
        self.role_var = tk.StringVar(value="Water Spider")
        self.status_var = tk.StringVar(value="No associate selected")
        self.indirect_hours_var = tk.StringVar(value="0.00")
        self.cloud_sync_var = tk.BooleanVar(value=self.config.get("cloud_sync", False))

        self._setup_style()
        self._build_layout()

    # ---------- STYLE ----------

    def _setup_style(self):
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Dark.TFrame", background="#1a1a1a")
        style.configure("DarkInner.TFrame", background="#222222")
        style.configure("Dark.TLabel", background="#1a1a1a", foreground="#e6e6e6", font=("Segoe UI", 11))
        style.configure("DarkBanner.TLabel", background="#222222", foreground="#ff9900",
                        font=("Segoe UI", 22, "bold"), padding=20)
        style.configure("DarkNav.TButton", background="#333333", foreground="#e6e6e6",
                        padding=8, font=("Segoe UI", 11, "bold"))
        style.map("DarkNav.TButton", background=[("active", "#444444")])
        style.configure("Dark.Treeview",
                        background="#222222",
                        fieldbackground="#222222",
                        foreground="#e6e6e6",
                        rowheight=24)
        style.configure("Dark.Treeview.Heading",
                        background="#333333",
                        foreground="#e6e6e6",
                        font=("Segoe UI", 10, "bold"))

    # ---------- LAYOUT ----------

    def _build_layout(self):
        banner_frame = tk.Frame(self.root, bg="#222222", height=80)
        banner_frame.pack(side="top", fill="x")

        try:
            logo_img = Image.open("ind8logo.png").resize((60, 60))
            self.logo = ImageTk.PhotoImage(logo_img)
            logo_label = tk.Label(banner_frame, image=self.logo, bg="#222222")
            logo_label.pack(side="left", padx=20)
        except Exception:
            pass

        title_label = ttk.Label(banner_frame, text="IND8", style="DarkBanner.TLabel")
        title_label.pack(side="left")

        version_label = tk.Label(
            banner_frame,
            text=f"v{APP_VERSION}",
            bg="#222222",
            fg="#cccccc",
            font=("Segoe UI", 10, "bold")
        )
        version_label.pack(side="right", padx=20)

        main_frame = tk.Frame(self.root, bg="#1a1a1a")
        main_frame.pack(fill="both", expand=True)

        nav = tk.Frame(main_frame, bg="#111111", width=200)
        nav.pack(side="left", fill="y")

        self._add_nav_button(nav, "Home", lambda: None)
        self._add_nav_button(nav, "Start Direct", self.start_direct)
        self._add_nav_button(nav, "Start Indirect", self.start_indirect)
        self._add_nav_button(nav, "End Current", self.end_current)
        self._add_nav_button(nav, "Dashboard", self.open_dashboard)
        self._add_nav_button(nav, "Export to CSV", self.export_to_excel)
        self._add_nav_button(nav, "Shift Report", self.export_shift_report)
        self._add_nav_button(nav, "Legend", self.open_legend)

        cloud_frame = tk.Frame(nav, bg="#111111")
        cloud_frame.pack(fill="x", pady=10, padx=8)
        cloud_cb = tk.Checkbutton(
            cloud_frame,
            text="Cloud sync (OneDrive)",
            variable=self.cloud_sync_var,
            bg="#111111",
            fg="#e6e6e6",
            selectcolor="#111111",
            activebackground="#111111",
            activeforeground="#e6e6e6",
            command=self.toggle_cloud_sync
        )
        cloud_cb.pack(anchor="w")

        content = tk.Frame(main_frame, bg="#1a1a1a")
        content.pack(side="left", fill="both", expand=True, padx=10, pady=10)

        top = tk.Frame(content, bg="#1a1a1a")
        top.pack(fill="x", pady=(0, 10))

        tk.Label(top, text="Badge ID:", bg="#1a1a1a", fg="#e6e6e6",
                 font=("Segoe UI", 10)).grid(row=0, column=0, sticky="w")
        self.badge_entry = tk.Entry(top, textvariable=self.badge_var, width=18,
                                    bg="#333333", fg="#e6e6e6", insertbackground="white",
                                    relief="flat")
        self.badge_entry.grid(row=0, column=1, sticky="w", padx=(5, 20))
        self.badge_entry.bind("<Return>", self.scan_badge)

        tk.Label(top, text="Name (optional):", bg="#1a1a1a", fg="#e6e6e6",
                 font=("Segoe UI", 10)).grid(row=0, column=2, sticky="w")
        self.name_entry = tk.Entry(top, textvariable=self.name_var, width=20,
                                   bg="#333333", fg="#e6e6e6", insertbackground="white",
                                   relief="flat")
        self.name_entry.grid(row=0, column=3, sticky="w", padx=(5, 20))

        tk.Label(top, text="Area:", bg="#1a1a1a", fg="#e6e6e6",
                 font=("Segoe UI", 10)).grid(row=1, column=0, sticky="w", pady=(8, 0))
        area_cb = ttk.Combobox(top, textvariable=self.area_var,
                               values=["CRET", "VRET", "TRANSHIP", "ILS"],
                               width=15, state="readonly")
        area_cb.grid(row=1, column=1, sticky="w", pady=(8, 0))

        tk.Label(top, text="Indirect Role:", bg="#1a1a1a", fg="#e6e6e6",
                 font=("Segoe UI", 10)).grid(row=1, column=2, sticky="w", pady=(8, 0))
        role_cb = ttk.Combobox(top, textvariable=self.role_var,
                               values=["Water Spider", "Down Stack", "Unloads", "N/A"],
                               width=18, state="readonly")
        role_cb.grid(row=1, column=3, sticky="w", pady=(8, 0))

        load_btn = ttk.Button(top, text="Load Associate", style="DarkNav.TButton",
                              command=self.load_associate)
        load_btn.grid(row=0, column=4, padx=(10, 0))

        mid = tk.Frame(content, bg="#1a1a1a")
        mid.pack(fill="both", expand=True)

        columns = ("start", "end", "type", "area", "role", "duration")
        self.tree = ttk.Treeview(mid, columns=columns, show="headings",
                                 style="Dark.Treeview")
        for col in columns:
            self.tree.heading(col, text=col.capitalize(), anchor="w")
            self.tree.column(col, anchor="w", width=100)
        self.tree.pack(fill="both", expand=True)

        bottom = tk.Frame(content, bg="#1a1a1a")
        bottom.pack(fill="x", pady=(10, 0))

        tk.Label(bottom, text="Indirect hours today:",
                 bg="#1a1a1a", fg="#e6e6e6",
                 font=("Segoe UI", 10)).grid(row=0, column=0, sticky="w")
        tk.Label(bottom, textvariable=self.indirect_hours_var,
                 bg="#1a1a1a", fg="#ffcc00",
                 font=("Segoe UI", 11, "bold")).grid(row=0, column=1, sticky="w", padx=(5, 20))

        tk.Label(bottom, text="Status:",
                 bg="#1a1a1a", fg="#e6e6e6",
                 font=("Segoe UI", 10)).grid(row=0, column=2, sticky="w")
        self.status_label = tk.Label(bottom, textvariable=self.status_var,
                                     bg="#1a1a1a", fg="#00cc66",
                                     font=("Segoe UI", 10, "bold"))
        self.status_label.grid(row=0, column=3, sticky="w", padx=(5, 0))

        db_label = tk.Label(
            bottom,
            text=f"DB: {DB_FILE}",
            bg="#1a1a1a",
            fg="#777777",
            font=("Segoe UI", 8),
            anchor="w",
            justify="left",
            wraplength=600
        )
        db_label.grid(row=1, column=0, columnspan=4, sticky="w", pady=(5, 0))

    # ---------- DIRECT / INDIRECT HELPERS ----------

    def has_direct_today(self, associate_id):
        sessions = get_today_sessions(associate_id)
        return any(work_type == "DIRECT" for _, _, _, work_type, _, _ in sessions)

    def has_indirect_today(self, associate_id):
        sessions = get_today_sessions(associate_id)
        return any(work_type == "INDIRECT" for _, _, _, work_type, _, _ in sessions)

    def get_indirect_roles_today(self, associate_id):
        sessions = get_today_sessions(associate_id)
        roles = set()
        for _, _, _, work_type, _, role in sessions:
            if work_type == "INDIRECT":
                roles.add(role)
        return roles

    # ---------- NAV BUTTON ----------

    def _add_nav_button(self, parent, text, command):
        frame = tk.Frame(parent, bg="#111111")
        frame.pack(fill="x", pady=4, padx=8)

        btn = tk.Label(frame, text=text, bg="#333333", fg="#e6e6e6",
                       font=("Segoe UI", 10, "bold"),
                       padx=12, pady=6)
        btn.pack(fill="x")
        btn.bind("<Button-1>", lambda e: command())
        btn.bind("<Enter>", lambda e: btn.config(bg="#444444"))
        btn.bind("<Leave>", lambda e: btn.config(bg="#333333"))

    # ---------- CLOUD SYNC ----------

    def toggle_cloud_sync(self):
        new_value = self.cloud_sync_var.get()
        self.config["cloud_sync"] = new_value
        save_config(self.config)
        messagebox.showinfo(
            "Cloud Sync",
            "Cloud sync setting saved.\n\n"
            "Please close and reopen the app for the change to take effect."
        )

    # ---------- LEGEND PANEL ----------

    def open_legend(self):
        legend = tk.Toplevel(self.root)
        legend.title("Dashboard Legend")
        legend.configure(bg="#1a1a1a")
        legend.geometry("350x260")

        items = [
            ("DIRECT only", "#003366"),
            ("INDIRECT only", "#665511"),
            ("DIRECT + INDIRECT", "#CC6600"),
            ("No work today", "#444444"),
            ("Warning (5+ hrs indirect)", "#665511"),
            ("Violation (6+ hrs indirect)", "#661111"),
        ]

        tk.Label(
            legend,
            text="Dashboard Color Legend",
            bg="#1a1a1a",
            fg="#ff9900",
            font=("Segoe UI", 14, "bold"),
            pady=10
        ).pack()

        for label, color in items:
            row = tk.Frame(legend, bg="#1a1a1a")
            row.pack(fill="x", pady=4, padx=10)

            swatch = tk.Label(row, bg=color, width=4, height=2)
            swatch.pack(side="left", padx=(0, 10))

            tk.Label(
                row,
                text=label,
                bg="#1a1a1a",
                fg="#e6e6e6",
                font=("Segoe UI", 11)
            ).pack(side="left")

    # ---------- BADGE / ASSOCIATE ----------

    def scan_badge(self, event=None):
        badge = self.badge_var.get().strip()
        if not badge:
            return
        assoc_id = get_or_create_associate(badge, self.name_var.get().strip() or None)
        self.current_associate_id = assoc_id
        self.refresh_view()
        self._update_direct_indirect_status(assoc_id)

    def load_associate(self):
        badge = self.badge_var.get().strip()
        if not badge:
            messagebox.showerror("Error", "Enter a badge ID.")
            return
        assoc_id = get_or_create_associate(badge, self.name_var.get().strip() or None)
        self.current_associate_id = assoc_id
        self.refresh_view()
        self._update_direct_indirect_status(assoc_id)

    def _update_direct_indirect_status(self, assoc_id):
        did_direct = self.has_direct_today(assoc_id)
        did_indirect = self.has_indirect_today(assoc_id)
        roles = self.get_indirect_roles_today(assoc_id)
        roles_text = ", ".join(sorted(roles)) if roles else "None"

        if did_direct and did_indirect:
            self.status_var.set(f"DIRECT + INDIRECT ({roles_text})")
            self.status_label.config(fg="#ffcc00")
        elif did_direct:
            self.status_var.set("DIRECT work today")
            self.status_label.config(fg="#00ccff")
        elif did_indirect:
            self.status_var.set(f"INDIRECT work today ({roles_text})")
            self.status_label.config(fg="#ffcc00")
        else:
            self.status_var.set("No DIRECT or INDIRECT work today")
            self.status_label.config(fg="#00cc66")

    # ---------- SESSION CONTROLS ----------

    def start_direct(self):
        if not self.current_associate_id:
            messagebox.showerror("Error", "Load an associate first.")
            return
        start_session(self.current_associate_id, "DIRECT", self.area_var.get(), "N/A")
        self.refresh_view()
        self._update_direct_indirect_status(self.current_associate_id)

    def start_indirect(self):
        if not self.current_associate_id:
            messagebox.showerror("Error", "Load an associate first.")
            return
        start_session(self.current_associate_id, "INDIRECT", self.area_var.get(), self.role_var.get())
        self.refresh_view()
        self._update_direct_indirect_status(self.current_associate_id)

    def end_current(self):
        if not self.current_associate_id:
            messagebox.showerror("Error", "Load an associate first.")
            return
        end_active_session(self.current_associate_id)
        self.refresh_view()
        self._update_direct_indirect_status(self.current_associate_id)

    # ---------- DASHBOARD (UPGRADED) ----------

    def open_dashboard(self):
        dash = tk.Toplevel(self.root)
        dash.title("Dashboard - All Associates")
        dash.configure(bg="#1a1a1a")
        dash.geometry("800x500")

        columns = ("name", "badge", "indirect_hours", "indirect_roles", "status")
        tree = ttk.Treeview(dash, columns=columns, show="headings",
                            style="Dark.Treeview")
        for col in columns:
            tree.heading(col, text=col.replace("_", " ").title(), anchor="w")
            tree.column(col, anchor="w", width=150)
        tree.pack(fill="both", expand=True, padx=10, pady=10)

        associates = get_all_associates()
        rows = []
        for assoc_id, name, badge in associates:
            hours = compute_indirect_hours_today(assoc_id)
            status = self.get_status_text(hours)
            did_direct = self.has_direct_today(assoc_id)
            did_indirect = self.has_indirect_today(assoc_id)
            roles = self.get_indirect_roles_today(assoc_id)
            roles_text = ", ".join(sorted(roles)) if roles else "None"
            rows.append((hours, name, badge, roles_text, status, did_direct, did_indirect, assoc_id))

        rows.sort(reverse=True, key=lambda x: x[0])

        for hours, name, badge, roles_text, status, did_direct, did_indirect, assoc_id in rows:
            if hours >= INDIRECT_LIMIT_HOURS:
                tag = "violation"
            elif hours >= WARNING_THRESHOLD_HOURS:
                tag = "warning"
            elif did_direct and did_indirect:
                tag = "both"
            elif did_direct:
                tag = "direct"
            elif did_indirect:
                tag = "indirect"
            else:
                tag = "none"

            tree.insert(
                "",
                "end",
                values=(name, badge, f"{hours:.2f}", roles_text, status),
                tags=(tag,)
            )

        tree.tag_configure("direct", background="#003366")      # blue
        tree.tag_configure("indirect", background="#665511")    # yellow-brown
        tree.tag_configure("both", background="#CC6600")        # orange
        tree.tag_configure("none", background="#444444")        # gray
        tree.tag_configure("warning", background="#665511")     # yellow-ish
        tree.tag_configure("violation", background="#661111")   # red

    # ---------- EXPORTS ----------

    def export_to_excel(self):
        file_path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv")],
            title="Save as"
        )
        if not file_path:
            return

        associates = get_all_associates()
        with open(file_path, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow(["Name", "Badge ID", "Indirect Hours Today", "Indirect Roles", "Status"])
            for assoc_id, name, badge in associates:
                hours = compute_indirect_hours_today(assoc_id)
                status = self.get_status_text(hours)
                roles = self.get_indirect_roles_today(assoc_id)
                roles_text = ", ".join(sorted(roles)) if roles else "None"
                writer.writerow([name, badge, f"{hours:.2f}", roles_text, status])

        messagebox.showinfo("Export", "Data exported successfully.")

    def export_shift_report(self):
        file_path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv")],
            title="Save shift report as"
        )
        if not file_path:
            return

        associates = get_all_associates()
        today = datetime.date.today().isoformat()

        with open(file_path, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow(["Date", "Name", "Badge ID", "Total Indirect Hours", "Indirect Roles", "Status"])
            for assoc_id, name, badge in associates:
                hours = compute_indirect_hours_today(assoc_id)
                status = self.get_status_text(hours)
                roles = self.get_indirect_roles_today(assoc_id)
                roles_text = ", ".join(sorted(roles)) if roles else "None"
                writer.writerow([today, name, badge, f"{hours:.2f}", roles_text, status])

        messagebox.showinfo("Shift Report", "Shift report exported successfully.")

    # ---------- VIEW / STATUS ----------

    def refresh_view(self):
        if not self.current_associate_id:
            return

        for row in self.tree.get_children():
            self.tree.delete(row)

        sessions = get_today_sessions(self.current_associate_id)
        now = datetime.datetime.now()
        active_path = "No active session"

        for _, start, end, work_type, area, role in sessions:
            start_dt = datetime.datetime.fromisoformat(start)
            end_dt = datetime.datetime.fromisoformat(end) if end else now
            duration_hours = (end_dt - start_dt).total_seconds() / 3600.0
            self.tree.insert("", "end", values=(
                start_dt.strftime("%H:%M"),
                end_dt.strftime("%H:%M") if end else "ACTIVE",
                work_type,
                area,
                role,
                f"{duration_hours:.2f}"
            ))
            if end is None:
                active_path = f"{work_type} → {area} → {role}"

        hours = compute_indirect_hours_today(self.current_associate_id)
        self.indirect_hours_var.set(f"{hours:.2f}")

        status = self.get_status_text(hours)
        self.status_var.set(f"{status} | Path: {active_path}")
        self.apply_status_color(hours)

        if WARNING_THRESHOLD_HOURS <= hours < INDIRECT_LIMIT_HOURS:
            messagebox.showwarning(
                "Warning",
                f"Associate is at {hours:.2f} indirect hours (approaching 6h violation)."
            )
        elif hours >= INDIRECT_LIMIT_HOURS:
            messagebox.showerror(
                "Violation",
                f"Associate is at {hours:.2f} indirect hours (VIOLATION threshold reached)."
            )

    def get_status_text(self, hours):
        if hours >= INDIRECT_LIMIT_HOURS:
            return "VIOLATION (MPV risk)"
        elif hours >= WARNING_THRESHOLD_HOURS:
            return "Near violation"
        else:
            return "OK"

    def apply_status_color(self, hours):
        if hours >= INDIRECT_LIMIT_HOURS:
            color = "#ff3333"
        elif hours >= WARNING_THRESHOLD_HOURS:
            color = "#ffcc00"
        else:
            color = "#00cc66"
        self.status_label.config(fg=color)

# ----------------- MAIN -----------------

def main():
    init_db()
    root = tk.Tk()
    root.withdraw()
    show_splash(root)
    root.after(2000, lambda: root.deiconify())
    app = App(root, CONFIG)
    root.mainloop()

if __name__ == "__main__":
    main()
