import os
import re
import sys
import math
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from dataclasses import dataclass
from typing import List, Optional, Dict

import pandas as pd
import json
import urllib.request
import urllib.error
import matplotlib
from PIL import Image, ImageTk

matplotlib.use("TkAgg")
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure
from matplotlib.patches import Patch

# Windows taskbar icon fix
if sys.platform == 'win32':
    import ctypes
    # Set app user model ID so Windows shows our icon in taskbar
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID('DryerCapacityDashboard')


# ======================================================================
# CONFIG
# ======================================================================
APP_TITLE = "Dryer Capacity Dashboard"

# current application version; bump this with each release
APP_VERSION = "1.0.0"

# url to a JSON file describing the latest release.  the file should look
# like { "version": "1.0.1", "url": "https://.../DryerCapacityDashboard.exe" }
# you can host this on GitHub Pages, S3, a private server, etc.
UPDATE_INFO_URL = "https://example.com/dryer_dashboard/latest.json"

# Handle icon path for both development and PyInstaller frozen exe
def get_resource_path(filename):
    """Get the correct path for resources, works with PyInstaller."""
    if getattr(sys, 'frozen', False):
        # Running as compiled exe
        base_path = sys._MEIPASS
    else:
        # Running as script
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, filename)

ICON_PATH = get_resource_path("dryer.ico")

# Helper for persisting settings

def get_settings_path() -> str:
    r"""Return the path to the settings file, creating parent dirs as needed.

    On Windows we store under %APPDATA%\DryerCapacityDashboard\settings.json;
    on other platforms use a hidden file in the user home directory.
    """
    if sys.platform == 'win32':
        base = os.getenv('APPDATA', os.path.expanduser('~'))
        folder = os.path.join(base, "DryerCapacityDashboard")
        os.makedirs(folder, exist_ok=True)
        return os.path.join(folder, "settings.json")
    else:
        home = os.path.expanduser('~')
        return os.path.join(home, ".dryer_capacity_dashboard.json")


# Default hours available per day (can be changed in settings)
DEFAULT_HOURS_PER_DAY = 24.0

# Default work days per week (for weekly view)
DEFAULT_WORK_DAYS_PER_WEEK = 7

# Default dryer capacity settings (kg per hour)
DEFAULT_DRYER_CAPACITIES = {
    2: 51.0,
    6: 51.0,
    9: 319.0,
    10: 130.0,
    11: 219.0,  # updated to actual throughput
}

# Default wash/cleaning time per dryer (in hours)
DEFAULT_DRYER_WASH_HOURS = {
    2: 1.50,
    6: 1.67,
    9: 3.0,
    10: 2.0,
    11: 2.5,
}


# ======================================================================
# EXCEL FORMAT (SAP export)
# ======================================================================
SAP_PROD_VERSION = "Production Version"        # ex: SD02
SAP_QTY = "Order quantity (GMEIN)"             # numeric
SAP_MATERIAL = "Material Number"               # label
SAP_DATE = "Basic finish date"                 # date column


# ======================================================================
# DATA TYPES
# ======================================================================
@dataclass
class Task:
    dryer: int
    qty: float
    material: str
    order: Optional[str] = None
    material_desc: Optional[str] = None
    date: Optional[str] = None  # Date string for filtering
    # Priority flag (True when Excel's Priority column contains '*')
    priority: bool = False


@dataclass
class DryerSummary:
    dryer: int
    total_units: float
    capacity_per_hour: float
    hours_available: float
    hours_for_drying: float
    wash_count: int
    hours_for_wash: float
    hours_needed: float  # drying + wash
    hours_remaining: float
    utilization_pct: float
    tasks: List[Task]


# ======================================================================
# HELPERS
# ======================================================================
def dryer_from_production_version(val: object) -> Optional[int]:
    """
    SD02 -> 2, SD10 -> 10, etc.
    """
    if val is None:
        return None
    m = re.search(r"(\d+)", str(val))
    if not m:
        return None
    n = int(m.group(1))
    if 1 <= n <= 99:
        return n
    return None


def format_order_label(val: Optional[object]) -> str:
    """Normalize Order value for display.

    - Converts numeric-like values that end in ".0" to integer strings ("12345.0" -> "12345").
    - Preserves non-numeric or already-string values ("ORD-1001" stays the same).
    - Returns empty string for None/NaN.
    """
    if val is None:
        return ""
    s = str(val).strip()
    # handle pandas NaN-like values
    if s.lower() in ("nan", "none", "na"):
        return ""
    # if it's a float-like integer (e.g. "12345.0"), remove the ".0"
    try:
        f = float(s)
        if f.is_integer():
            return str(int(f))
    except Exception:
        pass
    return s


def format_material_label(val: Optional[object]) -> str:
    """Normalize Material value for display.

    - Removes trailing ".0" from numeric-like material strings without converting non-numeric strings.
    - Preserves leading zeros when material is provided as a string.
    - Handles numeric inputs (int/float) by returning integer string.
    """
    if val is None:
        return ""
    # For numeric types, convert to integer string
    if isinstance(val, (int,)):
        return str(val)
    if isinstance(val, float):
        if val.is_integer():
            return str(int(val))
        return str(val)
    s = str(val).strip()
    if not s or s.lower() in ("nan", "none", "na"):
        return ""
    # strip trailing .0 (but preserve leading zeros in the string)
    m = re.match(r"^(\d+)\.0+$", s)
    if m:
        return m.group(1)
    # otherwise, return unchanged
    return s


def load_tasks_from_excel(path: str) -> List[Task]:
    """Load tasks from Excel file (SAP format)."""
    df = pd.read_excel(path)
    
    required = [SAP_PROD_VERSION, SAP_QTY, SAP_MATERIAL]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(
            "Excel format not recognized.\n"
            f"Missing columns: {missing}\n"
            f"Found columns: {list(df.columns)}"
        )

    # Check if date column exists
    has_date = SAP_DATE in df.columns

    # Filter to meaningful rows
    if "Order" in df.columns:
        df2 = df[df["Order"].notna()].copy()
    else:
        df2 = df[df[SAP_QTY].notna()].copy()

    tasks: List[Task] = []
    for _, r in df2.iterrows():
        dryer = dryer_from_production_version(r.get(SAP_PROD_VERSION))
        if not dryer:
            continue

        qty_val = r.get(SAP_QTY)
        if pd.isna(qty_val):
            continue

        try:
            qty = float(qty_val)
        except Exception:
            continue

        material = "" if pd.isna(r.get(SAP_MATERIAL)) else str(r.get(SAP_MATERIAL)).strip()
        if not material:
            material = "UNKNOWN"
        else:
            # normalize material display (remove trailing .0 when present)
            material = format_material_label(material)

        # Optional material description (use common SAP column names if present)
        material_desc = ""
        for desc_col in ("Material description", "Material Description", "Material text", "Material description (EN)"):
            if desc_col in df.columns:
                material_desc = "" if pd.isna(r.get(desc_col)) else str(r.get(desc_col)).strip()
                break

        # Parse date if available
        date_str = None
        if has_date:
            date_val = r.get(SAP_DATE)
            if pd.notna(date_val):
                try:
                    if isinstance(date_val, str):
                        date_str = date_val.strip()
                    else:
                        # Convert datetime to string format
                        date_str = pd.to_datetime(date_val).strftime("%Y-%m-%d")
                except Exception:
                    date_str = str(date_val).strip()

        # Optional Order number (some SAP exports include an 'Order' column)
        order_val = None
        if 'Order' in df.columns:
            ord_raw = r.get('Order')
            if pd.notna(ord_raw):
                order_val = format_order_label(ord_raw)

        # Optional Priority flag (some SAP exports include a 'Priority' column; '*' indicates high priority)
        priority_flag = False
        # find any column name that contains 'priority' (case-insensitive)
        priority_col = None
        for c in df.columns:
            if 'priority' in str(c).lower():
                priority_col = c
                break
        if priority_col is not None:
            pr_raw = r.get(priority_col)
            if pd.notna(pr_raw) and '*' in str(pr_raw):
                priority_flag = True

        tasks.append(Task(dryer=dryer, qty=qty, material=material, order=order_val, material_desc=material_desc, date=date_str, priority=priority_flag))

    return tasks


def count_washes_for_dryer(tasks: List[Task], mode: str = "Unique") -> int:
    """Count washes needed.

    mode: "Unique" = count unique materials; "Segment" = count contiguous material blocks (ordered by date).
    """
    if not tasks:
        return 0

    if mode == "Unique":
        unique_materials = set()
        for t in tasks:
            mat = (t.material or "").strip()
            if mat:
                unique_materials.add(mat)
        return max(1, len(unique_materials)) if tasks else 0

    # Segment mode: count contiguous blocks (needs ordering)
    def _parse_date(s: Optional[str]):
        from datetime import datetime
        if not s:
            return None
        try:
            return datetime.strptime(s, "%Y-%m-%d")
        except Exception:
            try:
                return datetime.strptime(s, "%m/%d/%Y")
            except Exception:
                return None

    # Sort tasks by date (stable)
    indexed = list(enumerate(tasks))
    indexed.sort(key=lambda iv: (_parse_date(iv[1].date) or iv[0]))

    blocks = 0
    last_mat = None
    for _, t in indexed:
        mat = (t.material or "").strip()
        if mat != last_mat:
            blocks += 1
            last_mat = mat
    return max(1, blocks) if tasks else 0


def _parse_date_safe(s: Optional[str]):
    from datetime import datetime
    if not s:
        return None
    try:
        return datetime.strptime(s, "%Y-%m-%d")
    except Exception:
        try:
            return datetime.strptime(s, "%m/%d/%Y")
        except Exception:
            return None


def compute_material_segments(tasks: List[Task], capacity: float, min_run_hours: float = 0.0):
    """Group ordered tasks into contiguous material segments and compute per-segment drying hours.

    Returns list of dicts: {'material', 'qty', 'hours', 'tasks'}
    """
    if not tasks:
        return []

    # Stable sort by date (preserve original order for same-date tasks)
    indexed = list(enumerate(tasks))
    indexed.sort(key=lambda iv: (_parse_date_safe(iv[1].date) or iv[0]))

    segments = []
    cur_mat = None
    cur_qty = 0.0
    cur_tasks = []
    for _, t in indexed:
        mat = (t.material or "").strip()
        if cur_mat is None:
            cur_mat = mat
            cur_qty = t.qty
            cur_tasks = [t]
        elif mat == cur_mat:
            cur_qty += t.qty
            cur_tasks.append(t)
        else:
            # flush
            hours = (cur_qty / capacity) if capacity > 0 else 0.0
            if min_run_hours and hours < min_run_hours:
                hours = min_run_hours
            segments.append({"material": cur_mat, "qty": cur_qty, "hours": hours, "tasks": cur_tasks})
            cur_mat = mat
            cur_qty = t.qty
            cur_tasks = [t]
    # flush last
    if cur_mat is not None:
        hours = (cur_qty / capacity) if capacity > 0 else 0.0
        if min_run_hours and hours < min_run_hours:
            hours = min_run_hours
        segments.append({"material": cur_mat, "qty": cur_qty, "hours": hours, "tasks": cur_tasks})

    return segments


def compute_dryer_summaries(
    tasks: List[Task],
    capacities: Dict[int, float],
    hours_available: float,
    wash_hours: Dict[int, float],
    wash_count_mode: str = "Segment",
    min_run_hours: float = 0.0
) -> List[DryerSummary]:
    """Compute summary for each dryer based on tasks and capacity settings.

    wash_count_mode: "Segment" (default) or "Unique".
    min_run_hours: if >0, round each material-segment drying time up to this minimum.
    """
    # Group tasks by dryer
    tasks_by_dryer: Dict[int, List[Task]] = {}
    for t in tasks:
        tasks_by_dryer.setdefault(t.dryer, []).append(t)

    # Always include all 5 dryers
    all_dryers = [2, 6, 9, 10, 11]

    summaries: List[DryerSummary] = []

    for dryer in all_dryers:
        dryer_tasks = tasks_by_dryer.get(dryer, [])
        total_units = sum(t.qty for t in dryer_tasks)

        capacity = capacities.get(dryer, 75.0)  # default 75 kg/hour

        # Compute drying time by summing per-segment hours (respects min_run_hours)
        segments = compute_material_segments(dryer_tasks, capacity, min_run_hours)
        hours_for_drying = sum(s["hours"] for s in segments) if segments else 0.0

        # Wash count depends on mode
        if wash_count_mode == "Unique":
            wash_count = count_washes_for_dryer(dryer_tasks, mode="Unique")
        else:
            wash_count = len(segments) if segments else 0
            wash_count = max(1, wash_count) if dryer_tasks else 0

        dryer_wash_hours = wash_hours.get(dryer, 2.0)  # default 2 hours
        hours_for_wash = wash_count * dryer_wash_hours

        hours_needed = hours_for_drying + hours_for_wash
        hours_remaining = max(0.0, hours_available - hours_needed)
        utilization = (hours_needed / hours_available * 100) if hours_available > 0 else 0.0

        summaries.append(DryerSummary(
            dryer=dryer,
            total_units=total_units,
            capacity_per_hour=capacity,
            hours_available=hours_available,
            hours_for_drying=hours_for_drying,
            wash_count=wash_count,
            hours_for_wash=hours_for_wash,
            hours_needed=hours_needed,
            hours_remaining=hours_remaining,
            utilization_pct=utilization,
            tasks=dryer_tasks
        ))

    return summaries


# ======================================================================
# GUI
# ======================================================================
class App(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title(APP_TITLE)
        self.geometry("1120x760")

        # ensure settings are saved when the app exits
        self.protocol("WM_DELETE_WINDOW", self.on_close)

        try:
            if os.path.exists(ICON_PATH):
                self.iconbitmap(ICON_PATH)
        except Exception:
            pass

        self.file_path: Optional[str] = None
        
        # Dryer capacity settings (kg per hour)
        self.dryer_capacities: Dict[int, float] = dict(DEFAULT_DRYER_CAPACITIES)
        # Per-dryer wash times (hours)
        self.dryer_wash_hours: Dict[int, float] = dict(DEFAULT_DRYER_WASH_HOURS)
        self.hours_per_day = tk.DoubleVar(value=DEFAULT_HOURS_PER_DAY)
        self.work_days_per_week = tk.IntVar(value=DEFAULT_WORK_DAYS_PER_WEEK)
        # Chart max hours (controls Y-axis maximum for capacity chart)
        # Default chart Y-axis increased to 30 hours per user request
        self.chart_max_hours = tk.DoubleVar(value=30.0)
        # Reference/break point (red dashed line) used in Day view (default 24h)
        self.reference_hours = tk.DoubleVar(value=24.0)
        # Wash counting mode: "Segment" (contiguous blocks) or "Unique" (unique materials)
        self.wash_count_mode = tk.StringVar(value="Segment")
        # Minimum run hours per material segment (0 => disabled)
        self.min_run_hours = tk.DoubleVar(value=0.0)

        # load persisted settings (overwrites defaults)
        self.load_settings()

        self.container = tk.Frame(self)
        self.container.pack(fill="both", expand=True)
        self.container.grid_rowconfigure(0, weight=1)
        self.container.grid_columnconfigure(0, weight=1)

        self.frames: Dict[str, tk.Frame] = {}
        for F in (StartPage, DashboardPage):
            frame = F(parent=self.container, controller=self)
            self.frames[F.__name__] = frame
            frame.grid(row=0, column=0, sticky="nsew")

        self.show_frame("StartPage")
        # automatically check for updates on launch (no popup if up‑to‑date)
        self.check_for_updates(silent=True)

    def show_frame(self, name: str):
        self.frames[name].tkraise()

    # ---------- settings persistence ----------
    def load_settings(self):
        """Read configuration from disk and apply to controller variables."""
        path = get_settings_path()
        if not os.path.exists(path):
            return
        try:
            with open(path, 'r') as f:
                data = json.load(f)
        except Exception:
            # ignore errors, keep defaults
            return
        self.hours_per_day.set(data.get('hours_per_day', self.hours_per_day.get()))
        self.work_days_per_week.set(data.get('work_days_per_week', self.work_days_per_week.get()))
        self.chart_max_hours.set(data.get('chart_max_hours', self.chart_max_hours.get()))
        self.reference_hours.set(data.get('reference_hours', self.reference_hours.get()))
        self.wash_count_mode.set(data.get('wash_count_mode', self.wash_count_mode.get()))
        self.min_run_hours.set(data.get('min_run_hours', self.min_run_hours.get()))
        # dictionaries stored directly
        # dictionaries may come back with string keys from JSON; convert to int
        for k, v in (data.get('dryer_capacities', {}) or {}).items():
            try:
                self.dryer_capacities[int(k)] = float(v)
            except Exception:
                pass
        for k, v in (data.get('dryer_wash_hours', {}) or {}).items():
            try:
                self.dryer_wash_hours[int(k)] = float(v)
            except Exception:
                pass
        # optional file path remembered
        self.file_path = data.get('last_file_path', self.file_path)

    def save_settings(self):
        """Write current settings values to disk."""
        path = get_settings_path()
        data = {
            'hours_per_day': self.hours_per_day.get(),
            'work_days_per_week': self.work_days_per_week.get(),
            'chart_max_hours': self.chart_max_hours.get(),
            'reference_hours': self.reference_hours.get(),
            'wash_count_mode': self.wash_count_mode.get(),
            'min_run_hours': self.min_run_hours.get(),
            'dryer_capacities': self.dryer_capacities,
            'dryer_wash_hours': self.dryer_wash_hours,
            'last_file_path': self.file_path,
        }
        try:
            with open(path, 'w') as f:
                json.dump(data, f, indent=2)
        except Exception:
            pass

    def on_close(self):
        # save settings before quitting
        self.save_settings()
        self.destroy()

    def check_for_updates(self, silent: bool = False) -> Optional[Dict[str, str]]:
        """Query the update server and return info if a newer version exists.

        If ``silent`` is False, a messagebox will notify the user of the result.
        Returns the parsed info dict when there is an update, otherwise None.

        When an update is found and ``silent`` is False the user is prompted to
        download and run the new executable immediately.
        """
        try:
            with urllib.request.urlopen(UPDATE_INFO_URL, timeout=5) as resp:
                raw = resp.read().decode('utf-8')
        except Exception as e:
            if not silent:
                messagebox.showwarning("Update Check", f"Could not reach update server: {e}")
            return None

        try:
            info = json.loads(raw)
            latest = info.get('version')
            if latest and self._version_greater(latest, APP_VERSION):
                if not silent:
                    answer = messagebox.askyesno(
                        "Update Available",
                        f"A new version {latest} is available.\n"
                        f"Would you like to download and run it now?"
                    )
                    if answer:
                        url = info.get('url') or UPDATE_INFO_URL
                        self._download_and_run_update(url)
                return info
            else:
                if not silent:
                    messagebox.showinfo("Update Check", "You are running the latest version.")
        except Exception as e:
            if not silent:
                messagebox.showwarning("Update Check", f"Bad response from update server: {e}")
        return None

    def _version_greater(self, a: str, b: str) -> bool:
        """Return True if version string *a* is strictly greater than *b*.

        Simple semantic comparison splitting on dots; non‑numeric parts are
        compared lexically.  This is just for basic use and can be replaced by
        packaging.version.parse if you add a dependency.
        """
        def norm(v):
            parts = v.split('.')
            out = []
            for p in parts:
                if p.isdigit():
                    out.append(int(p))
                else:
                    out.append(p)
            return out
        return norm(a) > norm(b)

    def _download_and_run_update(self, url: str):
        """Download the file at *url* to a temporary location and open it.

        The download occurs synchronously; any errors are reported via
        messageboxes.  After launching the downloaded executable the current
        application exits (on Windows the new installer can overwrite the
        running exe).
        """
        import tempfile, shutil
        try:
            # guess filename from URL
            filename = os.path.basename(url.split('?', 1)[0]) or "update.exe"
            dest = os.path.join(tempfile.gettempdir(), filename)
            # download
            with urllib.request.urlopen(url, timeout=30) as resp, open(dest, 'wb') as out:
                shutil.copyfileobj(resp, out)
        except Exception as e:
            messagebox.showerror("Update Error", f"Failed to download update: {e}")
            return
        try:
            # open the downloaded file; on Windows this will launch the exe
            os.startfile(dest)
        except Exception as e:
            messagebox.showwarning("Update", f"Downloaded update to {dest}\nUnable to launch automatically: {e}")
        # quit current app to allow installer to replace it
        self.destroy()

    def run_execute(self):
        try:
            dash: DashboardPage = self.frames["DashboardPage"]  # type: ignore

            if self.file_path:
                tasks = load_tasks_from_excel(self.file_path)
                dash.load_data(tasks)
            else:
                # Mock data for testing - uses dryers 2, 6, 9, 10, 11
                mock_tasks = [
                    Task(dryer=2, qty=800, material="10001234", order="ORD-1001", material_desc="STRAWBERRY FLV"),
                    Task(dryer=2, qty=400, material="10007890", order="ORD-1002", material_desc="VANILLA EXTRACT"),
                    Task(dryer=6, qty=1200, material="10001111", order="ORD-2001", material_desc="BLACKBERRY FLV"),
                    Task(dryer=6, qty=600, material="10002222", order="ORD-2002", material_desc="APPLE CINNAMON"),
                    Task(dryer=9, qty=900, material="10003333", order="ORD-3001", material_desc="LEMONADE FLV"),
                    Task(dryer=9, qty=1500, material="10004444", order="ORD-3002", material_desc="LIME FLV"),
                    Task(dryer=10, qty=700, material="10005555", order="ORD-4001", material_desc="WATERMELON EF100", priority=True),
                    Task(dryer=10, qty=500, material="10006666", order="ORD-4002", material_desc="CINNAMON ROLL DRY"),
                    Task(dryer=11, qty=1100, material="10007777", order="ORD-5001", material_desc="CARROT"),
                    Task(dryer=11, qty=800, material="10008888", order="ORD-5002", material_desc="POMEGRANATE FLV"),
                ]
                dash.load_data(mock_tasks)

            self.show_frame("DashboardPage")

        except Exception as e:
            messagebox.showerror("Error", str(e))


class StartPage(tk.Frame):
    def __init__(self, parent, controller: App):
        super().__init__(parent, bg="#f0f2f5")
        self.controller = controller

        # Center container
        center_frame = tk.Frame(self, bg="#f0f2f5")
        center_frame.place(relx=0.5, rely=0.5, anchor="center")

        # Main card - tighter width
        card = tk.Frame(center_frame, bg="white", highlightbackground="#e0e0e0", highlightthickness=1)
        card.pack()

        card_inner = tk.Frame(card, bg="white")
        card_inner.pack(padx=40, pady=32)

        # Icon/Logo area - load from ico file
        try:
            if os.path.exists(ICON_PATH):
                ico_image = Image.open(ICON_PATH)
                ico_image = ico_image.resize((64, 64), Image.Resampling.LANCZOS)
                self.logo_photo = ImageTk.PhotoImage(ico_image)
                logo_label = tk.Label(card_inner, image=self.logo_photo, bg="white")
                logo_label.pack(pady=(0, 16))
            else:
                raise FileNotFoundError()
        except Exception:
            # Fallback to text if ico not available
            icon_frame = tk.Frame(card_inner, bg="#0078d4", width=64, height=64)
            icon_frame.pack(pady=(0, 16))
            icon_frame.pack_propagate(False)
            tk.Label(icon_frame, text="📊", font=("Segoe UI", 24), bg="#0078d4", fg="white").pack(expand=True)

        # Title
        title = tk.Label(card_inner, text=APP_TITLE, font=("Segoe UI", 20, "bold"), 
                        bg="white", fg="#111827")
        title.pack(pady=(0, 4))

        # Subtitle
        subtitle = tk.Label(card_inner, text="Analyze dryer capacity and cleaning schedules",
                           font=("Segoe UI", 10), bg="white", fg="#6b7280")
        subtitle.pack(pady=(0, 24))

        # File selection area
        file_frame = tk.Frame(card_inner, bg="#f9fafb", highlightbackground="#e5e7eb", highlightthickness=1)
        file_frame.pack(fill="x", pady=(0, 16))

        file_inner = tk.Frame(file_frame, bg="#f9fafb")
        file_inner.pack(padx=24, pady=14)

        tk.Label(file_inner, text="📁", font=("Segoe UI", 14), bg="#f9fafb").pack()
        
        self.path_label = tk.Label(file_inner, text="No file selected", 
                                   font=("Segoe UI", 9), bg="#f9fafb", fg="#6b7280")
        self.path_label.pack(pady=(2, 6))
        # reflect last-used file if saved in settings
        if controller.file_path:
            display_path = controller.file_path if len(controller.file_path) < 50 else "..." + controller.file_path[-47:]
            self.path_label.config(text=display_path, fg="#059669")

        select_btn = tk.Button(file_inner, text="Select Excel File", 
                              command=self.select_file,
                              font=("Segoe UI", 9), bg="white", fg="#374151",
                              relief="solid", bd=1, cursor="hand2", padx=14, pady=3)
        select_btn.pack()

        # Execute button
        exec_btn = tk.Button(card_inner, text="Execute Analysis", 
                            command=self.controller.run_execute,
                            font=("Segoe UI", 10, "bold"), bg="#0078d4", fg="white",
                            relief="flat", bd=0, cursor="hand2", padx=28, pady=8,
                            activebackground="#005a9e", activeforeground="white")
        exec_btn.pack(pady=(8, 16))

        # Help text
        help_frame = tk.Frame(card_inner, bg="white")
        help_frame.pack()

        tk.Label(help_frame, text="Expected Excel columns:", 
                font=("Segoe UI", 9, "bold"), bg="white", fg="#374151").pack()
        
        columns_frame = tk.Frame(help_frame, bg="white")
        columns_frame.pack(pady=(6, 0))
        
        for col in ["Production Version (SD02)", "Order quantity (GMEIN)", "Material Number"]:
            col_label = tk.Label(columns_frame, text=f"• {col}", 
                                font=("Segoe UI", 9), bg="white", fg="#6b7280", anchor="w")
            col_label.pack(anchor="w")

        # Footer note
        note = tk.Label(card_inner, 
                       text="If no file is selected, the app will use mock data for demonstration.",
                       font=("Segoe UI", 9, "italic"), bg="white", fg="#9ca3af")
        note.pack(pady=(20, 0))

    def select_file(self):
        path = filedialog.askopenfilename(
            title="Select Excel file",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")],
        )
        if path:
            self.controller.file_path = path
            # Show truncated path if too long
            display_path = path if len(path) < 50 else "..." + path[-47:]
            self.path_label.config(text=display_path, fg="#059669")


class DashboardPage(tk.Frame):
    def __init__(self, parent, controller: App):
        super().__init__(parent, bg="#f0f2f5")
        self.controller = controller

        self.tasks: List[Task] = []
        self.filtered_tasks: List[Task] = []  # Tasks filtered by selected date
        self.summaries: List[DryerSummary] = []
        self.available_dates: List[str] = []  # Dates found in data
        self.week_dates: Dict[str, List[str]] = {}  # Week label -> list of dates
        self.selected_date = tk.StringVar(value="")
        self.view_mode = tk.StringVar(value="Day")  # "Day" or "Week"
        # mpl pick-event connection id for capacity chart (used for double-click audit)
        self._capacity_pick_cid = None
        # Track open audit windows by dryer to avoid duplicates
        self._open_audits: Dict[int, tk.Toplevel] = {}

        # Modern header bar - full width with accent color
        header = tk.Frame(self, bg="#0078d4")
        header.pack(fill="x")
        
        header_inner = tk.Frame(header, bg="#0078d4")
        header_inner.pack(fill="x", padx=24, pady=12)

        back_btn = tk.Button(header_inner, text="←", command=lambda: controller.show_frame("StartPage"),
                            font=("Segoe UI", 14), bg="#0078d4", fg="white", 
                            relief="flat", bd=0, cursor="hand2", activeforeground="#cce4f7",
                            activebackground="#0078d4")
        back_btn.pack(side="left")

        # Date selector in header
        date_frame = tk.Frame(header_inner, bg="#0078d4")
        date_frame.pack(side="left", padx=(20, 0))
        
        tk.Label(date_frame, text="Date:", font=("Segoe UI", 10), 
                bg="#0078d4", fg="white").pack(side="left", padx=(0, 8))
        
        self.date_combo = ttk.Combobox(date_frame, textvariable=self.selected_date, 
                                        state="readonly", width=18, font=("Segoe UI", 10))
        self.date_combo['values'] = []
        self.date_combo.pack(side="left")
        self.date_combo.bind("<<ComboboxSelected>>", lambda e: self.on_date_changed())

        # Day/Week/Month toggle
        toggle_frame = tk.Frame(header_inner, bg="#0078d4")
        toggle_frame.pack(side="left", padx=(25, 0))
        
        self.day_btn = tk.Button(toggle_frame, text="Day", font=("Segoe UI", 9, "bold"),
                                  bg="white", fg="#0078d4", relief="flat", bd=0,
                                  padx=12, pady=4, cursor="hand2",
                                  command=lambda: self.set_view_mode("Day"))
        self.day_btn.pack(side="left")
        
        self.week_btn = tk.Button(toggle_frame, text="Week", font=("Segoe UI", 9),
                                   bg="#0078d4", fg="white", relief="flat", bd=0,
                                   padx=12, pady=4, cursor="hand2",
                                   command=lambda: self.set_view_mode("Week"))
        self.week_btn.pack(side="left")
        
        self.month_btn = tk.Button(toggle_frame, text="Month", font=("Segoe UI", 9),
                                    bg="#0078d4", fg="white", relief="flat", bd=0,
                                    padx=12, pady=4, cursor="hand2",
                                    command=lambda: self.set_view_mode("Month"))
        self.month_btn.pack(side="left")

        # Load settings icon from image
        try:
            settings_img = Image.open(get_resource_path("setting1.png"))
            settings_img = settings_img.resize((24, 24), Image.Resampling.LANCZOS)
            self.settings_photo = ImageTk.PhotoImage(settings_img)
            settings_btn = tk.Button(header_inner, image=self.settings_photo, command=self.open_settings,
                                    bg="#0078d4", relief="flat", bd=0, cursor="hand2",
                                    activebackground="#0078d4")
        except Exception:
            # Fallback to text if image not available
            settings_btn = tk.Button(header_inner, text="⚙", command=self.open_settings,
                                    font=("Segoe UI", 16), bg="#0078d4", fg="white",
                                    relief="flat", bd=0, cursor="hand2",
                                    activebackground="#0078d4")
        settings_btn.pack(side="right")
        
        # Add subtle shadow line under header
        tk.Frame(self, bg="#005a9e", height=2).pack(fill="x")

        # Summary cards container - full width
        self.summary_container = tk.Frame(self, bg="#f0f2f5")
        self.summary_container.pack(fill="x", padx=24, pady=(16, 0))
        
        # Create individual stat cards
        self.stat_cards = {}
        card_configs = [
            ("units", "Total KG", "0", "#3b82f6"),
            ("drying", "Drying Time", "0h", "#10b981"),
            ("washes", "Washes", "0", "#f59e0b"),
            ("needed", "Total Needed", "0h", "#8b5cf6"),
            ("remaining", "Remaining", "0h", "#06b6d4"),
            ("util", "Utilization", "0%", "#ef4444"),
        ]
        
        for i, (key, label, default, color) in enumerate(card_configs):
            card = tk.Frame(self.summary_container, bg="white", highlightbackground="#e5e7eb", highlightthickness=1)
            card.grid(row=0, column=i, padx=4, pady=0, sticky="nsew")
            self.summary_container.columnconfigure(i, weight=1)
            
            card_inner = tk.Frame(card, bg="white")
            card_inner.pack(fill="both", expand=True, padx=12, pady=10)
            
            value_label = tk.Label(card_inner, text=default, font=("Segoe UI", 16, "bold"), 
                                   bg="white", fg=color)
            value_label.pack(anchor="w")
            
            title_label = tk.Label(card_inner, text=label, font=("Segoe UI", 9), 
                                   bg="white", fg="#6b7280")
            title_label.pack(anchor="w")
            
            self.stat_cards[key] = value_label

        # Notebook container - full width
        notebook_container = tk.Frame(self, bg="#f0f2f5")
        notebook_container.pack(fill="both", expand=True, padx=24, pady=16)
        
        # Style the notebook tabs
        style = ttk.Style()
        style.configure("Modern.TNotebook", background="#f0f2f5", borderwidth=0)
        style.configure("Modern.TNotebook.Tab", font=("Segoe UI", 10), padding=[16, 8])
        style.map("Modern.TNotebook.Tab", 
                  background=[("selected", "white"), ("!selected", "#e5e7eb")],
                  foreground=[("selected", "#111827"), ("!selected", "#6b7280")])

        self.notebook = ttk.Notebook(notebook_container, style="Modern.TNotebook")
        self.notebook.pack(fill="both", expand=True)

        self.tab_capacity = tk.Frame(self.notebook, bg="white")
        self.tab_details = tk.Frame(self.notebook, bg="white")

        self.notebook.add(self.tab_capacity, text="  Capacity Overview  ")
        self.notebook.add(self.tab_details, text="  Dryer Details  ")

        # Capacity chart - fill available space
        self.fig_capacity = Figure(figsize=(10, 5), dpi=100, facecolor="white")
        # Give more room on the left/right/top so axis labels and the 24h reference are not clipped
        self.fig_capacity.subplots_adjust(left=0.08, right=0.96, top=0.94, bottom=0.16)
        self.ax_capacity = self.fig_capacity.add_subplot(111)
        self.ax_capacity.set_facecolor("white")
        self.canvas_capacity = FigureCanvasTkAgg(self.fig_capacity, master=self.tab_capacity)
        self.canvas_capacity.get_tk_widget().pack(fill="both", expand=True)
        # Hint for users
        hint_lbl = tk.Label(self.tab_capacity, text="Double-click a bar to open dryer audit", font=("Segoe UI", 9), bg="white", fg="#6b7280")
        hint_lbl.pack(anchor="w", padx=12, pady=(6, 0))

        # Details list
        self.details_frame = tk.Frame(self.tab_details, bg="white")
        self.details_frame.pack(fill="both", expand=True, padx=10, pady=10)

    def load_data(self, tasks: List[Task]):
        self.tasks = tasks
        
        # Extract unique dates from tasks (sorted earliest first)
        dates = sorted(set(t.date for t in tasks if t.date))
        self.available_dates = dates
        
        # Build week options from available dates
        self.week_dates = {}  # Maps "Week of <date>" -> list of dates in that week
        if dates:
            from datetime import datetime
            # Group dates by week (using Monday as start of week)
            for date_str in dates:
                try:
                    dt = datetime.strptime(date_str, "%Y-%m-%d")
                except ValueError:
                    try:
                        dt = datetime.strptime(date_str, "%m/%d/%Y")
                    except ValueError:
                        continue
                # Find Monday of this week
                monday = dt - pd.Timedelta(days=dt.weekday())
                week_key = f"Week of {monday.strftime('%m/%d')}"
                if week_key not in self.week_dates:
                    self.week_dates[week_key] = []
                self.week_dates[week_key].append(date_str)
        
        # Update date dropdown - only individual dates (use Day/Week/Month toggle to switch views)
        if dates:
            self.date_combo['values'] = dates
            self.selected_date.set(dates[0])  # Select earliest date
        else:
            self.date_combo['values'] = []
            self.selected_date.set("")
        
        self.render_all()

    def set_view_mode(self, mode):
        """Toggle between Day, Week, and Month view."""
        self.view_mode.set(mode)
        active_style = {"bg": "white", "fg": "#0078d4", "font": ("Segoe UI", 9, "bold")}
        inactive_style = {"bg": "#0078d4", "fg": "white", "font": ("Segoe UI", 9)}
        
        self.day_btn.config(**(active_style if mode == "Day" else inactive_style))
        self.week_btn.config(**(active_style if mode == "Week" else inactive_style))
        self.month_btn.config(**(active_style if mode == "Month" else inactive_style))
        
        if mode == "Day":
            # Show all dates in Day mode and select the first one
            if self.available_dates:
                self.date_combo['values'] = self.available_dates
                self.selected_date.set(self.available_dates[0])
        elif mode == "Week":
            # Show only Mondays in Week mode
            if self.available_dates:
                from datetime import datetime
                mondays = []
                seen_weeks = set()
                for date_str in self.available_dates:
                    try:
                        dt = datetime.strptime(date_str, "%Y-%m-%d")
                    except ValueError:
                        try:
                            dt = datetime.strptime(date_str, "%m/%d/%Y")
                        except ValueError:
                            continue
                    monday = dt - pd.Timedelta(days=dt.weekday())
                    monday_str = monday.strftime("%Y-%m-%d")
                    if monday_str not in seen_weeks:
                        seen_weeks.add(monday_str)
                        mondays.append(monday.strftime("%m/%d/%Y") + " (Week)")
                self.date_combo['values'] = mondays if mondays else self.available_dates
                if mondays and self.selected_date.get() not in mondays:
                    self.selected_date.set(mondays[0])
        elif mode == "Month":
            # Show only month entries
            if self.available_dates:
                from datetime import datetime
                months = []
                seen_months = set()
                for date_str in self.available_dates:
                    try:
                        dt = datetime.strptime(date_str, "%Y-%m-%d")
                    except ValueError:
                        try:
                            dt = datetime.strptime(date_str, "%m/%d/%Y")
                        except ValueError:
                            continue
                    month_key = dt.strftime("%Y-%m")
                    if month_key not in seen_months:
                        seen_months.add(month_key)
                        months.append(dt.strftime("%B %Y"))
                self.date_combo['values'] = months if months else self.available_dates
                if months and self.selected_date.get() not in months:
                    self.selected_date.set(months[0])
        self.render_all()

    def on_date_changed(self):
        """Called when date selection changes."""
        self.render_all()

    def render_all(self):
        # Filter tasks by selected date, week, or month
        selected = self.selected_date.get()
        view_mode = self.view_mode.get()
        
        if not selected:
            self.filtered_tasks = self.tasks
        elif view_mode == "Week":
            # Week view - find all dates in the same week as selected Monday
            from datetime import datetime
            date_part = selected.replace(" (Week)", "")
            try:
                dt = datetime.strptime(date_part, "%Y-%m-%d")
            except ValueError:
                try:
                    dt = datetime.strptime(date_part, "%m/%d/%Y")
                except ValueError:
                    self.filtered_tasks = self.tasks
                    return
            monday = dt
            week_key = f"Week of {monday.strftime('%m/%d')}"
            week_dates = self.week_dates.get(week_key, [])
            self.filtered_tasks = [t for t in self.tasks if t.date in week_dates]
        elif view_mode == "Month":
            # Month view - find all dates in the selected month
            from datetime import datetime
            try:
                dt = datetime.strptime(selected, "%B %Y")
            except ValueError:
                self.filtered_tasks = self.tasks
                return
            target_year = dt.year
            target_month = dt.month
            month_dates = []
            for date_str in self.available_dates:
                try:
                    d = datetime.strptime(date_str, "%Y-%m-%d")
                except ValueError:
                    try:
                        d = datetime.strptime(date_str, "%m/%d/%Y")
                    except ValueError:
                        continue
                if d.year == target_year and d.month == target_month:
                    month_dates.append(date_str)
            self.filtered_tasks = [t for t in self.tasks if t.date in month_dates]
        else:
            self.filtered_tasks = [t for t in self.tasks if t.date == selected]
        
        # Calculate hours available based on view (day vs week vs month)
        hours_per_day = float(self.controller.hours_per_day.get())
        work_days_per_week = self.controller.work_days_per_week.get()
        if view_mode == "Week":
            total_hours = hours_per_day * work_days_per_week
        elif view_mode == "Month":
            import calendar
            from datetime import datetime
            try:
                dt = datetime.strptime(selected, "%B %Y")
            except ValueError:
                total_hours = hours_per_day
            else:
                days_in_month = calendar.monthrange(dt.year, dt.month)[1]
                work_days_in_month = (work_days_per_week / 7) * days_in_month
                total_hours = hours_per_day * work_days_in_month
        else:
            total_hours = hours_per_day
        
        self.summaries = compute_dryer_summaries(
            self.filtered_tasks,
            self.controller.dryer_capacities,
            total_hours,  # Use calculated hours (day or week)
            self.controller.dryer_wash_hours,
            wash_count_mode=self.controller.wash_count_mode.get(),
            min_run_hours=self.controller.min_run_hours.get()
        )
        
        # Show all dryers (not just ones with tasks)
        active_summaries = self.summaries
        
        if not active_summaries:
            # Reset stat cards
            self.stat_cards["units"].config(text="0")
            self.stat_cards["drying"].config(text="0h")
            self.stat_cards["washes"].config(text="0")
            self.stat_cards["needed"].config(text="0h")
            self.stat_cards["remaining"].config(text="0h")
            self.stat_cards["util"].config(text="0%")
            self.ax_capacity.clear()
            self.canvas_capacity.draw()
            return

        # Overall summary
        total_units = sum(s.total_units for s in active_summaries)
        total_hours_drying = sum(s.hours_for_drying for s in active_summaries)
        total_washes = sum(s.wash_count for s in active_summaries)
        total_hours_wash = sum(s.hours_for_wash for s in active_summaries)
        total_hours_needed = sum(s.hours_needed for s in active_summaries)
        total_hours_available = sum(s.hours_available for s in active_summaries)
        # Cap remaining: if a dryer exceeds capacity (negative), don't let it reduce overall remaining
        total_remaining = sum(max(0, s.hours_remaining) for s in active_summaries)
        overall_util = (total_hours_needed / total_hours_available * 100) if total_hours_available > 0 else 0

        # Update stat cards
        self.stat_cards["units"].config(text=f"{total_units:,.0f}")
        self.stat_cards["drying"].config(text=f"{total_hours_drying:.1f}h")
        self.stat_cards["washes"].config(text=f"{total_washes} ({total_hours_wash:.1f}h)")
        self.stat_cards["needed"].config(text=f"{total_hours_needed:.1f}h")
        self.stat_cards["remaining"].config(text=f"{total_remaining:.1f}h")
        self.stat_cards["util"].config(text=f"{overall_util:.1f}%")

        self.render_capacity_chart(active_summaries)
        self.render_details(active_summaries)

    def render_capacity_chart(self, summaries: List[DryerSummary]):
        self.ax_capacity.clear()

        labels = [f"Dryer {s.dryer}" for s in summaries]
        hours_drying = [s.hours_for_drying for s in summaries]
        hours_wash = [s.hours_for_wash for s in summaries]
        hours_remaining = [s.hours_remaining for s in summaries]
        hours_available = [s.hours_available for s in summaries]

        x = range(len(summaries))
        width = 0.6

        # Stacked bar: drying (single block per dryer) + wash + remaining (pickable — double-click to audit)
        bars_drying = self.ax_capacity.bar(x, hours_drying, width, label="Drying", color="tab:blue", picker=True)
        bars_wash = self.ax_capacity.bar(x, hours_wash, width, bottom=hours_drying, 
                             label="Cleaning", color="tab:orange", picker=True)
        bottoms_for_remaining = [hours_drying[i] + hours_wash[i] for i in range(len(summaries))]
        bars_remaining = self.ax_capacity.bar(x, hours_remaining, width, bottom=bottoms_for_remaining,
                             label="Remaining", color="tab:green", alpha=0.6, picker=True)

        # Tag each rectangle with its dryer id so the pick handler can find which dryer was clicked
        for i, s in enumerate(summaries):
            for bc in (bars_drying, bars_wash, bars_remaining):
                try:
                    rect = bc[i]
                    rect.set_gid(str(s.dryer))
                except Exception:
                    pass

        # Ensure we have a single pick-event connection (disconnect previous if present)
        try:
            if getattr(self, '_capacity_pick_cid', None):
                self.canvas_capacity.mpl_disconnect(self._capacity_pick_cid)
        except Exception:
            pass
        self._capacity_pick_cid = self.canvas_capacity.mpl_connect('pick_event', self._on_capacity_pick)

        # Draw cleaning and remaining bars stacked on top of the aggregated drying height
        bars_wash = self.ax_capacity.bar(x, hours_wash, width, bottom=hours_drying,
                             label="Cleaning", color="tab:orange", picker=True)
        bottoms_for_remaining = [hours_drying[i] + hours_wash[i] for i in range(len(summaries))]
        bars_remaining = self.ax_capacity.bar(x, hours_remaining, width, bottom=bottoms_for_remaining,
                             label="Remaining", color="tab:green", alpha=0.6, picker=True)

        # Tag wash/remaining rectangles with dryer id as before
        for i, s in enumerate(summaries):
            try:
                bars_wash[i].set_gid(str(s.dryer))
            except Exception:
                pass
            try:
                bars_remaining[i].set_gid(str(s.dryer))
            except Exception:
                pass

        # Ensure we have a single pick-event connection (disconnect previous if present)
        try:
            if getattr(self, '_capacity_pick_cid', None):
                self.canvas_capacity.mpl_disconnect(self._capacity_pick_cid)
        except Exception:
            pass
        self._capacity_pick_cid = self.canvas_capacity.mpl_connect('pick_event', self._on_capacity_pick) 

        # Add labels on bars (aggregate drying + washing + remaining)
        for i, s in enumerate(summaries):
            # Drying label (aggregate)
            if s.hours_for_drying > 0.5:
                self.ax_capacity.text(
                    i, s.hours_for_drying / 2, 
                    f"{s.hours_for_drying:.1f}h",
                    ha="center", va="center", fontsize=8, fontweight="bold", color="white"
                )
            # Wash label - show hours
            if s.hours_for_wash > 0.3:
                self.ax_capacity.text(
                    i, s.hours_for_drying + s.hours_for_wash / 2,
                    f"{s.hours_for_wash:.1f}h",
                    ha="center", va="center", fontsize=8, fontweight="bold", color="white"
                )
            # Remaining label
            if s.hours_remaining > 0.5:
                self.ax_capacity.text(
                    i, s.hours_for_drying + s.hours_for_wash + s.hours_remaining / 2,
                    f"{s.hours_remaining:.1f}h",
                    ha="center", va="center", fontsize=8, color="black"
                )

        # Per-task labels are drawn inline with each task block above — no further annotation required here.

        self.ax_capacity.set_xticks(x)
        self.ax_capacity.set_xticklabels(labels)
        self.ax_capacity.set_ylabel("Hours")
        self.ax_capacity.set_title("Dryer Capacity: Drying + Cleaning vs Available")
        
        # Y-axis: use user chart max but ensure it's at least large enough for data (include needed/over-capacity)
        user_chart_max = float(self.controller.chart_max_hours.get())
        data_max = max((max(hours_available) if hours_available else 0), max((s.hours_needed for s in summaries), default=0), max((s.hours_for_drying for s in summaries), default=0))
        y_max = max(user_chart_max, data_max)
        # set Y limits slightly above maximum value
        self.ax_capacity.set_ylim(0, y_max * 1.05)

        # Y-axis ticks: choose sensible interval per view mode
        # - Day view: ticks every 5 hours
        # - Week view: ticks every 24 hours
        # - Month / fallback: ticks every 24 hours
        try:
            view = (self.view_mode.get() if getattr(self, 'view_mode', None) else "Day")
            if view == "Day":
                step = 5
            elif view == "Week":
                step = 24
            elif view == "Month":
                # show ticks per week (24 * 7 hours)
                step = 24 * 7
            else:
                step = 24

            y_top = max(1.0, float(math.ceil(y_max)))
            # round up to the nearest step so the top tick covers the data
            last_tick = int(((y_top + step - 1) // step) * step)
            ticks = list(range(0, last_tick + 1, step))

            def _fmt_tick(v: float) -> str:
                if abs(v - round(v)) < 1e-6:
                    return str(int(round(v)))
                return f"{v:.1f}"

            self.ax_capacity.set_yticks(ticks)
            self.ax_capacity.set_yticklabels([_fmt_tick(t) for t in ticks])
        except Exception:
            pass

        # Legend below chart, horizontal, centered (explicit handles so 'Drying' appears once)
        handles = [Patch(facecolor='tab:blue', label='Drying'), Patch(facecolor='tab:orange', label='Cleaning'), Patch(facecolor='tab:green', alpha=0.6, label='Remaining')]
        self.ax_capacity.legend(handles=handles, loc="upper center", bbox_to_anchor=(0.5, -0.08), 
                                ncol=3, frameon=False, fontsize=9)
        
        # Red dashed reference line:
        # - In Day view show the fixed reference (default 24h)
        # - In Week/Month views use the computed hours_available for that period
        if getattr(self, 'view_mode', None) and self.view_mode.get() == "Day":
            available_line = float(self.controller.reference_hours.get())
        else:
            available_line = hours_available[0] if hours_available else float(self.controller.hours_per_day.get())
        # Reference line (don't clip it)
        self.ax_capacity.axhline(y=available_line, color="red", linestyle="--", alpha=0.5, clip_on=False)
        try:
            if getattr(self, 'view_mode', None) and self.view_mode.get() == "Day":
                # place the label *outside* the axes at the right edge, vertically aligned with the line
                ref_txt = f"{int(float(self.controller.reference_hours.get()))}h"
                # xy=(1.0, available_line) uses axes-fraction for x and data for y
                self.ax_capacity.annotate(ref_txt,
                                          xy=(1.0, available_line), xycoords=('axes fraction', 'data'),
                                          xytext=(6, 0), textcoords='offset points',
                                          va='center', ha='left', color='red', fontsize=9, clip_on=False)
        except Exception:
            pass

        # finalize drawing
        self.canvas_capacity.draw()

    def render_details(self, summaries: List[DryerSummary]):
        # Clear existing widgets
        for widget in self.details_frame.winfo_children():
            widget.destroy()

        # Define columns
        columns = ("dryer", "units", "rate", "drying", "washes", "clean", "total", "left", "util")
        
        # Create Treeview with gridlines
        style = ttk.Style()
        style.configure("Details.Treeview", font=("Consolas", 10), rowheight=24)
        style.configure("Details.Treeview.Heading", font=("Consolas", 10, "bold"))
        
        tree = ttk.Treeview(
            self.details_frame, 
            columns=columns, 
            show="headings",
            style="Details.Treeview"
        )
        
        # Configure headings and column widths
        headings = {
            "dryer": ("Dryer", 60),
            "units": ("KG", 80),
            "rate": ("Rate/Hr", 70),
            "drying": ("Drying", 70),
            "washes": ("Washes", 65),
            "clean": ("Clean", 65),
            "total": ("Total", 70),
            "left": ("Left", 65),
            "util": ("Util", 75)
        }
        
        for col, (heading, width) in headings.items():
            tree.heading(col, text=heading, anchor="center")
            tree.column(col, width=width, anchor="center", minwidth=width)
        
        # Add scrollbar
        scrollbar = ttk.Scrollbar(self.details_frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)
        
        # Configure tag for over-capacity rows
        tree.tag_configure("over_capacity", background="#ffcccc")
        
        # Insert data rows
        for s in summaries:
            tag = "over_capacity" if s.utilization_pct > 100 else ""
            tree.insert("", "end", values=(
                f"D{s.dryer}",
                f"{s.total_units:,.0f}",
                f"{s.capacity_per_hour:.0f}",
                f"{s.hours_for_drying:.1f}h",
                f"{s.wash_count}",
                f"{s.hours_for_wash:.1f}h",
                f"{s.hours_needed:.1f}h",
                f"{s.hours_remaining:.1f}h",
                f"{s.utilization_pct:.1f}%"
            ), tags=(tag,))
        
        # Allow double-click to view task audit for a dryer
        tree.bind("<Double-1>", lambda e, tr=tree: self._on_details_double_click(e, tr))
        hint = tk.Label(self.details_frame, text="Double-click a row to view dryer audit", font=("Segoe UI", 8), bg="white", fg="#6b7280")
        hint.pack(side="bottom", anchor="w", padx=8, pady=(6, 0))

        # Pack widgets
        tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

    def _on_details_double_click(self, event, tree):
        """Handle double-click on details row and open dryer audit."""
        item = tree.focus()
        if not item:
            return
        vals = tree.item(item, "values")
        if not vals:
            return
        dryer_label = vals[0]  # e.g. 'D11'
        import re
        m = re.search(r"\d+", str(dryer_label))
        if not m:
            return
        dryer = int(m.group())
        self.open_audit_window(dryer)

    def _on_capacity_pick(self, event):
        """Matplotlib pick-event handler for capacity bars — open audit on double-click."""
        mouse = getattr(event, 'mouseevent', None)
        # Require a double-click to open the audit (avoids accidental single-clicks)
        if not mouse or not getattr(mouse, 'dblclick', False):
            return
        artist = getattr(event, 'artist', None)
        if not artist:
            return
        # Try gid first, then fallback to attribute
        dryer = None
        try:
            gid = artist.get_gid()
            if gid:
                dryer = int(gid)
        except Exception:
            dryer = getattr(artist, '_dryer', None)
        if dryer:
            self.open_audit_window(dryer)


    def open_audit_window(self, dryer: int):
        """Show a detailed list of tasks (material + description + qty + date) for a dryer.

        - Prevent duplicate windows (bring existing to front).
        - Support column sorting on Order, Quantity, Date (priority always on top).
        """
        # If audit for this dryer already open, focus it
        existing = self._open_audits.get(dryer)
        if existing and tk.Toplevel.winfo_exists(existing):
            try:
                existing.lift()
                existing.focus_force()
            except Exception:
                pass
            return

        summary = next((s for s in self.summaries if s.dryer == dryer), None)
        tasks = summary.tasks if summary else []
        # Initial display order: priority first, then heaviest -> lightest
        try:
            tasks = sorted(tasks, key=lambda t: (not bool(getattr(t, 'priority', False)), -float(getattr(t, 'qty', 0))))
        except Exception:
            pass

        win = tk.Toplevel(self)
        win.title(f"Audit - Dryer {dryer}")
        # remember this window to avoid duplicates
        self._open_audits[dryer] = win
        win.protocol("WM_DELETE_WINDOW", lambda d=dryer, w=win: (w.destroy(), self._open_audits.pop(d, None)))

        # Wider audit window so footer / table columns do not get clipped
        win.geometry("1200x460")
        win.transient(self)
        win.grab_set()
        try:
            if os.path.exists(ICON_PATH):
                win.iconbitmap(ICON_PATH)
        except Exception:
            pass

        header = tk.Frame(win, bg="white")
        header.grid(row=0, column=0, columnspan=2, sticky="ew", padx=12, pady=8)
        tk.Label(header, text=f"Dryer {dryer} — {len(tasks)} task(s)", font=("Segoe UI", 11, "bold"), bg="white").pack(anchor="w")
        # small legend for priority
        tk.Label(header, text="* = Priority", font=("Segoe UI", 9), bg="white", fg="#b45309").pack(anchor="w")

        cols = ("order", "material", "desc", "qty", "runhrs", "date")
        # Style the audit tree so headings are bold
        style = ttk.Style(win)
        style.configure("Audit.Treeview", font=("Segoe UI", 10))
        style.configure("Audit.Treeview.Heading", font=("Segoe UI", 10, "bold"))

        tree = ttk.Treeview(win, columns=cols, show="headings", height=12, style="Audit.Treeview")
        # Headings; Order, Quantity and Date are clickable to sort (priority always remains on top)
        tree.heading("order", text="Order #", anchor="center")
        tree.heading("material", text="Material No", anchor="center")
        tree.heading("desc", text="Description", anchor="center")
        tree.heading("qty", text="Quantity", anchor="center")
        tree.heading("runhrs", text="Run Hrs", anchor="center")
        tree.heading("date", text="Date", anchor="center")
        tree.column("order", width=110, anchor="center")
        tree.column("material", width=120, anchor="center")
        tree.column("desc", width=520, anchor="center")
        tree.column("qty", width=100, anchor="center")
        tree.column("runhrs", width=90, anchor="center")
        tree.column("date", width=200, anchor="center")

        # Sorting helpers for this audit tree
        heading_labels = {"order": "Order #", "material": "Material No", "desc": "Description", "qty": "Quantity", "runhrs": "Run Hrs", "date": "Date"}
        tree._sort_state = {"col": None, "desc": True}

        def _refresh_tree(items: list):
            # clear and re-insert rows from Task objects
            tree.delete(*tree.get_children())
            for tt in items:
                runhrs = task_to_hours.get(id(tt), 0.0)
                runhrs_str = f"{runhrs:.1f}h" if runhrs else ""
                order_display = format_order_label(getattr(tt, 'order', None)) or ""
                row_tags = ()
                if getattr(tt, 'priority', False):
                    order_display = f"*{order_display}" if order_display else "*"
                    row_tags = ("priority",)
                tree.insert("", "end", values=(order_display, format_material_label(tt.material) or "", tt.material_desc or "", f"{tt.qty:,.2f}", runhrs_str, tt.date or ""), tags=row_tags)

        def _sort_by(col: str):
            # toggle behavior: first click sorts descending (big->small), second click toggles
            state = tree._sort_state
            if state.get("col") == col:
                state["desc"] = not state.get("desc", True)
            else:
                state["col"] = col
                state["desc"] = True
            # determine key function
            from datetime import datetime
            def _parse_date(s: Optional[str]):
                if not s:
                    return datetime.min
                try:
                    return datetime.strptime(s, "%Y-%m-%d")
                except Exception:
                    try:
                        return datetime.strptime(s, "%m/%d/%Y")
                    except Exception:
                        return datetime.min

            if col == "order":
                def keyfn(t: Task):
                    s = format_order_label(getattr(t, 'order', None)) or ""
                    try:
                        return int(s)
                    except Exception:
                        return s.lower()
            elif col == "qty":
                def keyfn(t: Task):
                    try:
                        return float(getattr(t, 'qty', 0.0))
                    except Exception:
                        return 0.0
            elif col == "date":
                def keyfn(t: Task):
                    return _parse_date(getattr(t, 'date', None))
            elif col == "runhrs":
                # sort by computed run hours (qty / capacity)
                def keyfn(t: Task):
                    try:
                        return float(task_to_hours.get(id(t), 0.0))
                    except Exception:
                        try:
                            # fallback: compute from qty/capacity
                            cap = float(self.controller.dryer_capacities.get(dryer, 0.0))
                            return float(getattr(t, 'qty', 0.0)) / cap if cap > 0 else 0.0
                        except Exception:
                            return 0.0
            elif col == "material":
                def keyfn(t: Task):
                    m = format_material_label(getattr(t, 'material', None)) or ""
                    # try numeric comparison when material is numeric
                    try:
                        return int(m)
                    except Exception:
                        return m.lower()
            elif col == "desc":
                def keyfn(t: Task):
                    return (getattr(t, 'material_desc', '') or '').lower()
            else:
                # fallback: sort by material string
                def keyfn(t: Task):
                    return (getattr(t, 'material', '') or '').lower()

            # apply priority-first, then sort within groups
            pri = [t for t in tasks if getattr(t, 'priority', False)]
            nonpri = [t for t in tasks if not getattr(t, 'priority', False)]
            pri_sorted = sorted(pri, key=keyfn, reverse=state["desc"])
            nonpri_sorted = sorted(nonpri, key=keyfn, reverse=state["desc"])
            new_order = pri_sorted + nonpri_sorted
            # update the in-scope tasks order so subsequent sorts operate on the same list
            tasks[:] = new_order
            _refresh_tree(new_order)

            # update header labels with sort indicator (only one active at a time)
            for c, lbl in heading_labels.items():
                suffix = ''
                if c == state.get('col'):
                    suffix = ' ↓' if state.get('desc') else ' ↑'
                tree.heading(c, text=lbl + suffix)

        # attach header click handlers for sortable columns
        tree.heading("order", command=lambda c="order": _sort_by(c))
        tree.heading("qty", command=lambda c="qty": _sort_by(c))
        tree.heading("date", command=lambda c="date": _sort_by(c))
        tree.heading("runhrs", command=lambda c="runhrs": _sort_by(c))
        tree.heading("material", command=lambda c="material": _sort_by(c))
        tree.heading("desc", command=lambda c="desc": _sort_by(c))

        # Highlight priority rows
        tree.tag_configure("priority", background="#fff4e5", foreground="#7a4100")

        vsb = ttk.Scrollbar(win, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=vsb.set)

        # Compute per-task run hours for display (task.qty / capacity)
        capacity = self.controller.dryer_capacities.get(dryer, 75.0)
        task_to_hours = {}
        for tt in tasks:
            try:
                th = float(tt.qty) / capacity if capacity > 0 else 0.0
            except Exception:
                th = 0.0
            task_to_hours[id(tt)] = th

        for t in tasks:
            runhrs = task_to_hours.get(id(t), 0.0)
            runhrs_str = f"{runhrs:.1f}h" if runhrs else ""
            order_display = format_order_label(getattr(t, 'order', None)) or ""
            row_tags = ()
            if getattr(t, 'priority', False):
                # prepend a star and apply priority tag
                order_display = f"*{order_display}" if order_display else "*"
                row_tags = ("priority",)
            tree.insert("", "end", values=(order_display, format_material_label(t.material) or "", t.material_desc or "", f"{t.qty:,.2f}", runhrs_str, t.date or ""), tags=row_tags)

        # Use grid so footer is always visible and scrollbar aligns correctly
        tree.grid(row=1, column=0, sticky="nsew", padx=(12, 0), pady=(8, 12))
        vsb.grid(row=1, column=1, sticky="ns", padx=(0, 12), pady=(8, 12))
        # Horizontal scrollbar for wide tables
        hsb = ttk.Scrollbar(win, orient="horizontal", command=tree.xview)
        tree.configure(xscrollcommand=hsb.set)
        hsb.grid(row=2, column=0, columnspan=2, sticky="ew", padx=(12, 12), pady=(0, 6))

        win.grid_rowconfigure(1, weight=1)
        win.grid_columnconfigure(0, weight=1)

        # Footer: show wash summary (count + total wash time) + hours needed + utilization
        wash_count = summary.wash_count if summary else count_washes_for_dryer(tasks, mode=self.controller.wash_count_mode.get())
        dryer_wash_hours = self.controller.dryer_wash_hours.get(dryer, 2.0)
        total_wash_time = wash_count * dryer_wash_hours

        # Determine hours needed and utilization (fall back to calculated values if summary missing)
        if summary:
            hours_needed = summary.hours_needed
            utilization_pct = summary.utilization_pct
            hours_available = summary.hours_available
        else:
            total_units = sum(t.qty for t in tasks)
            capacity = self.controller.dryer_capacities.get(dryer, 75.0)
            hours_for_drying = total_units / capacity if capacity > 0 else 0.0
            hours_needed = hours_for_drying + total_wash_time
            hours_available = float(self.controller.hours_per_day.get())
            utilization_pct = (hours_needed / hours_available * 100) if hours_available > 0 else 0.0

        footer = tk.Frame(win, bg="white")
        footer.grid(row=3, column=0, columnspan=2, sticky="ew", padx=12, pady=(0, 8))

        # Total KG (left-most)
        total_kg = sum(t.qty for t in tasks)
        # Show exact quantity (two decimals) instead of rounding
        totalkg_lbl = tk.Label(footer, text=f"Total KG: {total_kg:,.2f}",
                               font=("Segoe UI", 10, "bold"), bg="#F3F4F6", fg="#111827", bd=1, relief="solid", padx=8, pady=4)
        totalkg_lbl.pack(side="left", padx=(0, 8))

        # Remaining KG capacity = hours_remaining * capacity (don't go negative)
        capacity = self.controller.dryer_capacities.get(dryer, 75.0)
        if summary:
            hours_remaining_val = summary.hours_remaining
        else:
            hours_remaining_val = max(0.0, hours_available - hours_needed)
        remaining_kg = max(0.0, hours_remaining_val) * capacity
        remainingkg_lbl = tk.Label(footer, text=f"Remaining KG: {remaining_kg:,.2f}",
                                   font=("Segoe UI", 10), bg="#EEF2FF", fg="#0B3D91", bd=1, relief="solid", padx=8, pady=4)
        remainingkg_lbl.pack(side="left", padx=(0, 8))

        # Left: yellow wash info
        wash_lbl = tk.Label(footer, text=f"Washes: {wash_count}    Total wash time: {total_wash_time:.1f}h",
                            font=("Segoe UI", 10), bg="#FDE68A", fg="#111827", bd=1, relief="solid", padx=8, pady=4)
        wash_lbl.pack(side="left", padx=(0, 8))

        # Middle: hours needed (green)
        need_lbl = tk.Label(footer, text=f"Hours needed: {hours_needed:.1f}h",
                            font=("Segoe UI", 10, "bold"), bg="#10B981", fg="white", bd=1, relief="solid", padx=10, pady=4)
        need_lbl.pack(side="left", padx=(0, 8))

        # Right: utilization (red) and Available (blue)
        util_lbl = tk.Label(footer, text=f"Utilization: {utilization_pct:.1f}%",
                            font=("Segoe UI", 10, "bold"), bg="#EF4444", fg="white", bd=1, relief="solid", padx=10, pady=4)
        util_lbl.pack(side="right")

        remaining_lbl = tk.Label(footer, text=f"Hours remaining: {hours_remaining_val:.1f}h",
                                 font=("Segoe UI", 10, "bold"), bg="#60A5FA", fg="white", bd=1, relief="solid", padx=10, pady=4)
        remaining_lbl.pack(side="right", padx=(0, 8))

    def open_settings(self):
        win = tk.Toplevel(self)
        win.title("Settings")
        # Make the settings dialog larger by default and allow the user to resize it
        win.geometry("760x760")
        win.minsize(530, 540)
        win.transient(self)
        win.grab_set()
        win.configure(bg="#f5f5f5")
        win.resizable(True, True)

        try:
            if os.path.exists(ICON_PATH):
                win.iconbitmap(ICON_PATH)
        except Exception:
            pass

        # Modern styling
        style = ttk.Style()
        style.configure("Card.TFrame", background="white")
        style.configure("CardLabel.TLabel", background="white", font=("Segoe UI", 10))
        style.configure("CardHeader.TLabel", background="white", font=("Segoe UI", 11, "bold"))
        style.configure("Settings.TEntry", padding=5)
        style.configure("Primary.TButton", font=("Segoe UI", 10))
        
        # Main container with padding (scrollable)
        container = tk.Frame(win, bg="#f5f5f5")
        container.pack(fill="both", expand=True)

        settings_canvas = tk.Canvas(container, bg="#f5f5f5", highlightthickness=0)
        settings_vsb = ttk.Scrollbar(container, orient="vertical", command=settings_canvas.yview)
        settings_canvas.configure(yscrollcommand=settings_vsb.set)
        settings_vsb.pack(side="right", fill="y")
        settings_canvas.pack(side="left", fill="both", expand=True)

        # Inner frame that holds the actual settings content (keeps the original name)
        main_frame = tk.Frame(settings_canvas, bg="#f5f5f5")
        window_id = settings_canvas.create_window((0, 0), window=main_frame, anchor="nw")

        def _on_main_config(e):
            settings_canvas.configure(scrollregion=settings_canvas.bbox("all"))
        main_frame.bind("<Configure>", _on_main_config)

        def _on_canvas_config(e):
            # keep inner frame width matched to canvas width
            settings_canvas.itemconfigure(window_id, width=e.width)
        settings_canvas.bind("<Configure>", _on_canvas_config)

        # Mousewheel support for the settings canvas
        def _on_mousewheel(e):
            settings_canvas.yview_scroll(int(-1 * (e.delta / 120)), "units")
        settings_canvas.bind_all("<MouseWheel>", _on_mousewheel)

        # Add padding inside the inner frame
        main_frame.configure(padx=20, pady=20)

        # ---- Hours Per Day Card ----
        hours_card = tk.Frame(main_frame, bg="white", highlightbackground="#e0e0e0", highlightthickness=1)
        hours_card.pack(fill="x", pady=(0, 12))
        
        hours_inner = tk.Frame(hours_card, bg="white")
        hours_inner.pack(fill="x", padx=20, pady=16)
        
        tk.Label(hours_inner, text="Hours Available Per Day", font=("Segoe UI", 11, "bold"), 
                 bg="white", fg="#333333").pack(anchor="w")
        tk.Label(hours_inner, text="Set the total working hours available each day", 
                 font=("Segoe UI", 9), bg="white", fg="#666666").pack(anchor="w", pady=(2, 10))
        
        hours_entry_frame = tk.Frame(hours_inner, bg="white")
        hours_entry_frame.pack(anchor="w")
        hours_entry = ttk.Entry(hours_entry_frame, width=12, justify="center", font=("Segoe UI", 11))
        hours_entry.insert(0, str(self.controller.hours_per_day.get()))
        hours_entry.pack(side="left")
        tk.Label(hours_entry_frame, text="hours", font=("Segoe UI", 10), bg="white", fg="#666666").pack(side="left", padx=(8, 0))

        # Chart maximum (Y-axis) — lets user expand the graph without changing calculations
        chart_frame = tk.Frame(hours_inner, bg="white")
        chart_frame.pack(anchor="w", pady=(8, 0))
        tk.Label(chart_frame, text="Chart max (Y‑axis)", font=("Segoe UI", 9), bg="white", fg="#666666").pack(side="left")
        chart_entry = ttk.Entry(chart_frame, width=8, justify="center", font=("Segoe UI", 10))
        chart_entry.insert(0, str(self.controller.chart_max_hours.get()))
        chart_entry.pack(side="left", padx=(8, 0))
        tk.Label(chart_frame, text="hours", font=("Segoe UI", 9), bg="white", fg="#666666").pack(side="left", padx=(8, 0))
        tk.Label(hours_inner, text="(Y-axis maximum for the capacity chart — reference line stays at 24h)",
                 font=("Segoe UI", 8), bg="white", fg="#999999").pack(anchor="w", pady=(6, 0))

        # ---- Work Days Per Week Card ----
        week_card = tk.Frame(main_frame, bg="white", highlightbackground="#e0e0e0", highlightthickness=1)
        week_card.pack(fill="x", pady=(0, 12))
        
        week_inner = tk.Frame(week_card, bg="white")
        week_inner.pack(fill="x", padx=20, pady=16)
        
        tk.Label(week_inner, text="Work Days Per Week", font=("Segoe UI", 11, "bold"), 
                 bg="white", fg="#333333").pack(anchor="w")
        tk.Label(week_inner, text="Number of work days when viewing weekly capacity", 
                 font=("Segoe UI", 9), bg="white", fg="#666666").pack(anchor="w", pady=(2, 10))
        
        week_entry_frame = tk.Frame(week_inner, bg="white")
        week_entry_frame.pack(anchor="w")
        week_entry = ttk.Entry(week_entry_frame, width=12, justify="center", font=("Segoe UI", 11))
        week_entry.insert(0, str(self.controller.work_days_per_week.get()))
        week_entry.pack(side="left")
        tk.Label(week_entry_frame, text="days", font=("Segoe UI", 10), bg="white", fg="#666666").pack(side="left", padx=(8, 0))



        # ---- Dryer Settings Card (Capacity + Wash Time) ----
        dryer_card = tk.Frame(main_frame, bg="white", highlightbackground="#e0e0e0", highlightthickness=1)
        dryer_card.pack(fill="x")
        
        dryer_header = tk.Frame(dryer_card, bg="white")
        dryer_header.pack(fill="x", padx=20, pady=(16, 0))
        
        tk.Label(dryer_header, text="Dryer Settings", font=("Segoe UI", 11, "bold"), 
                 bg="white", fg="#333333").pack(anchor="w")
        tk.Label(dryer_header, text="Capacity (kg/hr) and wash time (hours) for each dryer", 
                 font=("Segoe UI", 9), bg="white", fg="#666666").pack(anchor="w", pady=(2, 0))

        # Column headers
        header_frame = tk.Frame(dryer_card, bg="white")
        header_frame.pack(fill="x", padx=20, pady=(10, 0))
        
        tk.Label(header_frame, text="Dryer", width=10, anchor="w",
                 font=("Segoe UI", 9, "bold"), bg="white", fg="#666666").pack(side="left")
        tk.Label(header_frame, text="Capacity", width=12, 
                 font=("Segoe UI", 9, "bold"), bg="white", fg="#666666").pack(side="left", padx=(0, 8))
        tk.Label(header_frame, text="Wash Time", width=12, 
                 font=("Segoe UI", 9, "bold"), bg="white", fg="#666666").pack(side="left")

        # Dryer list (scrollable section)
        list_container_outer = tk.Frame(dryer_card, bg="white")
        list_container_outer.pack(fill="both", padx=20, pady=(4, 12), expand=True)

        list_canvas = tk.Canvas(list_container_outer, bg="white", highlightthickness=0, height=220)
        list_vsb = ttk.Scrollbar(list_container_outer, orient="vertical", command=list_canvas.yview)
        list_canvas.configure(yscrollcommand=list_vsb.set)
        list_vsb.pack(side="right", fill="y")
        list_canvas.pack(side="left", fill="both", expand=True)

        list_container = tk.Frame(list_canvas, bg="white")
        list_win_id = list_canvas.create_window((0, 0), window=list_container, anchor="nw")
        def _on_list_config(e):
            list_canvas.configure(scrollregion=list_canvas.bbox("all"))
        list_container.bind("<Configure>", _on_list_config)
        def _on_list_canvas_cfg(e):
            list_canvas.itemconfigure(list_win_id, width=e.width)
        list_canvas.bind("<Configure>", _on_list_canvas_cfg)
        # Allow vertical scrolling with mouse wheel over the dryer list
        list_canvas.bind_all("<MouseWheel>", lambda ev: list_canvas.yview_scroll(int(-1 * (ev.delta / 120)), "units"))

        # Create entries for each dryer
        dryer_entries: Dict[int, ttk.Entry] = {}
        wash_entries: Dict[int, ttk.Entry] = {}
        
        # Always show all 5 dryers
        all_dryers = [2, 6, 9, 10, 11]

        for i, dryer in enumerate(all_dryers):
            row_bg = "#fafafa" if i % 2 == 0 else "white"
            row = tk.Frame(list_container, bg=row_bg)
            row.pack(fill="x", pady=1)
            
            row_inner = tk.Frame(row, bg=row_bg)
            row_inner.pack(fill="x", padx=8, pady=8)
            
            tk.Label(row_inner, text=f"Dryer {dryer}", width=10, anchor="w", 
                     font=("Segoe UI", 10), bg=row_bg, fg="#333333").pack(side="left")
            
            # Capacity entry
            cap_entry = ttk.Entry(row_inner, width=8, justify="center", font=("Segoe UI", 10))
            current_cap = self.controller.dryer_capacities.get(dryer, 75.0)
            cap_entry.insert(0, str(current_cap))
            cap_entry.pack(side="left", padx=(0, 4))
            tk.Label(row_inner, text="kg/hr", font=("Segoe UI", 8), bg=row_bg, fg="#888888").pack(side="left", padx=(0, 12))
            
            # Wash time entry
            wash_entry = ttk.Entry(row_inner, width=8, justify="center", font=("Segoe UI", 10))
            current_wash = self.controller.dryer_wash_hours.get(dryer, 2.0)
            wash_entry.insert(0, str(current_wash))
            wash_entry.pack(side="left", padx=(0, 4))
            tk.Label(row_inner, text="hrs", font=("Segoe UI", 8), bg=row_bg, fg="#888888").pack(side="left")
            
            dryer_entries[dryer] = cap_entry
            wash_entries[dryer] = wash_entry

        # ---- Button Row ----
        btn_frame = tk.Frame(main_frame, bg="#f5f5f5")
        btn_frame.pack(fill="x", pady=(16, 0))

        def apply():
            try:
                hours = float(hours_entry.get())
                if hours <= 0:
                    messagebox.showwarning("Settings", "Hours per day must be greater than 0.")
                    return
                self.controller.hours_per_day.set(hours)
            except ValueError:
                messagebox.showwarning("Settings", "Hours per day must be a number.")
                return

            # Chart max hours (Y-axis)
            try:
                chart_max = float(chart_entry.get())
                if chart_max <= 0:
                    messagebox.showwarning("Settings", "Chart max hours must be greater than 0.")
                    return
                self.controller.chart_max_hours.set(chart_max)
            except ValueError:
                messagebox.showwarning("Settings", "Chart max hours must be a number.")
                return

            try:
                work_days = int(week_entry.get())
                if work_days <= 0 or work_days > 7:
                    messagebox.showwarning("Settings", "Work days per week must be between 1 and 7.")
                    return
                self.controller.work_days_per_week.set(work_days)
            except ValueError:
                messagebox.showwarning("Settings", "Work days per week must be a whole number.")
                return



            for dryer in dryer_entries.keys():
                # Validate capacity
                try:
                    cap = float(dryer_entries[dryer].get())
                    if cap <= 0:
                        messagebox.showwarning("Settings", f"Dryer {dryer} capacity must be positive.")
                        return
                    self.controller.dryer_capacities[dryer] = cap
                except ValueError:
                    messagebox.showwarning("Settings", f"Dryer {dryer} capacity must be a number.")
                    return
                
                # Validate wash time
                try:
                    wash = float(wash_entries[dryer].get())
                    if wash < 0:
                        messagebox.showwarning("Settings", f"Dryer {dryer} wash time cannot be negative.")
                        return
                    self.controller.dryer_wash_hours[dryer] = wash
                except ValueError:
                    messagebox.showwarning("Settings", f"Dryer {dryer} wash time must be a number.")
                    return

            # Persist new settings immediately
            self.controller.save_settings()

            win.destroy()
            self.render_all()

        # Modern styled buttons
        cancel_btn = tk.Button(btn_frame, text="Cancel", width=12, command=win.destroy,
                               font=("Segoe UI", 10), bg="white", fg="#333333", 
                               relief="solid", bd=1, cursor="hand2")
        cancel_btn.pack(side="left")
        
        # add a small link/button for manual update checks
        def _check_updates():
            self.controller.check_for_updates(silent=False)
        upd_btn = tk.Button(btn_frame, text="Check for updates", command=_check_updates,
                            font=("Segoe UI", 9, "underline"), bg="#f5f5f5", fg="#0078d4",
                            relief="flat", bd=0, cursor="hand2", activebackground="#f5f5f5",
                            activeforeground="#106ebe")
        upd_btn.pack(side="left", padx=(8, 0))

        save_btn = tk.Button(btn_frame, text="Save Changes", width=14, command=apply,
                             font=("Segoe UI", 10, "bold"), bg="#0078d4", fg="white",
                             relief="flat", bd=0, cursor="hand2", activebackground="#106ebe",
                             activeforeground="white")
        save_btn.pack(side="right")


# ======================================================================
# MAIN
# ======================================================================
if __name__ == "__main__":
    app = App()
    app.mainloop()
