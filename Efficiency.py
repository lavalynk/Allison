import os
import re
import random
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from dataclasses import dataclass
from datetime import datetime, date, time, timedelta
from typing import List, Optional, Tuple, Dict, Literal

import pandas as pd
import matplotlib

matplotlib.use("TkAgg")
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure
from matplotlib.patches import Patch


# ======================================================================
# CONFIG
# ======================================================================
APP_TITLE = "Dryer Schedule Dashboard"
ICON_PATH = "dryer.ico"  # keep this in the same folder as this script

DEFAULT_WASH_MINUTES = 30
DEFAULT_MINUTES_PER_UNIT = 0.80  # minutes per "Order quantity (GMEIN)" unit

# 24 hour operation: midnight to next midnight
DAY_START_TIME = time(0, 0)

# Optional gap when packing tasks into a 24 hour day (0 means more continuous)
PACK_GAP_MINUTES = 0

# Timeline labels spacing
TIMELINE_LABEL_STEP_MINUTES = 60


# ======================================================================
# EXCEL FORMAT A (optional, true time schedule)
# If file contains these columns, it will use them directly as a timeline.
# ======================================================================
COL_DRYER = "Dryer"
COL_START = "Start"
COL_END = "End"
COL_CODE = "Code"   # in your case you can use Material Number here too
COL_DATE = "Date"   # optional


# ======================================================================
# EXCEL FORMAT B (your SAP export)
# ======================================================================
SAP_PROD_VERSION = "Production Version"        # ex: SD02
SAP_QTY = "Order quantity (GMEIN)"             # numeric
SAP_MATERIAL = "Material Number"               # label and wash key
SAP_START_DATE = "Basic start date"            # date only, optional


# ======================================================================
# DATA TYPES
# ======================================================================
@dataclass
class Batch:
    dryer: int
    start_dt: datetime
    end_dt: datetime
    material: str  # label and wash key


@dataclass
class PackedTask:
    dryer: int
    qty: float
    material: str


LoadMode = Literal["timed", "packed"]


# ======================================================================
# HELPERS
# ======================================================================
def get_day_window(day: date) -> Tuple[datetime, datetime]:
    start_dt = datetime(day.year, day.month, day.day, DAY_START_TIME.hour, DAY_START_TIME.minute, 0)
    end_dt = start_dt + timedelta(days=1)
    return start_dt, end_dt


def minutes_between(a: datetime, b: datetime) -> float:
    return (b - a).total_seconds() / 60.0


def fmt_hour_label(dt_val: datetime) -> str:
    hr = dt_val.hour
    ampm = "a" if hr < 12 else "p"
    hr12 = hr % 12
    if hr12 == 0:
        hr12 = 12
    return f"{hr12}{ampm}"


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


def compute_wash_segments(batches_sorted: List[Batch], wash_minutes: int) -> List[Tuple[datetime, datetime]]:
    """
    Wash occurs after a batch when the next batch has a different material number.
    Wash is clamped so it does not overlap the next batch start.
    """
    segs: List[Tuple[datetime, datetime]] = []
    if wash_minutes <= 0 or len(batches_sorted) < 2:
        return segs

    for i in range(len(batches_sorted) - 1):
        cur = batches_sorted[i]
        nxt = batches_sorted[i + 1]

        cur_mat = (cur.material or "").strip()
        nxt_mat = (nxt.material or "").strip()

        if not cur_mat or not nxt_mat:
            continue

        if cur_mat == nxt_mat:
            continue

        wash_start = cur.end_dt
        wash_end = wash_start + timedelta(minutes=wash_minutes)

        if wash_end > nxt.start_dt:
            wash_end = nxt.start_dt

        if wash_end > wash_start:
            segs.append((wash_start, wash_end))

    return segs


def compute_gaps_for_dryer(batches_sorted: List[Batch], window_start: datetime, window_end: datetime) -> List[Tuple[datetime, datetime]]:
    gaps: List[Tuple[datetime, datetime]] = []
    cur = window_start

    for b in batches_sorted:
        if b.start_dt > cur:
            gaps.append((cur, b.start_dt))
        if b.end_dt > cur:
            cur = b.end_dt

    if cur < window_end:
        gaps.append((cur, window_end))

    return [(a, b) for (a, b) in gaps if b > a]


def group_batches_by_dryer(batches: List[Batch]) -> Dict[int, List[Batch]]:
    grouped: Dict[int, List[Batch]] = {}
    for b in batches:
        grouped.setdefault(b.dryer, []).append(b)
    for d in grouped:
        grouped[d].sort(key=lambda x: x.start_dt)
    return grouped


def group_tasks_by_dryer(tasks: List[PackedTask]) -> Dict[int, List[PackedTask]]:
    grouped: Dict[int, List[PackedTask]] = {}
    for t in tasks:
        grouped.setdefault(t.dryer, []).append(t)
    return grouped


def active_dryers_from_batches(batches: List[Batch]) -> List[int]:
    dryers = sorted({b.dryer for b in batches})
    return dryers


def active_dryers_from_tasks(tasks: List[PackedTask]) -> List[int]:
    dryers = sorted({t.dryer for t in tasks})
    return dryers


def short_label(s: str, max_len: int = 10) -> str:
    s = (s or "").strip()
    if len(s) <= max_len:
        return s
    return s[:max_len - 1] + "…"


# ======================================================================
# MOCK DATA (no Excel selected)
# Generates lots more batches.
# ======================================================================
def build_mock_batches(day: Optional[date] = None) -> List[Batch]:
    if day is None:
        day = date.today()

    day_start, day_end = get_day_window(day)
    day_minutes = int(minutes_between(day_start, day_end))  # 1440

    target_fill = 0.72
    min_batch = 20
    max_batch = 70
    min_gap = 2
    max_gap = 12

    # Fake "material numbers"
    materials = ["10001234", "10004567", "10007890", "10001111", "10002222", "10003333", "10004444"]

    random.seed(day.toordinal())

    batches: List[Batch] = []

    # Mock uses 10 dryers (all active) to stress test the UI
    for dryer in range(1, 11):
        scheduled_target = int(day_minutes * target_fill)
        scheduled_so_far = 0

        t = day_start + timedelta(minutes=random.randint(0, 40))

        while scheduled_so_far < scheduled_target and t < day_end:
            dur = random.randint(min_batch, max_batch)
            end_t = t + timedelta(minutes=dur)
            if end_t > day_end:
                break

            material = random.choice(materials)

            batches.append(Batch(dryer, t, end_t, material))
            scheduled_so_far += dur

            gap = random.randint(min_gap, max_gap)
            t = end_t + timedelta(minutes=gap)

    return batches


# ======================================================================
# EXCEL LOADING
# Supports:
#  - Format A: Dryer/Start/End/Code (true time)
#  - Format B: Production Version + Qty + Material Number (packed sequential)
# ======================================================================
def _parse_dt_with_fallback(val: object, assumed_day: date) -> datetime:
    if pd.isna(val):
        raise ValueError("Found a blank Start or End value.")

    if isinstance(val, datetime):
        return val

    ts = pd.to_datetime(val, errors="coerce")
    if not pd.isna(ts):
        py = ts.to_pydatetime()
        if py.year >= 2000:
            return py
        return datetime(assumed_day.year, assumed_day.month, assumed_day.day, py.hour, py.minute, py.second)

    if isinstance(val, str):
        parts = val.strip().split(":")
        if len(parts) >= 2:
            h = int(parts[0])
            m = int(parts[1])
            return datetime(assumed_day.year, assumed_day.month, assumed_day.day, h, m, 0)

    raise ValueError(f"Could not parse datetime or time value: {val}")


def load_from_excel(path: str) -> Tuple[LoadMode, date, List[Batch], List[PackedTask]]:
    df = pd.read_excel(path)

    # Format A
    has_a = all(c in df.columns for c in [COL_DRYER, COL_START, COL_END, COL_CODE])
    if has_a:
        if COL_DATE in df.columns and df[COL_DATE].notna().any():
            day = pd.to_datetime(df[COL_DATE].dropna().iloc[0]).date()
        else:
            day = pd.to_datetime(df[COL_START].dropna().iloc[0]).date()

        batches: List[Batch] = []
        for _, r in df.iterrows():
            if pd.isna(r.get(COL_DRYER)) or pd.isna(r.get(COL_START)) or pd.isna(r.get(COL_END)):
                continue

            try:
                dryer = int(r.get(COL_DRYER))
            except Exception:
                continue

            start_dt = _parse_dt_with_fallback(r.get(COL_START), day)
            end_dt = _parse_dt_with_fallback(r.get(COL_END), day)
            if end_dt <= start_dt:
                end_dt = end_dt + timedelta(days=1)

            material = "" if pd.isna(r.get(COL_CODE)) else str(r.get(COL_CODE)).strip()
            if not material:
                material = "UNKNOWN"

            batches.append(Batch(dryer, start_dt, end_dt, material))

        return "timed", day, batches, []

    # Format B
    required_b = [SAP_PROD_VERSION, SAP_QTY, SAP_MATERIAL]
    missing_b = [c for c in required_b if c not in df.columns]
    if missing_b:
        raise ValueError(
            "Excel format not recognized.\n"
            f"Missing columns: {missing_b}\n"
            f"Found columns: {list(df.columns)}"
        )

    if SAP_START_DATE in df.columns and df[SAP_START_DATE].notna().any():
        day = pd.to_datetime(df[SAP_START_DATE].dropna().iloc[0]).date()
    else:
        day = date.today()

    # Filter to meaningful rows if possible
    if "Order" in df.columns:
        df2 = df[df["Order"].notna()].copy()
    else:
        df2 = df[df[SAP_QTY].notna()].copy()

    tasks: List[PackedTask] = []
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

        tasks.append(PackedTask(dryer=dryer, qty=qty, material=material))

    return "packed", day, [], tasks


def build_batches_from_tasks(day: date, tasks: List[PackedTask], minutes_per_unit: float) -> List[Batch]:
    """
    Packs tasks sequentially per dryer from midnight to midnight.
    Duration = qty * minutes_per_unit (minimum 5 minutes).
    """
    day_start, day_end = get_day_window(day)
    tasks_by_dryer = group_tasks_by_dryer(tasks)

    batches: List[Batch] = []

    for dryer in sorted(tasks_by_dryer.keys()):
        t = day_start
        for task in tasks_by_dryer.get(dryer, []):
            dur = int(round(task.qty * minutes_per_unit))
            if dur < 5:
                dur = 5

            end_t = t + timedelta(minutes=dur)
            if end_t > day_end:
                break

            batches.append(Batch(dryer=dryer, start_dt=t, end_dt=end_t, material=task.material))
            t = end_t + timedelta(minutes=PACK_GAP_MINUTES)

    return batches


# ======================================================================
# GUI
# ======================================================================
class App(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title(APP_TITLE)
        self.geometry("1120x760")

        try:
            if os.path.exists(ICON_PATH):
                self.iconbitmap(ICON_PATH)
        except Exception:
            pass

        self.file_path: Optional[str] = None

        self.container = tk.Frame(self)
        self.container.pack(fill="both", expand=True)

        self.frames: Dict[str, tk.Frame] = {}
        for F in (StartPage, DashboardPage):
            frame = F(parent=self.container, controller=self)
            self.frames[F.__name__] = frame
            frame.grid(row=0, column=0, sticky="nsew")

        self.show_frame("StartPage")

    def show_frame(self, name: str):
        self.frames[name].tkraise()

    def run_execute(self):
        try:
            dash: DashboardPage = self.frames["DashboardPage"]  # type: ignore

            if self.file_path:
                mode, day, timed_batches, packed_tasks = load_from_excel(self.file_path)
                dash.load_data(mode=mode, day=day, timed_batches=timed_batches, packed_tasks=packed_tasks)
            else:
                day = date.today()
                batches = build_mock_batches(day)
                dash.load_data(mode="timed", day=day, timed_batches=batches, packed_tasks=[])

            self.show_frame("DashboardPage")

        except Exception as e:
            messagebox.showerror("Error", str(e))


class StartPage(tk.Frame):
    def __init__(self, parent, controller: App):
        super().__init__(parent)
        self.controller = controller

        title = tk.Label(self, text=APP_TITLE, font=("Segoe UI", 18, "bold"))
        title.pack(pady=20)

        instructions = tk.Label(
            self,
            text="Select an Excel file and click Execute.\nIf you do not select a file, the app will run using mock data.",
            font=("Segoe UI", 11),
            justify="center",
        )
        instructions.pack(pady=10)

        self.path_label = tk.Label(self, text="No file selected", font=("Segoe UI", 10))
        self.path_label.pack(pady=10)

        btn_row = tk.Frame(self)
        btn_row.pack(pady=10)

        select_btn = tk.Button(btn_row, text="Select Excel File", width=20, command=self.select_file)
        select_btn.grid(row=0, column=0, padx=8)

        exec_btn = tk.Button(btn_row, text="Execute", width=20, command=self.controller.run_execute)
        exec_btn.grid(row=0, column=1, padx=8)

        note = tk.Label(
            self,
            text="Excel supported formats:\n"
                 "1) Dryer, Start, End, Code\n"
                 "2) Production Version (SD02), Order quantity (GMEIN), Material Number",
            font=("Segoe UI", 10),
            justify="center",
        )
        note.pack(pady=12)

    def select_file(self):
        path = filedialog.askopenfilename(
            title="Select Excel file",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")],
        )
        if path:
            self.controller.file_path = path
            self.path_label.config(text=path)


class DashboardPage(tk.Frame):
    def __init__(self, parent, controller: App):
        super().__init__(parent)
        self.controller = controller

        self.wash_minutes = tk.IntVar(value=DEFAULT_WASH_MINUTES)
        self.minutes_per_unit = tk.DoubleVar(value=DEFAULT_MINUTES_PER_UNIT)

        self.mode: LoadMode = "timed"
        self.day: date = date.today()
        self.timed_batches: List[Batch] = []
        self.packed_tasks: List[PackedTask] = []

        self.window_start_dt, self.window_end_dt = get_day_window(self.day)

        # Active dryers and y mapping for click handling
        self.active_dryers: List[int] = []
        self.y_to_dryer: Dict[int, int] = {}

        top = tk.Frame(self)
        top.pack(fill="x", padx=12, pady=8)

        back_btn = tk.Button(top, text="Back", command=lambda: controller.show_frame("StartPage"))
        back_btn.pack(side="left")

        settings_btn = tk.Button(top, text="Settings", command=self.open_settings)
        settings_btn.pack(side="right")

        self.summary_label = tk.Label(self, text="", font=("Segoe UI", 11))
        self.summary_label.pack(fill="x", padx=12, pady=6)

        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True, padx=12, pady=10)

        self.tab_timeline = tk.Frame(self.notebook)
        self.tab_totals = tk.Frame(self.notebook)

        self.notebook.add(self.tab_timeline, text="Timeline")
        self.notebook.add(self.tab_totals, text="Totals (minutes)")

        self.fig_timeline = Figure(figsize=(10, 5.6), dpi=100)
        self.ax_timeline = self.fig_timeline.add_subplot(111)
        self.canvas_timeline = FigureCanvasTkAgg(self.fig_timeline, master=self.tab_timeline)
        self.canvas_timeline.get_tk_widget().pack(fill="both", expand=True)

        self.fig_timeline.canvas.mpl_connect("button_press_event", self.on_timeline_click)

        self.fig_totals = Figure(figsize=(10, 5.6), dpi=100)
        self.ax_totals = self.fig_totals.add_subplot(111)
        self.canvas_totals = FigureCanvasTkAgg(self.fig_totals, master=self.tab_totals)
        self.canvas_totals.get_tk_widget().pack(fill="both", expand=True)

    def load_data(self, mode: LoadMode, day: date, timed_batches: List[Batch], packed_tasks: List[PackedTask]):
        self.mode = mode
        self.day = day
        self.timed_batches = timed_batches
        self.packed_tasks = packed_tasks
        self.window_start_dt, self.window_end_dt = get_day_window(self.day)
        self.render_all()

    def current_batches(self) -> List[Batch]:
        if self.mode == "timed":
            return self.timed_batches
        return build_batches_from_tasks(self.day, self.packed_tasks, float(self.minutes_per_unit.get()))

    def open_settings(self):
        win = tk.Toplevel(self)
        win.title("Settings")
        win.geometry("380x220")
        win.transient(self)
        win.grab_set()

        try:
            if os.path.exists(ICON_PATH):
                win.iconbitmap(ICON_PATH)
        except Exception:
            pass

        tk.Label(win, text="Wash time (minutes):", font=("Segoe UI", 10)).pack(pady=(16, 6))
        tk.Entry(win, textvariable=self.wash_minutes, width=12, justify="center").pack()

        tk.Label(win, text="Minutes per unit (GMEIN):", font=("Segoe UI", 10)).pack(pady=(14, 6))
        tk.Entry(win, textvariable=self.minutes_per_unit, width=12, justify="center").pack()

        hint = tk.Label(
            win,
            text="Wash is added when the next batch Material Number is different.",
            font=("Segoe UI", 9),
            justify="center",
        )
        hint.pack(pady=(12, 0))

        def apply():
            try:
                w = int(self.wash_minutes.get())
                if w < 0 or w > 240:
                    messagebox.showwarning("Settings", "Wash time must be between 0 and 240.")
                    return
            except Exception:
                messagebox.showwarning("Settings", "Wash time must be a whole number.")
                return

            try:
                mpu = float(self.minutes_per_unit.get())
                if mpu <= 0 or mpu > 20:
                    messagebox.showwarning("Settings", "Minutes per unit must be between 0 and 20.")
                    return
            except Exception:
                messagebox.showwarning("Settings", "Minutes per unit must be a number.")
                return

            win.destroy()
            self.render_all()

        btn_row = tk.Frame(win)
        btn_row.pack(pady=16)

        tk.Button(btn_row, text="Cancel", width=10, command=win.destroy).grid(row=0, column=0, padx=8)
        tk.Button(btn_row, text="Apply", width=10, command=apply).grid(row=0, column=1, padx=8)

    def render_all(self):
        batches = self.current_batches()

        # Only show dryers that have at least one batch
        self.active_dryers = active_dryers_from_batches(batches)
        if not self.active_dryers:
            self.summary_label.config(text=f"Day: {self.day}   No batches found.")
            self.ax_timeline.clear()
            self.ax_totals.clear()
            self.canvas_timeline.draw()
            self.canvas_totals.draw()
            return

        self.render_timeline(batches)
        self.render_totals(batches)

    def render_timeline(self, batches: List[Batch]):
        self.ax_timeline.clear()

        grouped = group_batches_by_dryer(batches)
        wash_mins = int(self.wash_minutes.get())

        row_height_full = 0.78
        row_height_idle = 0.22

        total_scheduled = 0.0
        total_wash = 0.0

        # Map active dryers to y positions 1..N
        self.y_to_dryer = {}
        dryer_to_y: Dict[int, int] = {}
        for idx, dryer in enumerate(self.active_dryers, start=1):
            dryer_to_y[dryer] = idx
            self.y_to_dryer[idx] = dryer

        for dryer in self.active_dryers:
            y = dryer_to_y[dryer]
            b_list = sorted(grouped.get(dryer, []), key=lambda x: x.start_dt)

            # Scheduled
            for b in b_list:
                seg_start = max(b.start_dt, self.window_start_dt)
                seg_end = min(b.end_dt, self.window_end_dt)
                if seg_end <= seg_start:
                    continue

                start_min = minutes_between(self.window_start_dt, seg_start)
                dur_min = minutes_between(seg_start, seg_end)
                if dur_min <= 0:
                    continue

                total_scheduled += dur_min
                self.ax_timeline.broken_barh(
                    [(start_min, dur_min)],
                    (y - row_height_full / 2, row_height_full),
                    facecolors="tab:blue",
                )

                mat = (b.material or "").strip()
                if mat and dur_min >= 25:
                    self.ax_timeline.text(
                        start_min + 2,
                        y,
                        short_label(mat, 10),
                        va="center",
                        fontsize=9,
                        fontweight="bold",
                        color="white",
                    )

            # Wash
            wash_segs = compute_wash_segments(b_list, wash_mins)
            for (ws, we) in wash_segs:
                seg_start = max(ws, self.window_start_dt)
                seg_end = min(we, self.window_end_dt)
                if seg_end <= seg_start:
                    continue

                start_min = minutes_between(self.window_start_dt, seg_start)
                dur_min = minutes_between(seg_start, seg_end)
                if dur_min <= 0:
                    continue

                total_wash += dur_min
                self.ax_timeline.broken_barh(
                    [(start_min, dur_min)],
                    (y - row_height_full / 2, row_height_full),
                    facecolors="tab:orange",
                )

            # Idle (thin background)
            gaps = compute_gaps_for_dryer(b_list, self.window_start_dt, self.window_end_dt)
            for (gs, ge) in gaps:
                start_min = minutes_between(self.window_start_dt, gs)
                dur_min = minutes_between(gs, ge)
                if dur_min <= 0:
                    continue
                self.ax_timeline.broken_barh(
                    [(start_min, dur_min)],
                    (y - row_height_idle / 2, row_height_idle),
                    facecolors="tab:gray",
                    alpha=0.18,
                )

        # Summary math based on active dryers only
        day_minutes_per_dryer = minutes_between(self.window_start_dt, self.window_end_dt)  # 1440
        total_available = day_minutes_per_dryer * len(self.active_dryers)
        used = total_scheduled + total_wash
        remaining = max(0.0, total_available - used)
        fill = (used / total_available) if total_available > 0 else 0.0

        self.summary_label.config(
            text=(
                f"Day: {self.day}   Scheduled: {total_scheduled:.0f} min   "
                f"Wash: {total_wash:.0f} min   Remaining: {remaining:.0f} min   "
                f"Fill: {fill * 100:.1f}%   Dryers shown: {len(self.active_dryers)}"
            )
        )

        # Axes formatting
        n = len(self.active_dryers)
        self.ax_timeline.set_ylim(0.5, n + 0.5)

        yticks = list(range(1, n + 1))
        ylabels = [f"Dryer {d}" for d in self.active_dryers]
        self.ax_timeline.set_yticks(yticks)
        self.ax_timeline.set_yticklabels(ylabels)

        x_max = day_minutes_per_dryer
        self.ax_timeline.set_xlim(0, x_max)

        ticks, labels = [], []
        step = int(TIMELINE_LABEL_STEP_MINUTES)
        for m in range(0, int(x_max) + 1, step):
            dt_val = self.window_start_dt + timedelta(minutes=m)
            ticks.append(m)
            labels.append(fmt_hour_label(dt_val))

        self.ax_timeline.set_xticks(ticks)
        self.ax_timeline.set_xticklabels(labels)

        self.ax_timeline.set_title("Dryer Schedule Timeline (click a dryer row for details)")
        self.ax_timeline.set_xlabel("Time")
        self.ax_timeline.set_ylabel("")

        legend_items = [
            Patch(facecolor="tab:blue", label="Scheduled"),
            Patch(facecolor="tab:orange", label=f"Wash ({wash_mins} min)"),
            Patch(facecolor="tab:gray", label="Idle"),
        ]

        self.fig_timeline.subplots_adjust(bottom=0.18)
        self.ax_timeline.legend(
            handles=legend_items,
            loc="upper right",
            bbox_to_anchor=(1.0, -0.12),
            frameon=True,
            ncol=1
        )

        self.fig_timeline.tight_layout()
        self.canvas_timeline.draw()

    def render_totals(self, batches: List[Batch]):
        self.ax_totals.clear()

        grouped = group_batches_by_dryer(batches)
        wash_mins = int(self.wash_minutes.get())
        day_minutes = minutes_between(self.window_start_dt, self.window_end_dt)  # 1440

        labels = [f"D{d}" for d in self.active_dryers]
        scheduled_list: List[float] = []
        wash_list: List[float] = []
        idle_list: List[float] = []

        for dryer in self.active_dryers:
            b_list = sorted(grouped.get(dryer, []), key=lambda x: x.start_dt)

            scheduled = 0.0
            for b in b_list:
                seg_start = max(b.start_dt, self.window_start_dt)
                seg_end = min(b.end_dt, self.window_end_dt)
                if seg_end > seg_start:
                    scheduled += minutes_between(seg_start, seg_end)

            wash_segs = compute_wash_segments(b_list, wash_mins)
            wash = 0.0
            for (ws, we) in wash_segs:
                seg_start = max(ws, self.window_start_dt)
                seg_end = min(we, self.window_end_dt)
                if seg_end > seg_start:
                    wash += minutes_between(seg_start, seg_end)

            used = scheduled + wash
            idle = max(0.0, day_minutes - used)

            scheduled_list.append(scheduled)
            wash_list.append(wash)
            idle_list.append(idle)

        self.ax_totals.bar(labels, scheduled_list, label="Scheduled", color="tab:blue")
        self.ax_totals.bar(labels, wash_list, bottom=scheduled_list, label="Wash", color="tab:orange")

        bottoms = [scheduled_list[i] + wash_list[i] for i in range(len(self.active_dryers))]
        self.ax_totals.bar(labels, idle_list, bottom=bottoms, label="Idle", color="tab:gray", alpha=0.5)

        self.ax_totals.set_title(f"Totals vs Available (per dryer, {day_minutes:.0f} min)")
        self.ax_totals.set_xlabel("Dryer")
        self.ax_totals.set_ylabel("Minutes")
        self.ax_totals.set_ylim(0, day_minutes)

        self.fig_totals.subplots_adjust(bottom=0.18)
        self.ax_totals.legend(
            loc="upper right",
            bbox_to_anchor=(1.0, -0.12),
            frameon=True,
            ncol=3
        )

        self.fig_totals.tight_layout()
        self.canvas_totals.draw()

    def on_timeline_click(self, event):
        if event.inaxes != self.ax_timeline or event.ydata is None:
            return

        y_clicked = int(round(event.ydata))
        dryer = self.y_to_dryer.get(y_clicked)
        if not dryer:
            return

        self.open_dryer_details(dryer)

    def open_dryer_details(self, dryer: int):
        batches = self.current_batches()
        grouped = group_batches_by_dryer(batches)

        b_list = sorted(grouped.get(dryer, []), key=lambda x: x.start_dt)
        wash_mins = int(self.wash_minutes.get())
        day_minutes = minutes_between(self.window_start_dt, self.window_end_dt)

        scheduled = sum(
            minutes_between(max(b.start_dt, self.window_start_dt), min(b.end_dt, self.window_end_dt))
            for b in b_list
            if min(b.end_dt, self.window_end_dt) > max(b.start_dt, self.window_start_dt)
        )

        wash_segs = compute_wash_segments(b_list, wash_mins)
        wash = sum(
            minutes_between(max(ws, self.window_start_dt), min(we, self.window_end_dt))
            for (ws, we) in wash_segs
            if min(we, self.window_end_dt) > max(ws, self.window_start_dt)
        )

        used = scheduled + wash
        idle = max(0.0, day_minutes - used)
        fill = (used / day_minutes) if day_minutes > 0 else 0.0

        gaps = compute_gaps_for_dryer(b_list, self.window_start_dt, self.window_end_dt)
        gaps_sorted = sorted(gaps, key=lambda ab: minutes_between(ab[0], ab[1]), reverse=True)

        win = tk.Toplevel(self)
        win.title(f"Dryer {dryer} Details")
        win.geometry("760x580")

        try:
            if os.path.exists(ICON_PATH):
                win.iconbitmap(ICON_PATH)
        except Exception:
            pass

        header = tk.Label(
            win,
            text=f"Dryer {dryer}   Scheduled: {scheduled:.0f} min   Wash: {wash:.0f} min   Idle: {idle:.0f} min   Fill: {fill * 100:.1f}%",
            font=("Segoe UI", 12, "bold"),
            wraplength=720
        )
        header.pack(pady=10)

        gap_frame = tk.Frame(win)
        gap_frame.pack(fill="x", padx=12, pady=8)

        tk.Label(gap_frame, text="Largest gaps:", font=("Segoe UI", 11)).pack(anchor="w")

        gap_box = tk.Listbox(gap_frame, height=9)
        gap_box.pack(fill="x", pady=6)

        if gaps_sorted:
            for (a, b) in gaps_sorted[:12]:
                mins = minutes_between(a, b)
                a_str = a.strftime("%I:%M %p").lstrip("0")
                b_str = b.strftime("%I:%M %p").lstrip("0")
                gap_box.insert(tk.END, f"{a_str} to {b_str}  =  {mins:.0f} min")
        else:
            gap_box.insert(tk.END, "No gaps found in the window.")

        fig = Figure(figsize=(6.8, 3.2), dpi=100)
        ax = fig.add_subplot(111)
        ax.bar(["Scheduled", "Wash", "Idle"], [scheduled, wash, idle])
        ax.set_title("Day Breakdown (minutes)")
        ax.set_ylabel("Minutes")
        fig.tight_layout()

        canvas = FigureCanvasTkAgg(fig, master=win)
        canvas.get_tk_widget().pack(fill="both", expand=True, padx=12, pady=10)
        canvas.draw()


# ======================================================================
# MAIN
# ======================================================================
if __name__ == "__main__":
    app = App()
    app.mainloop()
