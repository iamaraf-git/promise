import threading
import queue
import sys
import os
import requests
import datetime
import calendar
import time
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
try:
    from PIL import Image, ImageTk
    HAS_PIL = True
except ImportError:
    HAS_PIL = False

# ----------------------------
# CORE AUTOMATION LINKAGES
# ----------------------------
try:
    from main import is_cdp_running, launch_edge_with_cdp, ensure_promise_page
except ImportError:
    def is_cdp_running(): return False
    def launch_edge_with_cdp(): return False
    def ensure_promise_page(): pass

try:
    from main import run_automation
except Exception:
    run_automation = None

# ============================================================
# DESIGN COLORS & TYPOGRAPHY
# ============================================================
COLOR_BG = "#FAFAFA"          # Off-white app background
COLOR_SURFACE = "#FFFFFF"     # White widget cards
COLOR_BORDER = "#E5E7EB"      # Clean light-gray borders
COLOR_BORDER_HOVER = "#D1D5DB"
COLOR_TEXT = "#111827"        # Dark slate text
COLOR_TEXT_BODY = "#374151"
COLOR_TEXT_MUTED = "#6B7280"
COLOR_PRIMARY = "#4F46E5"     # Indigo accent
COLOR_HINT_BG = "#EEF2FF"     # Soft light-indigo hint card background
COLOR_HINT_TEXT = "#3730A3"
COLOR_LOG_BG = "#F9FAFB"      # Light-gray logging box
COLOR_LOG_INFO = "#2563EB"    # Blue info
COLOR_LOG_OK = "#16A34A"      # Green success
COLOR_LOG_ERR = "#DC2626"     # Red errors
COLOR_PROGRESS_TRACK = "#E5E7EB"
COLOR_PROGRESS_FILL = "#4F46E5"

FONT_TITLE = ("Segoe UI", 16, "bold")
FONT_LABEL = ("Segoe UI", 10)
FONT_HINT = ("Segoe UI", 9)
FONT_BUTTON = ("Segoe UI", 10, "bold")
FONT_INPUT = ("Segoe UI", 10)
FONT_LOG = ("Consolas", 9)
FONT_PROGRESS_HEADER = ("Segoe UI", 10, "bold")
FONT_PROGRESS_PERCENT = ("Segoe UI", 11, "bold")
FONT_PROGRESS_STATUS = ("Segoe UI", 9)

# ============================================================
# GLOBAL STATE & THREADING FLAGGING
# ============================================================
cdp_connected = False
stop_requested = False
log_queue = queue.Queue()
worker_thread = None
connect_thread = None

# ============================================================
# FILE & DIRECTORY PICKERS
# ============================================================
def load_file_statistics(file_path):
    """Calculates statistics for processable rows in the chosen sheet and logs them."""
    try:
        df = pd.read_excel(file_path, dtype=str, keep_default_na=False).fillna("")
        total_rows = len(df)
        
        # Verify required columns are present
        required = ["Contract Name", "Medicaid Number", "Date of Birth", "Last Name", "First Name"]
        missing = [col for col in required if col not in df.columns]
        
        log_to_gui("📊 SPREADSHEET STATISTICS:", "info")
        log_to_gui(f"  📄 Total Rows: {total_rows}", "info")
        if missing:
            log_to_gui(f"  ❌ Missing required columns: {', '.join(missing)}", "error")
        else:
            log_to_gui("  ✅ All required columns are present. Ready to process.", "success")
            
    except Exception as e:
        log_to_gui(f"❌ Failed loading stats: {e}", "error")

def browse_excel():
    """Opens file dialogue for selecting spreadsheet input."""
    filename = filedialog.askopenfilename(
        title="Select Promise Excel File",
        filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
    )
    if filename:
        excel_path.set(filename)
        log_box.config(state="normal")
        log_box.delete("1.0", tk.END)
        log_box.config(state="disabled")
        log_to_gui(f"📂 Selected input file: {filename}\n", "info")
        
        p_check = os.path.join(os.path.dirname(filename), f"{os.path.splitext(os.path.basename(filename))[0]}_progress.csv")
        if os.path.exists(p_check):
            log_to_gui("ℹ Found an existing session progress log file. System will automatically resume from the last saved milestone.\n", "info")
            
        load_file_statistics(filename)

def browse_folder():
    """Opens directory dialogue for target output save folder."""
    folder = filedialog.askdirectory(title="Select Output Folder")
    if folder:
        output_folder.set(folder)
        log_to_gui(f"📁 Selected output folder: {folder}", "info")

# ============================================================
# SAFE LOGGING & PROGRESS UPDATE HELPERS
# ============================================================
def log(message):
    """Callback logger logging onto the queue with timestamp prefix."""
    timestamp = datetime.datetime.now().strftime("%H:%M:%S")
    formatted = f"[{timestamp}] {message}"
    log_queue.put(formatted)

def log_to_gui(msg, tag="info"):
    """Thread-safe logging onto the Consolas styled viewport."""
    def _insert():
        log_box.config(state="normal")
        formatted = msg if msg.endswith("\n") else msg + "\n"
        log_box.insert(tk.END, formatted, tag)
        log_box.see(tk.END)
        log_box.config(state="disabled")
    app.after(0, _insert)

def update_progress(current, total, member_id="-"):
    """Updates progress horizontal bar and labels thread-safely."""
    def _apply():
        total_safe = max(int(total), 1)
        current_safe = max(0, min(int(current), total_safe))
        pct = int(round(100 * current_safe / total_safe))
        progress_bar.configure(maximum=total_safe, value=current_safe)
        progress_percent_var.set(f"{pct}%")
        progress_status_var.set(f"{current_safe} of {total_safe} rows processed (Member ID: {member_id})")
    app.after(0, _apply)

def process_log_queue():
    """Consumes queue messages thread-safely and color-codes tags."""
    while not log_queue.empty():
        message = log_queue.get()
        tag = "info"
        if any(err in message.lower() for err in ["❌", "error", "warning", "⚠️"]):
            tag = "error"
        elif any(ok in message.lower() for ok in ["✅", "🟢", "success", "perfect match"]):
            tag = "success"
        log_to_gui(message, tag)
    app.after(100, process_log_queue)

# ============================================================
# BUTTON COLORS & STATES ORCHESTRATION
# ============================================================
BUTTON_STYLES = {
    "connect": {
        "active":   {"bg": "#4F46E5", "fg": "white",   "border": "#4F46E5"},
        "disabled": {"bg": "#E0E7FF", "fg": "#818CF8", "border": "#E0E7FF"},
    },
    "start": {
        "active":   {"bg": "#16A34A", "fg": "white",   "border": "#16A34A"},
        "disabled": {"bg": "#DCFCE7", "fg": "#16A34A", "border": "#86EFAC"},
    },
    "stop": {
        "active":   {"bg": "#DC2626", "fg": "white",   "border": "#DC2626"},
        "disabled": {"bg": "#FEE2E2", "fg": "#DC2626", "border": "#FCA5A5"},
    },
}

def _apply_button_style(btn, variant, enabled):
    """Controls color settings dynamically, guaranteeing white button text."""
    if btn is None:
        return
    palette = BUTTON_STYLES[variant]["active" if enabled else "disabled"]
    btn.config(
        state="normal" if enabled else "disabled",
        bg=palette["bg"],
        activebackground=palette["bg"],
        fg=palette["fg"],
        activeforeground=palette["fg"],
        disabledforeground=palette["fg"],
        highlightbackground=palette["border"],
        highlightcolor=palette["border"],
    )

def _set_ui_state(ui_state):
    """Switches operational states and button styling configs on main thread."""
    def _apply():
        if ui_state == "idle":
            _apply_button_style(connect_btn, "connect", True)
            _apply_button_style(start_btn, "start", False)
            _apply_button_style(stop_btn, "stop", False)
            excel_browse_btn.config(state="normal")
            output_browse_btn.config(state="normal")
        elif ui_state == "connecting":
            _apply_button_style(connect_btn, "connect", False)
            _apply_button_style(start_btn, "start", False)
            _apply_button_style(stop_btn, "stop", True)
            excel_browse_btn.config(state="disabled")
            output_browse_btn.config(state="disabled")
        elif ui_state == "ready":
            _apply_button_style(connect_btn, "connect", True)
            _apply_button_style(start_btn, "start", True)
            _apply_button_style(stop_btn, "stop", False)
            excel_browse_btn.config(state="normal")
            output_browse_btn.config(state="normal")
        elif ui_state == "running":
            _apply_button_style(connect_btn, "connect", False)
            _apply_button_style(start_btn, "start", False)
            _apply_button_style(stop_btn, "stop", True)
            excel_browse_btn.config(state="disabled")
            output_browse_btn.config(state="disabled")
    app.after(0, _apply)

# ============================================================
# CONNECT BROWSER LOGIC (BACKGROUND THREADED)
# ============================================================
def connect_browser_thread():
    """Validates remote debugger port 9222 and connects Playwright browser."""
    global cdp_connected
    log("🔍 Auditing browser debugger port 9222 connection status...")
    cdp_connected = is_cdp_running()
    
    if cdp_connected:
        log("✅ Existing Edge browser debugging session detected.")
        try:
            ensure_promise_page()
            log("✅ Promise portal connection verified and page initialized.")
            log("🔑 Please log in to the Promise portal if needed, then click 'Start'.")
            _set_ui_state("ready")
        except Exception as e:
            log(f"❌ Failed connecting to active Promise portal page: {e}")
            _set_ui_state("idle")
        return
        
    log("🚀 Browser debug port not active. Launching isolated Edge instance...")
    if launch_edge_with_cdp():
        cdp_connected = True
        log("✅ Debugging browser initialized successfully.")
        try:
            ensure_promise_page()
            log("✅ Promise portal connection verified and page initialized.")
            log("🔑 Please log in to the Promise portal if needed, then click 'Start'.")
            _set_ui_state("ready")
        except Exception as e:
            log(f"❌ Failed establishing portal target: {e}")
            _set_ui_state("idle")
    else:
        cdp_connected = False
        log("❌ Edge remote debug launcher failed. Please close all other Edge instances and retry.")
        _set_ui_state("idle")

def start_connect_browser():
    """Trigger background check/connection logic thread."""
    global connect_thread
    log_box.config(state="normal")
    log_box.delete("1.0", tk.END)
    log_box.config(state="disabled")
    _set_ui_state("connecting")
    
    connect_thread = threading.Thread(target=connect_browser_thread, daemon=True)
    connect_thread.start()

# ============================================================
# AUTOMATION RUN WORKER WRAPPERS
# ============================================================
def run_wrapper(excel_file, output_folder_path):
    """Invokes core scraping sequences on background thread."""
    try:
        if run_automation:
            run_automation(
                excel_path=excel_file,
                output_base_folder=output_folder_path,
                log_callback=log,
                progress_callback=update_progress,
                stop_check=lambda: stop_requested,
            )
        else:
            log("❌ Error: run_automation function is missing from main.py")
    except Exception as e:
        log(f"❌ Automation Error: {e}")
    finally:
        _set_ui_state("ready")

def start_automation():
    """Validates parameters, resets progress, and spawns crawler thread."""
    global worker_thread
    excel_file = excel_path.get().strip()
    output_f = output_folder.get().strip()

    if not excel_file:
        log_to_gui("❌ Please select a source Excel file first", "error")
        return
    if not output_f:
        log_to_gui("❌ Please select an output folder first", "error")
        return
    if not cdp_connected:
        log_to_gui("❌ Browser connection is not established. Connect first.", "error")
        return

    global stop_requested
    stop_requested = False
    
    log_box.config(state="normal")
    log_box.delete("1.0", tk.END)
    log_box.config(state="disabled")
    
    progress_bar.configure(maximum=100, value=0)
    progress_percent_var.set("0%")
    progress_status_var.set("Connecting and verifying...")
    _set_ui_state("running")

    worker_thread = threading.Thread(
        target=run_wrapper,
        args=(excel_file, output_f),
        daemon=True
    )
    worker_thread.start()

def stop_automation():
    """Flags worker run loop to abort immediately."""
    global stop_requested
    stop_requested = True
    stop_btn.config(state="disabled")
    log("🛑 Stop request sent. Finalizing spreadsheet outputs and closing page context...")

def on_closing():
    """Intercepts window close events to prevent data corruption during active runs."""
    global worker_thread
    if worker_thread and worker_thread.is_alive():
        if messagebox.askokcancel(
            "Exit Warning",
            "An automation is currently running.\n\nWould you like to stop it and exit safely?"
        ):
            stop_automation()
            log("⏳ Waiting briefly for automation to safely write spreadsheet progress...")
            app.after(1500, app.destroy)
    else:
        app.destroy()

# ============================================================
# TKINTER WINDOW BUILDING & GRIDING
# ============================================================
def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

if getattr(sys, "frozen", False):
    base_path = sys._MEIPASS
else:
    base_path = os.path.dirname(__file__)

app = tk.Tk()
app.title("AZ Billing Automation - Promise Eligibility Checker")
app.geometry("850x780")
app.minsize(800, 650)
app.configure(bg=COLOR_BG)

# Load title bar icon
try:
    icon_path = os.path.join(base_path, "azbilling-logo.ico")
    if os.path.exists(icon_path):
        app.iconbitmap(icon_path)
except Exception:
    pass

# Helper GUI Builders
def _label(parent, text, font=FONT_LABEL, fg=COLOR_TEXT_BODY):
    return tk.Label(parent, text=text, font=font, bg=COLOR_BG, fg=fg, anchor="w")

def _entry(parent, var, width=46):
    return tk.Entry(
        parent, textvariable=var, width=width,
        font=FONT_INPUT, bg=COLOR_SURFACE, fg=COLOR_TEXT,
        relief="flat", bd=0,
        highlightthickness=1,
        highlightbackground=COLOR_BORDER,
        highlightcolor=COLOR_PRIMARY,
        insertbackground=COLOR_PRIMARY,
    )

def _browse(parent, command):
    return tk.Button(
        parent, text="Browse", command=command,
        font=("Segoe UI", 9), bg=COLOR_SURFACE, fg=COLOR_TEXT_BODY,
        activebackground=COLOR_LOG_BG, activeforeground=COLOR_TEXT,
        relief="flat", bd=0,
        highlightthickness=1,
        highlightbackground=COLOR_BORDER,
        highlightcolor=COLOR_BORDER_HOVER,
        padx=16, pady=5, cursor="hand2",
    )

app.grid_columnconfigure(0, weight=1)
app.grid_rowconfigure(0, weight=1)

# Core layout wrapper frame
main_frame = tk.Frame(app, padx=28, pady=20, bg=COLOR_BG)
main_frame.pack(fill=tk.BOTH, expand=True)

# Header Row containing Logo and Title
header_frame = tk.Frame(main_frame, bg=COLOR_BG)
header_frame.grid(row=0, column=0, columnspan=3, sticky="ew", pady=(0, 10))

logo_png_path = os.path.join(base_path, "azbilling-new-logo.png")
if HAS_PIL and os.path.exists(logo_png_path):
    try:
        pil_img = Image.open(logo_png_path)
        pil_img = pil_img.resize((70, 70), Image.Resampling.LANCZOS)
        logo_photo = ImageTk.PhotoImage(pil_img)
        logo_label = tk.Label(header_frame, image=logo_photo, bg=COLOR_BG)
        logo_label.image = logo_photo  # Keep reference
        logo_label.pack(side=tk.LEFT, padx=(0, 15))
    except Exception:
        pass

title_label = tk.Label(
    header_frame, text="Promise Eligibility Checker", 
    font=FONT_TITLE, fg=COLOR_TEXT, bg=COLOR_BG, anchor="w"
)
title_label.pack(side=tk.LEFT, fill=tk.Y)

# Divider Line
tk.Frame(main_frame, height=1, bg=COLOR_BORDER).grid(
    row=1, column=0, columnspan=3, sticky="ew", pady=(0, 18),
)

# Text Variables Setup
excel_path = tk.StringVar()
output_folder = tk.StringVar()

# Inputs Sections
_label(main_frame, "Excel File").grid(row=2, column=0, sticky="w", pady=8, padx=(0, 16))
_entry(main_frame, excel_path).grid(row=2, column=1, sticky="ew", pady=8, padx=(0, 10), ipady=6)
excel_browse_btn = _browse(main_frame, browse_excel)
excel_browse_btn.grid(row=2, column=2, sticky="w", pady=8)

_label(main_frame, "Output Folder").grid(row=3, column=0, sticky="w", pady=8, padx=(0, 16))
_entry(main_frame, output_folder).grid(row=3, column=1, sticky="ew", pady=8, padx=(0, 10), ipady=6)
output_browse_btn = _browse(main_frame, browse_folder)
output_browse_btn.grid(row=3, column=2, sticky="w", pady=8)

# Elegant Instructions Panel
hint_frame = tk.Frame(main_frame, bg=COLOR_HINT_BG)
hint_frame.grid(row=4, column=0, columnspan=3, sticky="ew", pady=(16, 0))

tk.Frame(hint_frame, width=3, bg=COLOR_PRIMARY).pack(side=tk.LEFT, fill=tk.Y)

hint_label = tk.Label(
    hint_frame,
    text=(
        "Step 1: Select your source Excel spreadsheet and an output folder.\n"
        "Step 2: Click 'Connect Browser' to launch Edge and load the Promise login page. Log in if needed.\n"
        "Step 3: Click 'Start' to begin automated eligibility lookups. Already processed rows will be skipped.\n"
        "Step 4: Click 'Stop' to halt processing, save your current spreadsheet progress, and reset."
    ),
    font=FONT_HINT, bg=COLOR_HINT_BG, fg=COLOR_HINT_TEXT,
    justify=tk.LEFT, anchor="w", wraplength=730,
)
hint_label.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=14, pady=10)

hint_frame.bind(
    "<Configure>",
    lambda e: hint_label.config(wraplength=max(200, e.width - 40)),
)

# Button Navigation Console
btn_frame = tk.Frame(main_frame, bg=COLOR_BG)
btn_frame.grid(row=5, column=0, columnspan=3, pady=20)

button_kwargs = dict(
    font=FONT_BUTTON, relief="flat", bd=0,
    padx=22, pady=9, cursor="hand2",
    highlightthickness=2,
)

connect_btn = tk.Button(
    btn_frame, text="Connect Browser",
    command=start_connect_browser, **button_kwargs
)
connect_btn.pack(side=tk.LEFT, padx=8)

start_btn = tk.Button(
    btn_frame, text="Start",
    command=start_automation, **button_kwargs
)
start_btn.pack(side=tk.LEFT, padx=8)

stop_btn = tk.Button(
    btn_frame, text="Stop",
    command=stop_automation, **button_kwargs
)
stop_btn.pack(side=tk.LEFT, padx=8)

# Initialize button colors
_apply_button_style(connect_btn, "connect", True)
_apply_button_style(start_btn, "start", False)
_apply_button_style(stop_btn, "stop", False)

# Automation Progress Card
progress_frame = tk.Frame(
    main_frame, bg=COLOR_SURFACE, highlightthickness=1, highlightbackground=COLOR_BORDER
)
progress_frame.grid(row=6, column=0, columnspan=3, sticky="ew", pady=(0, 16))

progress_inner = tk.Frame(progress_frame, bg=COLOR_SURFACE, padx=16, pady=14)
progress_inner.pack(fill=tk.X, expand=True)

progress_header = tk.Frame(progress_inner, bg=COLOR_SURFACE)
progress_header.pack(fill=tk.X)

_label(progress_header, "Automation Progress", font=FONT_PROGRESS_HEADER, fg=COLOR_TEXT).pack(side=tk.LEFT)

progress_percent_var = tk.StringVar(value="0%")
progress_status_var = tk.StringVar(value="Waiting to start…")

tk.Label(
    progress_header, textvariable=progress_percent_var,
    font=FONT_PROGRESS_PERCENT, bg=COLOR_SURFACE, fg=COLOR_PRIMARY,
).pack(side=tk.RIGHT)

# ttk Style Configuration for Progressbar
style = ttk.Style()
try:
    style.theme_use("clam")
except tk.TclError:
    pass

style.configure(
    "Automation.Horizontal.TProgressbar",
    troughcolor=COLOR_PROGRESS_TRACK,
    background=COLOR_PROGRESS_FILL,
    bordercolor=COLOR_BORDER,
    lightcolor=COLOR_PROGRESS_FILL,
    darkcolor=COLOR_PROGRESS_FILL,
    thickness=14,
)

progress_bar = ttk.Progressbar(
    progress_inner,
    style="Automation.Horizontal.TProgressbar",
    orient="horizontal",
    mode="determinate",
    maximum=100,
    value=0,
)
progress_bar.pack(fill=tk.X, pady=(10, 8))

tk.Label(
    progress_inner, textvariable=progress_status_var,
    font=FONT_PROGRESS_STATUS, bg=COLOR_SURFACE, fg=COLOR_TEXT_MUTED, anchor="w",
).pack(fill=tk.X)

# System Log Header
_label(main_frame, "Activity Log", font=("Segoe UI", 10, "bold"), fg=COLOR_TEXT).grid(
    row=7, column=0, sticky="w", pady=(6, 8),
)

# Scrollable Console Viewport
log_wrap = tk.Frame(main_frame, bg=COLOR_BORDER)
log_wrap.grid(row=8, column=0, columnspan=3, sticky="nsew")

log_inner = tk.Frame(log_wrap, bg=COLOR_LOG_BG)
log_inner.pack(fill=tk.BOTH, expand=True, padx=1, pady=1)

scrollbar = tk.Scrollbar(log_inner, bd=0, highlightthickness=0)
scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

log_box = tk.Text(
    log_inner, height=12, width=90,
    yscrollcommand=scrollbar.set, state="disabled", wrap=tk.WORD,
    bg=COLOR_LOG_BG, fg=COLOR_TEXT_BODY,
    font=FONT_LOG, relief="flat", bd=0,
    padx=14, pady=12,
    insertbackground=COLOR_PRIMARY,
    selectbackground="#DBEAFE", selectforeground=COLOR_TEXT,
)
log_box.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
scrollbar.config(command=log_box.yview)

# Activity Tag Configurations
log_box.tag_config("info", foreground=COLOR_LOG_INFO)
log_box.tag_config("success", foreground=COLOR_LOG_OK)
log_box.tag_config("error", foreground=COLOR_LOG_ERR)

# Layout resizing constraints
main_frame.grid_rowconfigure(8, weight=1)
main_frame.grid_columnconfigure(1, weight=1)

# Register window closing safety handler
app.protocol("WM_DELETE_WINDOW", on_closing)

# Start log queue listener and main window event loop
if __name__ == "__main__":
    process_log_queue()
    app.mainloop()
