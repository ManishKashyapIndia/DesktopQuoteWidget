import colorsys
import ctypes
import json
import os
import random
import shutil
import subprocess
import sys
import time
import traceback
import webbrowser
import winreg

# Hide console window immediately if launched via python.exe
_console = ctypes.windll.kernel32.GetConsoleWindow()
if _console:
    ctypes.windll.user32.ShowWindow(_console, 0)  # 0 = SW_HIDE


# Windows native MessageBox helper (works even if Tkinter or libraries fail to load)
def show_windows_popup(title, message, is_error=True):
    MB_OK = 0x0
    MB_ICONERROR = 0x10
    MB_ICONINFO = 0x40
    icon = MB_ICONERROR if is_error else MB_ICONINFO
    try:
        ctypes.windll.user32.MessageBoxW(0, str(message), str(title), MB_OK | icon)
    except Exception:
        print(f"{title}: {message}", file=sys.stderr)


# Single-instance lock using a Windows Named Mutex
MUTEX_NAME = "Global\\DesktopQuoteWidget_SingleInstance_Mutex_v1"
ERROR_ALREADY_EXISTS = 183


def acquire_single_instance_lock():
    kernel32 = ctypes.windll.kernel32
    mutex = kernel32.CreateMutexW(None, False, MUTEX_NAME)
    if not mutex:
        return None, False
    if kernel32.GetLastError() == ERROR_ALREADY_EXISTS:
        kernel32.CloseHandle(mutex)
        return None, True
    return mutex, False


def release_single_instance_lock(mutex):
    if mutex:
        try:
            ctypes.windll.kernel32.ReleaseMutex(mutex)
        except Exception:
            pass
        try:
            ctypes.windll.kernel32.CloseHandle(mutex)
        except Exception:
            pass


# Check required libraries early and show a clear error if anything is missing
try:
    import tkinter as tk
    from tkinter import colorchooser, filedialog, messagebox, ttk
    import openpyxl
    from openpyxl.styles import Alignment, Font, PatternFill
    from PIL import Image, ImageChops, ImageDraw, ImageFilter, ImageFont, ImageTk
except ImportError as exc:
    show_windows_popup(
        "Quote Widget - Failed to Start",
        f"The program failed to start because a required library is missing:\n\n"
        f"{exc}\n\n"
        f"Fix: Open Command Prompt and run:\n"
        f"pip install openpyxl pillow",
        is_error=True,
    )
    sys.exit(1)
except Exception as exc:
    show_windows_popup(
        "Quote Widget - Failed to Start",
        f"The program failed to start due to an unexpected error:\n\n{exc}",
        is_error=True,
    )
    sys.exit(1)


# Portable paths: works both as a .pyw script and as a PyInstaller frozen .exe
if getattr(sys, "frozen", False):
    BASE_DIR = os.path.dirname(os.path.abspath(sys.executable))
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

APP_DISPLAY_NAME = "Windows Quote Widget"
DEFAULT_INSTALL_DIR = r"C:\Windows Quote Widget"
EXCEL_FILENAME = "quotes.xlsx"
CONFIG_FILENAME = "widget_config.json"
INSTALL_MARKER_FILENAME = ".installed"
ICON_FILENAME = "app_icon.ico"

# Central pointer file in %LOCALAPPDATA% so the app always remembers its install folder
LOCAL_STATE_DIR = os.path.join(
    os.environ.get("LOCALAPPDATA", os.path.expanduser("~")), "WindowsQuoteWidget"
)
INSTALL_STATE_FILE = os.path.join(LOCAL_STATE_DIR, "install_state.json")

DEFAULT_EXCEL_FILE = os.path.join(BASE_DIR, EXCEL_FILENAME)
CONFIG_FILE = os.path.join(BASE_DIR, CONFIG_FILENAME)
GITHUB_URL = "https://github.com/ManishKashyapIndia"
CHROMA_KEY = "#010203"
CHROMA_RGB = (1, 2, 3)

# Windows Startup & Theme Registry constants
APP_REG_NAME = "DesktopQuoteWidget"
RUN_REG_PATH = r"Software\Microsoft\Windows\CurrentVersion\Run"
THEME_REG_PATH = r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize"

# Win32 window constants (desktop pinning & taskbar hiding)
HWND_BOTTOM = 1
SWP_NOSIZE = 0x0001
SWP_NOMOVE = 0x0002
SWP_NOZORDER = 0x0004
SWP_NOACTIVATE = 0x0010
SWP_FRAMECHANGED = 0x0020
GWL_EXSTYLE = -20
WS_EX_APPWINDOW = 0x00040000
WS_EX_TOOLWINDOW = 0x00000080

# History limits
MAX_PREV_HISTORY = 2          # Remember at most 2 previous quotes
DEFAULT_UNDO_LIMIT = 5        # Default undo history limit (user-configurable 1..20)
MIN_UNDO_LIMIT = 1
MAX_UNDO_LIMIT = 20

AVAILABLE_SHAPES = [
    "Rounded",
    "Pill",
    "Leaf",
    "Ticket",
    "Scalloped",
    "Cloud",
    "Banner",
    "Ribbon",
    "Chamfered",
    "Speech",
    "Rectangle",
]

AVAILABLE_BORDER_PATTERNS = [
    "Solid",
    "Thick",
    "Dashed",
    "Dotted",
    "Double",
    "Triple",
    "Glow",
    "Neon",
    "Gradient",
    "Rainbow",
    "Glitter",
]

ANIMATION_TYPES = [
    "Color Flash",
    "Border Pulse",
    "Jitter",
    "Fade In",
    "Pop Scale",
]

# Granularity units and their maximum allowed values (capped at 1 year)
INTERVAL_UNITS = {
    "Minutes": (60, 525600),
    "Hours": (3600, 8760),
    "Days": (86400, 365),
    "Weeks": (604800, 52),
    "Months": (2592000, 12),
    "Years": (31536000, 1),
}

# 100 Original, Copyright-Free Quotes (20 per category) for auto-generating quotes.xlsx
DEFAULT_QUOTES = [
    # Atomic Habits (20)
    ("Small daily actions quietly build the architecture of your future.", "Atomic Habits"),
    ("Focus on building reliable systems rather than chasing distant goals.", "Atomic Habits"),
    ("Every repeated choice casts a vote for the person you are becoming.", "Atomic Habits"),
    ("Make good habits obvious in your environment and bad ones harder to reach.", "Atomic Habits"),
    ("A tiny improvement repeated every day compounds into remarkable change.", "Atomic Habits"),
    ("Consistency matters far more than occasional bursts of intensity.", "Atomic Habits"),
    ("If you slip up once, return to your routine the very next day.", "Atomic Habits"),
    ("Design your surroundings so the right action is the easiest one to take.", "Atomic Habits"),
    ("Start a new routine with a version so small you cannot say no to it.", "Atomic Habits"),
    ("Before optimizing a habit, simply show up and establish the rhythm.", "Atomic Habits"),
    ("Pay more attention to your daily direction than to your current position.", "Atomic Habits"),
    ("Planning feels productive, but taking action creates real results.", "Atomic Habits"),
    ("True mastery requires pushing through the boredom of daily repetition.", "Atomic Habits"),
    ("Helpful routines turn time into an ally; harmful ones turn it into an obstacle.", "Atomic Habits"),
    ("Shape your environment deliberately instead of passively reacting to it.", "Atomic Habits"),
    ("Actions that feel rewarding right away are the ones you will repeat.", "Atomic Habits"),
    ("Lasting change begins when a habit becomes part of your identity.", "Atomic Habits"),
    ("Pair a habit you need to do with an activity you genuinely enjoy.", "Atomic Habits"),
    ("Track your streak visually—seeing progress fuels momentum.", "Atomic Habits"),
    ("Reduce friction for good behaviors and add friction to distractions.", "Atomic Habits"),

    # Health Motivation (20)
    ("Your body is your lifelong home; maintain it with daily care.", "Health Motivation"),
    ("Outer vitality is a direct reflection of inner nourishment and rest.", "Health Motivation"),
    ("Daily movement refreshes both the body and the mind.", "Health Motivation"),
    ("Make time for wellness today so you do not lose time to illness tomorrow.", "Health Motivation"),
    ("A steady daily walk outperforms an occasional exhausting workout.", "Health Motivation"),
    ("Stay active, breathe deeply, and let movement clear mental clutter.", "Health Motivation"),
    ("Good health is the most valuable asset you can invest in every morning.", "Health Motivation"),
    ("Simple whole foods, clean water, and regular movement form the core of health.", "Health Motivation"),
    ("Walking briskly each day is one of the simplest medicines available.", "Health Motivation"),
    ("Fuel your body well so your energy matches your ambitions.", "Health Motivation"),
    ("Quality sleep is the quiet foundation of physical and mental recovery.", "Health Motivation"),
    ("Measure your fitness against who you were last month, not against others.", "Health Motivation"),
    ("Strong health is the platform upon which all other joys are built.", "Health Motivation"),
    ("Choose long-term vitality over momentary comfort.", "Health Motivation"),
    ("Every step you take moves you closer to a stronger version of yourself.", "Health Motivation"),
    ("Morning sunlight and fresh air set a positive tone for the entire day.", "Health Motivation"),
    ("Peace of mind and a healthy body are the truest forms of wealth.", "Health Motivation"),
    ("Modest, consistent exercise habits yield lifelong resilience.", "Health Motivation"),
    ("You will never regret completing a workout, even on a tired day.", "Health Motivation"),
    ("Physical strength grows each time you persist through mild discomfort.", "Health Motivation"),

    # Study Motivation (20)
    ("Dedicated learning today opens doors that luck alone never will.", "Study Motivation"),
    ("Every skilled scholar started by patiently learning the basics.", "Study Motivation"),
    ("Keep moving forward steadily, one page and one problem at a time.", "Study Motivation"),
    ("Focused preparation in quiet hours builds confidence for big tests.", "Study Motivation"),
    ("Take ownership of your learning; no one else can build your mind for you.", "Study Motivation"),
    ("Genuine understanding has no shortcuts—embrace the process.", "Study Motivation"),
    ("Academic excellence is built through small, daily study sessions.", "Study Motivation"),
    ("Aim to grasp the underlying principle rather than memorizing surface facts.", "Study Motivation"),
    ("The concentration you practice today becomes your capability tomorrow.", "Study Motivation"),
    ("Delaying a task only makes it feel heavier; start with five minutes now.", "Study Motivation"),
    ("Small daily gains in knowledge accumulate into deep expertise.", "Study Motivation"),
    ("Your future opportunities are shaped by what you choose to learn today.", "Study Motivation"),
    ("Stay curious—asking thoughtful questions is the engine of learning.", "Study Motivation"),
    ("When your fundamentals are solid, advanced topics become approachable.", "Study Motivation"),
    ("Self-discipline turns academic goals into concrete achievements.", "Study Motivation"),
    ("Errors are not setbacks; they are signposts showing what to refine next.", "Study Motivation"),
    ("Read actively—question, connect, and apply what you discover.", "Study Motivation"),
    ("Investing time in knowledge delivers dividends for an entire lifetime.", "Study Motivation"),
    ("Stay humble in victory, persistent in difficulty, and steady in effort.", "Study Motivation"),
    ("Each completed chapter brings you closer to true command of your subject.", "Study Motivation"),

    # Finance (20)
    ("Set aside your savings first, and plan your spending around what remains.", "Finance"),
    ("Patient compounding rewards those who start early and stay consistent.", "Finance"),
    ("Financial risk shrinks when you take time to understand what you own.", "Finance"),
    ("True prosperity is needing less, not constantly acquiring more.", "Finance"),
    ("Long-term markets reward patience far more than impulsive reactions.", "Finance"),
    ("Look beyond the price tag and focus on lasting utility and value.", "Finance"),
    ("Watch small recurring expenses; over years they quietly drain wealth.", "Finance"),
    ("Build additional streams of income so your security never rests on one.", "Finance"),
    ("Staying invested steadily beats trying to guess short-term market swings.", "Finance"),
    ("A clear budget directs your money with purpose instead of guesswork.", "Finance"),
    ("Financial independence grows from learning the basics and practicing restraint.", "Finance"),
    ("Wealth depends less on what you earn and more on what you keep and invest.", "Finance"),
    ("Protect your capital first; avoid reckless bets you cannot afford to lose.", "Finance"),
    ("Avoid spending hard-earned money merely to impress people around you.", "Finance"),
    ("Spreading your investments across asset classes cushions unexpected shocks.", "Finance"),
    ("Lasting wealth is built quietly through discipline, not overnight hype.", "Finance"),
    ("The best time to start investing was years ago; the next best time is today.", "Finance"),
    ("Manage your finances proactively so money remains a tool, not a master.", "Finance"),
    ("Living comfortably below your means is the surest path to financial calm.", "Finance"),
    ("Ignore daily market noise and keep your eyes on decade-long horizons.", "Finance"),

    # Good Moral Values (20)
    ("True integrity is choosing the honest path even when nobody is looking.", "Good Moral Values"),
    ("Simple kindness speaks a universal language everyone understands.", "Good Moral Values"),
    ("Truthfulness is the cornerstone of lasting respect and wisdom.", "Good Moral Values"),
    ("Treat others with courtesy because of your own character, not theirs.", "Good Moral Values"),
    ("Judge a person's character by how they treat those who cannot repay them.", "Good Moral Values"),
    ("Practicing gratitude transforms ordinary moments into abundance.", "Good Moral Values"),
    ("Genuine humility is staying grounded regardless of how much you achieve.", "Good Moral Values"),
    ("One thoughtful act of generosity can ripple farther than you ever see.", "Good Moral Values"),
    ("Earn respect through fairness, keep trust through honesty, and show loyalty.", "Good Moral Values"),
    ("Stand by what is right, even when it is inconvenient or unpopular.", "Good Moral Values"),
    ("Letting go of resentment frees your future from the weight of the past.", "Good Moral Values"),
    ("Choose words that heal, clarify, or uplift; otherwise, value silence.", "Good Moral Values"),
    ("Compassion for others is the truest measure of ethical strength.", "Good Moral Values"),
    ("Patience in trying moments yields the sweetest long-term rewards.", "Good Moral Values"),
    ("Model the fairness and decency you hope to see in society.", "Good Moral Values"),
    ("Real growth is striving to be kinder and wiser than you were yesterday.", "Good Moral Values"),
    ("Good manners and genuine respect open doors that status alone cannot.", "Good Moral Values"),
    ("Empathy begins when you pause to see the world from another's perspective.", "Good Moral Values"),
    ("Do good quietly without expecting applause or recognition.", "Good Moral Values"),
    ("Harmony starts with a calm mind, a gentle word, and a forgiving heart.", "Good Moral Values"),
]


# ==============================================================================
# WINDOWS DARK / LIGHT MODE DETECTION
# ==============================================================================
def is_windows_dark_mode():
    """Return True if Windows AppsUseLightTheme is 0 (Dark Mode), else False (Light Mode)."""
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, THEME_REG_PATH) as key:
            val, _ = winreg.QueryValueEx(key, "AppsUseLightTheme")
            return val == 0
    except Exception:
        return False


# ==============================================================================
# PERSISTENT INSTALLATION STATE HELPERS
# ==============================================================================
def save_install_state(install_dir):
    try:
        marker_path = os.path.join(install_dir, INSTALL_MARKER_FILENAME)
        with open(marker_path, "w", encoding="utf-8") as f:
            json.dump({"installed": True, "install_dir": install_dir}, f, indent=2)
    except Exception:
        pass

    try:
        os.makedirs(LOCAL_STATE_DIR, exist_ok=True)
        with open(INSTALL_STATE_FILE, "w", encoding="utf-8") as f:
            json.dump({"installed": True, "install_dir": install_dir}, f, indent=2)
    except Exception:
        pass


def load_saved_install_dir():
    local_marker = os.path.join(BASE_DIR, INSTALL_MARKER_FILENAME)
    local_cfg = os.path.join(BASE_DIR, CONFIG_FILENAME)
    if os.path.isfile(local_marker) and os.path.isfile(local_cfg):
        return BASE_DIR

    if os.path.isfile(INSTALL_STATE_FILE):
        try:
            with open(INSTALL_STATE_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
            saved_dir = data.get("install_dir", "")
            if (
                data.get("installed")
                and saved_dir
                and os.path.isdir(saved_dir)
                and os.path.isfile(os.path.join(saved_dir, CONFIG_FILENAME))
            ):
                return saved_dir
        except Exception:
            pass

    return None


def clear_install_state():
    try:
        if os.path.isfile(INSTALL_STATE_FILE):
            os.remove(INSTALL_STATE_FILE)
        if os.path.isdir(LOCAL_STATE_DIR) and not os.listdir(LOCAL_STATE_DIR):
            os.rmdir(LOCAL_STATE_DIR)
    except Exception:
        pass


# ==============================================================================
# EXCEL VALIDATION, AUTO-GENERATION & SHORTCUT HELPERS
# ==============================================================================
def get_bundled_resource(filename):
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        candidate = os.path.join(sys._MEIPASS, filename)
        if os.path.isfile(candidate):
            return candidate
    candidate = os.path.join(BASE_DIR, filename)
    return candidate if os.path.isfile(candidate) else None


def is_valid_excel_format(filepath):
    if not filepath or not os.path.isfile(filepath):
        return False
    try:
        wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
        ws = wb.active
        rows = list(ws.iter_rows(min_row=1, max_row=3, values_only=True))
        wb.close()
        if len(rows) < 2 or len(rows[0]) < 2:
            return False
        col_a = str(rows[0][0] or "").strip().lower()
        col_b = str(rows[0][1] or "").strip().lower()
        if col_a != "quote" or col_b != "category":
            return False
        return bool(rows[1][0] and str(rows[1][0]).strip())
    except Exception:
        return False


def create_default_excel_file(filepath):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Quotes"

    header_fill = PatternFill(start_color="2C3E50", end_color="2C3E50", fill_type="solid")
    header_font = Font(name="Calibri", size=12, bold=True, color="FFFFFF")

    ws["A1"] = "Quote"
    ws["B1"] = "Category"
    for cell_ref in ("A1", "B1"):
        ws[cell_ref].fill = header_fill
        ws[cell_ref].font = header_font
        ws[cell_ref].alignment = Alignment(horizontal="left", vertical="center")

    for row_idx, (quote_text, category_name) in enumerate(DEFAULT_QUOTES, start=2):
        ws.cell(row=row_idx, column=1, value=quote_text).alignment = Alignment(
            wrap_text=True, vertical="center"
        )
        ws.cell(row=row_idx, column=2, value=category_name).alignment = Alignment(
            vertical="center"
        )

    ws.column_dimensions["A"].width = 85
    ws.column_dimensions["B"].width = 24
    ws.freeze_panes = "A2"
    wb.save(filepath)


def get_desktop_shortcut_path():
    desktop_dir = os.path.join(os.path.expanduser("~"), "Desktop")
    return os.path.join(desktop_dir, f"{APP_DISPLAY_NAME}.lnk")


def get_start_menu_shortcut_path():
    appdata = os.environ.get("APPDATA", os.path.expanduser("~"))
    programs_dir = os.path.join(appdata, "Microsoft", "Windows", "Start Menu", "Programs")
    return os.path.join(programs_dir, f"{APP_DISPLAY_NAME}.lnk")


def create_windows_shortcut(shortcut_path, target_exe, icon_path=None):
    os.makedirs(os.path.dirname(shortcut_path), exist_ok=True)
    work_dir = os.path.dirname(target_exe)

    ps_commands = [
        '$WshShell = New-Object -ComObject WScript.Shell',
        f'$Shortcut = $WshShell.CreateShortcut("{shortcut_path}")',
        f'$Shortcut.TargetPath = "{target_exe}"',
        f'$Shortcut.WorkingDirectory = "{work_dir}"',
    ]
    if icon_path and os.path.isfile(icon_path):
        ps_commands.append(f'$Shortcut.IconLocation = "{icon_path}"')
    ps_commands.append('$Shortcut.Save()')

    subprocess.run(
        ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", "; ".join(ps_commands)],
        check=False,
        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
    )


def create_all_shortcuts(target_exe, icon_path=None):
    create_windows_shortcut(get_desktop_shortcut_path(), target_exe, icon_path)
    create_windows_shortcut(get_start_menu_shortcut_path(), target_exe, icon_path)


def create_exit_menu_icon():
    size = 64
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    draw.ellipse([4, 4, size - 4, size - 4], fill=(220, 38, 38, 255))
    pad = 19
    draw.line([(pad, pad), (size - pad, size - pad)], fill=(255, 255, 255, 255), width=7)
    draw.line([(size - pad, pad), (pad, size - pad)], fill=(255, 255, 255, 255), width=7)
    img = img.resize((16, 16), Image.LANCZOS)
    return ImageTk.PhotoImage(img)


def get_startup_command(target_exe=None):
    if target_exe:
        return f'"{os.path.abspath(target_exe)}"'
    if getattr(sys, "frozen", False):
        return f'"{os.path.abspath(sys.executable)}"'
    pythonw = sys.executable.replace("python.exe", "pythonw.exe")
    if not os.path.exists(pythonw):
        pythonw = sys.executable
    return f'"{pythonw}" "{os.path.abspath(__file__)}"'


def set_windows_startup(enable=True, target_exe=None):
    try:
        with winreg.OpenKey(
            winreg.HKEY_CURRENT_USER, RUN_REG_PATH, 0, winreg.KEY_SET_VALUE
        ) as key:
            if enable:
                winreg.SetValueEx(
                    key, APP_REG_NAME, 0, winreg.REG_SZ, get_startup_command(target_exe)
                )
            else:
                try:
                    winreg.DeleteValue(key, APP_REG_NAME)
                except FileNotFoundError:
                    pass
    except Exception:
        pass


def hex_to_rgb(hex_color):
    hex_color = hex_color.lstrip("#")
    if len(hex_color) != 6:
        return (255, 255, 255)
    return tuple(int(hex_color[i : i + 2], 16) for i in (0, 2, 4))


def rgb_to_hex(r, g, b):
    r = max(0, min(255, int(r)))
    g = max(0, min(255, int(g)))
    b = max(0, min(255, int(b)))
    return f"#{r:02x}{g:02x}{b:02x}"


def discover_windows_fonts():
    font_dirs = [
        os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts"),
        os.path.expanduser(r"~\AppData\Local\Microsoft\Windows\Fonts"),
    ]
    fonts = {}
    for fdir in font_dirs:
        if not os.path.isdir(fdir):
            continue
        for fname in os.listdir(fdir):
            if fname.lower().endswith((".ttf", ".otf", ".ttc")):
                label = os.path.splitext(fname)[0].replace("_", " ").title()
                fonts[label] = os.path.join(fdir, fname)
    return dict(sorted(fonts.items()))


def detect_script_family(text):
    for ch in text:
        cp = ord(ch)
        if 0x0900 <= cp <= 0x0DFF:
            return "indic"
        if 0x0590 <= cp <= 0x06FF or 0x0750 <= cp <= 0x077F or 0xFB50 <= cp <= 0xFDFF:
            return "rtl"
        if (
            0x3040 <= cp <= 0x30FF
            or 0x4E00 <= cp <= 0x9FFF
            or 0xAC00 <= cp <= 0xD7AF
        ):
            return "cjk"
    return "latin"


# ==============================================================================
# MODE-CATEGORIZED PRESETS (Dark Mode: 28 Presets | Light Mode: 24 Presets)
# ==============================================================================
DARK_MODE_PRESETS = {
    "Executive Slate": {
        "font_name": "Calibri", "font_size": 20, "text_color": "#F1F5F9", "text_alpha": 1.0,
        "text_fill": "Solid", "text_shadow": True, "bg_color": "#1E293B", "bg_alpha": 0.90,
        "bg_fill": "Solid", "bg_shape": "Chamfered", "border_enabled": True, "border_color": "#94A3B8", "border_pattern": "Solid",
    },
    "Midnight Navy": {
        "font_name": "Georgia", "font_size": 20, "text_color": "#E2E8F0", "text_alpha": 1.0,
        "text_fill": "Solid", "text_shadow": True, "bg_color": "#0F172A", "bg_alpha": 0.92,
        "bg_fill": "Solid", "bg_shape": "Rounded", "border_enabled": True, "border_color": "#CBD5E1", "border_pattern": "Solid",
    },
    "Velvet Heritage": {
        "font_name": "Georgia", "font_size": 22, "text_color": "#F5D061", "text_alpha": 1.0,
        "text_fill": "Glitter", "text_shadow": True, "bg_color": "#3B0A18", "bg_alpha": 0.92,
        "bg_fill": "Dots", "bg_shape": "Rounded", "border_enabled": True, "border_color": "#D4AF37", "border_pattern": "Glitter",
    },
    "Oxford Library": {
        "font_name": "Georgia", "font_size": 21, "text_color": "#F5EFE6", "text_alpha": 1.0,
        "text_fill": "Solid", "text_shadow": True, "bg_color": "#1B4332", "bg_alpha": 0.92,
        "bg_fill": "Solid", "bg_shape": "Rounded", "border_enabled": True, "border_color": "#D4AF37", "border_pattern": "Triple",
    },
    "Antique Brass": {
        "font_name": "Georgia", "font_size": 21, "text_color": "#E6C77A", "text_alpha": 1.0,
        "text_fill": "Solid", "text_shadow": True, "bg_color": "#261C15", "bg_alpha": 0.92,
        "bg_fill": "Solid", "bg_shape": "Chamfered", "border_enabled": True, "border_color": "#B8860B", "border_pattern": "Double",
    },
    "Silver Linings": {
        "font_name": "Calibri", "font_size": 20, "text_color": "#E5E7EB", "text_alpha": 0.96,
        "text_fill": "Solid", "text_shadow": True, "bg_color": "#27272A", "bg_alpha": 0.90,
        "bg_fill": "Solid", "bg_shape": "Leaf", "border_enabled": True, "border_color": "#A1A1AA", "border_pattern": "Dotted",
    },
    "Chalk & Slate": {
        "font_name": "Georgia", "font_size": 20, "text_color": "#F3F4F6", "text_alpha": 0.96,
        "text_fill": "Solid", "text_shadow": True, "bg_color": "#2D3138", "bg_alpha": 0.92,
        "bg_fill": "Solid", "bg_shape": "Rounded", "border_enabled": True, "border_color": "#D1D5DB", "border_pattern": "Dashed",
    },
    "Emerald Velvet": {
        "font_name": "Georgia", "font_size": 21, "text_color": "#FDE68A", "text_alpha": 1.0,
        "text_fill": "Glitter", "text_shadow": True, "bg_color": "#064E3B", "bg_alpha": 0.92,
        "bg_fill": "Dots", "bg_shape": "Scalloped", "border_enabled": True, "border_color": "#F59E0B", "border_pattern": "Double",
    },
    "Pure Dark Minimal": {
        "font_name": "Georgia", "font_size": 20, "text_color": "#FFFFFF", "text_alpha": 0.96,
        "text_fill": "Solid", "text_shadow": True, "bg_color": "#000000", "bg_alpha": 0.0,
        "bg_fill": "Solid", "bg_shape": "Rounded", "border_enabled": False, "border_color": "#FFFFFF", "border_pattern": "Solid",
    },
    "Obsidian Gold": {
        "font_name": "Georgia", "font_size": 21, "text_color": "#FCD34D", "text_alpha": 1.0,
        "text_fill": "Solid", "text_shadow": True, "bg_color": "#09090B", "bg_alpha": 0.94,
        "bg_fill": "Solid", "bg_shape": "Chamfered", "border_enabled": True, "border_color": "#D97706", "border_pattern": "Double",
    },
    "Ocean Breeze": {
        "font_name": "Segoeui", "font_size": 20, "text_color": "#E0F2FE", "text_alpha": 1.0,
        "text_fill": "Solid", "text_shadow": True, "bg_color": "#0369A1", "bg_alpha": 0.88,
        "bg_fill": "Gradient", "bg_shape": "Pill", "border_enabled": True, "border_color": "#7DD3FC", "border_pattern": "Glow",
    },
    "Desert Sunset": {
        "font_name": "Georgia", "font_size": 21, "text_color": "#FEF3C7", "text_alpha": 1.0,
        "text_fill": "Solid", "text_shadow": True, "bg_color": "#9A3412", "bg_alpha": 0.90,
        "bg_fill": "Gradient", "bg_shape": "Rounded", "border_enabled": True, "border_color": "#FDBA74", "border_pattern": "Double",
    },
    "Forest Mist": {
        "font_name": "Calibri", "font_size": 20, "text_color": "#D1FAE5", "text_alpha": 1.0,
        "text_fill": "Solid", "text_shadow": True, "bg_color": "#14532D", "bg_alpha": 0.90,
        "bg_fill": "Solid", "bg_shape": "Leaf", "border_enabled": True, "border_color": "#6EE7B7", "border_pattern": "Solid",
    },
    "Synthwave Horizon": {
        "font_name": "Impact", "font_size": 21, "text_color": "#38BDF8", "text_alpha": 1.0,
        "text_fill": "Gradient", "text_shadow": True, "bg_color": "#2E1065", "bg_alpha": 0.90,
        "bg_fill": "Gradient", "bg_shape": "Chamfered", "border_enabled": True, "border_color": "#EC4899", "border_pattern": "Neon",
    },
    "Blueprint Grid": {
        "font_name": "Consola", "font_size": 19, "text_color": "#FFFFFF", "text_alpha": 1.0,
        "text_fill": "Solid", "text_shadow": True, "bg_color": "#0B3D91", "bg_alpha": 0.92,
        "bg_fill": "Dots", "bg_shape": "Rectangle", "border_enabled": True, "border_color": "#93C5FD", "border_pattern": "Dashed",
    },
    "Amber Terminal": {
        "font_name": "Consola", "font_size": 20, "text_color": "#FFB000", "text_alpha": 1.0,
        "text_fill": "Solid", "text_shadow": True, "bg_color": "#1A1000", "bg_alpha": 0.92,
        "bg_fill": "Stripes", "bg_shape": "Chamfered", "border_enabled": True, "border_color": "#FFB000", "border_pattern": "Solid",
    },
    "Deep Space Nebula": {
        "font_name": "Segoeui", "font_size": 20, "text_color": "#F8FAFC", "text_alpha": 1.0,
        "text_fill": "Glitter", "text_shadow": True, "bg_color": "#0F0728", "bg_alpha": 0.92,
        "bg_fill": "Glitter", "bg_shape": "Cloud", "border_enabled": True, "border_color": "#C084FC", "border_pattern": "Rainbow",
    },
    "Steampunk Copper": {
        "font_name": "Georgia", "font_size": 21, "text_color": "#FDBA74", "text_alpha": 1.0,
        "text_fill": "Solid", "text_shadow": True, "bg_color": "#291811", "bg_alpha": 0.94,
        "bg_fill": "Stripes", "bg_shape": "Ticket", "border_enabled": True, "border_color": "#B45309", "border_pattern": "Triple",
    },
    "Cyber Pulse": {
        "font_name": "Impact", "font_size": 22, "text_color": "#00F5FF", "text_alpha": 1.0,
        "text_fill": "Gradient", "text_shadow": True, "bg_color": "#1A0B2E", "bg_alpha": 0.90,
        "bg_fill": "Gradient", "bg_shape": "Chamfered", "border_enabled": True, "border_color": "#FF007F", "border_pattern": "Neon",
    },
    "Acid Lime Street": {
        "font_name": "Arialbd", "font_size": 21, "text_color": "#CCFF00", "text_alpha": 1.0,
        "text_fill": "Solid", "text_shadow": True, "bg_color": "#0E0E0E", "bg_alpha": 0.92,
        "bg_fill": "Stripes", "bg_shape": "Banner", "border_enabled": True, "border_color": "#CCFF00", "border_pattern": "Dashed",
    },
    "Lo-Fi Midnight": {
        "font_name": "Calibri", "font_size": 20, "text_color": "#E2D9F3", "text_alpha": 0.96,
        "text_fill": "Solid", "text_shadow": True, "bg_color": "#1F1D36", "bg_alpha": 0.88,
        "bg_fill": "Dots", "bg_shape": "Rounded", "border_enabled": True, "border_color": "#864879", "border_pattern": "Dotted",
    },
    "Electric Cobalt": {
        "font_name": "Arialbd", "font_size": 21, "text_color": "#FDE047", "text_alpha": 1.0,
        "text_fill": "Solid", "text_shadow": True, "bg_color": "#1D4ED8", "bg_alpha": 0.92,
        "bg_fill": "Solid", "bg_shape": "Banner", "border_enabled": True, "border_color": "#FDE047", "border_pattern": "Thick",
    },
    "Digital Rain": {
        "font_name": "Consola", "font_size": 20, "text_color": "#00FF41", "text_alpha": 1.0,
        "text_fill": "Solid", "text_shadow": True, "bg_color": "#001405", "bg_alpha": 0.86,
        "bg_fill": "Stripes", "bg_shape": "Chamfered", "border_enabled": True, "border_color": "#00FF41", "border_pattern": "Dashed",
    },
    "Retro Arcade 84": {
        "font_name": "Verdana", "font_size": 20, "text_color": "#FFDD00", "text_alpha": 1.0,
        "text_fill": "Gradient", "text_shadow": True, "bg_color": "#2C0735", "bg_alpha": 0.90,
        "bg_fill": "Stripes", "bg_shape": "Ticket", "border_enabled": True, "border_color": "#FF6EC7", "border_pattern": "Neon",
    },
    "Cafe Chalkboard": {
        "font_name": "Comic", "font_size": 21, "text_color": "#F4F1DE", "text_alpha": 0.95,
        "text_fill": "Solid", "text_shadow": True, "bg_color": "#262A2C", "bg_alpha": 0.92,
        "bg_fill": "Solid", "bg_shape": "Ticket", "border_enabled": True, "border_color": "#E07A5F", "border_pattern": "Dashed",
    },
    "Desi Shaadi Sparkle": {
        "font_name": "Comicbd", "font_size": 23, "text_color": "#FFD700", "text_alpha": 1.0,
        "text_fill": "Glitter", "text_shadow": True, "bg_color": "#C2185B", "bg_alpha": 0.90,
        "bg_fill": "Glitter", "bg_shape": "Ribbon", "border_enabled": True, "border_color": "#FFEB3B", "border_pattern": "Glitter",
    },
    "Marigold Utsav": {
        "font_name": "Georgia", "font_size": 21, "text_color": "#FFF176", "text_alpha": 1.0,
        "text_fill": "Glitter", "text_shadow": True, "bg_color": "#8E1616", "bg_alpha": 0.92,
        "bg_fill": "Dots", "bg_shape": "Scalloped", "border_enabled": True, "border_color": "#FF9800", "border_pattern": "Double",
    },
    "Peacock Jewel": {
        "font_name": "Times", "font_size": 22, "text_color": "#7DF9FF", "text_alpha": 1.0,
        "text_fill": "Glitter", "text_shadow": True, "bg_color": "#004D40", "bg_alpha": 0.90,
        "bg_fill": "Gradient", "bg_shape": "Leaf", "border_enabled": True, "border_color": "#FFD700", "border_pattern": "Glow",
    },
}

LIGHT_MODE_PRESETS = {
    "Royal Parchment": {
        "font_name": "Georgia", "font_size": 21, "text_color": "#3E2723", "text_alpha": 1.0,
        "text_fill": "Solid", "text_shadow": False, "bg_color": "#F3E5AB", "bg_alpha": 0.94,
        "bg_fill": "Solid", "bg_shape": "Scalloped", "border_enabled": True, "border_color": "#8D6E63", "border_pattern": "Double",
    },
    "Ivory & Charcoal": {
        "font_name": "Times", "font_size": 21, "text_color": "#212121", "text_alpha": 1.0,
        "text_fill": "Solid", "text_shadow": False, "bg_color": "#FAF8F5", "bg_alpha": 0.95,
        "bg_fill": "Solid", "bg_shape": "Rectangle", "border_enabled": True, "border_color": "#424242", "border_pattern": "Solid",
    },
    "Monochrome Ink": {
        "font_name": "Times", "font_size": 21, "text_color": "#111111", "text_alpha": 1.0,
        "text_fill": "Solid", "text_shadow": False, "bg_color": "#FFFFFF", "bg_alpha": 0.94,
        "bg_fill": "Solid", "bg_shape": "Ticket", "border_enabled": True, "border_color": "#111111", "border_pattern": "Double",
    },
    "Marble & Gold": {
        "font_name": "Georgia", "font_size": 20, "text_color": "#2D3748", "text_alpha": 1.0,
        "text_fill": "Solid", "text_shadow": False, "bg_color": "#EDF2F7", "bg_alpha": 0.94,
        "bg_fill": "Dots", "bg_shape": "Pill", "border_enabled": True, "border_color": "#C5A059", "border_pattern": "Gradient",
    },
    "Sepia Chronicle": {
        "font_name": "Times", "font_size": 21, "text_color": "#3B2F2F", "text_alpha": 1.0,
        "text_fill": "Solid", "text_shadow": False, "bg_color": "#E6D7C3", "bg_alpha": 0.94,
        "bg_fill": "Solid", "bg_shape": "Scalloped", "border_enabled": True, "border_color": "#6F4E37", "border_pattern": "Dashed",
    },
    "Sakura Blossom": {
        "font_name": "Georgia", "font_size": 20, "text_color": "#581C4F", "text_alpha": 1.0,
        "text_fill": "Solid", "text_shadow": False, "bg_color": "#FCE7F3", "bg_alpha": 0.92,
        "bg_fill": "Dots", "bg_shape": "Cloud", "border_enabled": True, "border_color": "#F472B6", "border_pattern": "Dotted",
    },
    "Arctic Frost": {
        "font_name": "Segoeui", "font_size": 20, "text_color": "#0F172A", "text_alpha": 1.0,
        "text_fill": "Solid", "text_shadow": False, "bg_color": "#E0F2FE", "bg_alpha": 0.92,
        "bg_fill": "Gradient", "bg_shape": "Chamfered", "border_enabled": True, "border_color": "#38BDF8", "border_pattern": "Glow",
    },
    "Pure Light Minimal": {
        "font_name": "Georgia", "font_size": 20, "text_color": "#111827", "text_alpha": 0.96,
        "text_fill": "Solid", "text_shadow": False, "bg_color": "#FFFFFF", "bg_alpha": 0.0,
        "bg_fill": "Solid", "bg_shape": "Rounded", "border_enabled": False, "border_color": "#111827", "border_pattern": "Solid",
    },
    "Mint Whisper": {
        "font_name": "Calibri", "font_size": 20, "text_color": "#064E3B", "text_alpha": 1.0,
        "text_fill": "Solid", "text_shadow": False, "bg_color": "#ECFDF5", "bg_alpha": 0.94,
        "bg_fill": "Solid", "bg_shape": "Leaf", "border_enabled": True, "border_color": "#34D399", "border_pattern": "Solid",
    },
    "Lavender Mist": {
        "font_name": "Georgia", "font_size": 20, "text_color": "#3B0764", "text_alpha": 1.0,
        "text_fill": "Solid", "text_shadow": False, "bg_color": "#F3E8FF", "bg_alpha": 0.94,
        "bg_fill": "Dots", "bg_shape": "Cloud", "border_enabled": True, "border_color": "#A855F7", "border_pattern": "Dotted",
    },
    "Sunlit Lemon": {
        "font_name": "Segoeui", "font_size": 20, "text_color": "#713F12", "text_alpha": 1.0,
        "text_fill": "Solid", "text_shadow": False, "bg_color": "#FEF9C3", "bg_alpha": 0.95,
        "bg_fill": "Solid", "bg_shape": "Rounded", "border_enabled": True, "border_color": "#EAB308", "border_pattern": "Double",
    },
    "Peach Sorbet": {
        "font_name": "Georgia", "font_size": 20, "text_color": "#7C2D12", "text_alpha": 1.0,
        "text_fill": "Solid", "text_shadow": False, "bg_color": "#FFEDD5", "bg_alpha": 0.94,
        "bg_fill": "Gradient", "bg_shape": "Pill", "border_enabled": True, "border_color": "#FB923C", "border_pattern": "Solid",
    },
    "Sky Paper": {
        "font_name": "Calibri", "font_size": 20, "text_color": "#1E3A8A", "text_alpha": 1.0,
        "text_fill": "Solid", "text_shadow": False, "bg_color": "#DBEAFE", "bg_alpha": 0.94,
        "bg_fill": "Solid", "bg_shape": "Banner", "border_enabled": True, "border_color": "#60A5FA", "border_pattern": "Dashed",
    },
    "Rose Quartz": {
        "font_name": "Georgia", "font_size": 20, "text_color": "#881337", "text_alpha": 1.0,
        "text_fill": "Solid", "text_shadow": False, "bg_color": "#FFE4E6", "bg_alpha": 0.94,
        "bg_fill": "Solid", "bg_shape": "Ribbon", "border_enabled": True, "border_color": "#FB7185", "border_pattern": "Double",
    },
    "Sage Notebook": {
        "font_name": "Times", "font_size": 20, "text_color": "#14532D", "text_alpha": 1.0,
        "text_fill": "Solid", "text_shadow": False, "bg_color": "#DCFCE7", "bg_alpha": 0.94,
        "bg_fill": "Stripes", "bg_shape": "Ticket", "border_enabled": True, "border_color": "#4ADE80", "border_pattern": "Solid",
    },
    "Warm Cream": {
        "font_name": "Georgia", "font_size": 21, "text_color": "#44403C", "text_alpha": 1.0,
        "text_fill": "Solid", "text_shadow": False, "bg_color": "#FFFBEB", "bg_alpha": 0.95,
        "bg_fill": "Solid", "bg_shape": "Rounded", "border_enabled": True, "border_color": "#D97706", "border_pattern": "Solid",
    },
    "Coral Breeze": {
        "font_name": "Segoeui", "font_size": 20, "text_color": "#9A3412", "text_alpha": 1.0,
        "text_fill": "Solid", "text_shadow": False, "bg_color": "#FFF7ED", "bg_alpha": 0.94,
        "bg_fill": "Dots", "bg_shape": "Scalloped", "border_enabled": True, "border_color": "#F97316", "border_pattern": "Dotted",
    },
    "Pearl Gray": {
        "font_name": "Calibri", "font_size": 20, "text_color": "#1F2937", "text_alpha": 1.0,
        "text_fill": "Solid", "text_shadow": False, "bg_color": "#F3F4F6", "bg_alpha": 0.95,
        "bg_fill": "Solid", "bg_shape": "Chamfered", "border_enabled": True, "border_color": "#9CA3AF", "border_pattern": "Double",
    },
    "Aqua Glass": {
        "font_name": "Segoeui", "font_size": 20, "text_color": "#155E75", "text_alpha": 1.0,
        "text_fill": "Solid", "text_shadow": False, "bg_color": "#CFFAFE", "bg_alpha": 0.92,
        "bg_fill": "Gradient", "bg_shape": "Rounded", "border_enabled": True, "border_color": "#22D3EE", "border_pattern": "Glow",
    },
    "Sandstone Classic": {
        "font_name": "Times", "font_size": 21, "text_color": "#292524", "text_alpha": 1.0,
        "text_fill": "Solid", "text_shadow": False, "bg_color": "#F5F5F4", "bg_alpha": 0.95,
        "bg_fill": "Solid", "bg_shape": "Rectangle", "border_enabled": True, "border_color": "#78716C", "border_pattern": "Triple",
    },
    "Buttercup Note": {
        "font_name": "Comic", "font_size": 20, "text_color": "#422006", "text_alpha": 1.0,
        "text_fill": "Solid", "text_shadow": False, "bg_color": "#FEF08A", "bg_alpha": 0.94,
        "bg_fill": "Solid", "bg_shape": "Speech", "border_enabled": True, "border_color": "#CA8A04", "border_pattern": "Thick",
    },
    "Lilac Cloud": {
        "font_name": "Georgia", "font_size": 20, "text_color": "#4C1D95", "text_alpha": 1.0,
        "text_fill": "Solid", "text_shadow": False, "bg_color": "#EDE9FE", "bg_alpha": 0.94,
        "bg_fill": "Solid", "bg_shape": "Cloud", "border_enabled": True, "border_color": "#8B5CF6", "border_pattern": "Solid",
    },
    "Pistachio Fresh": {
        "font_name": "Calibri", "font_size": 20, "text_color": "#365314", "text_alpha": 1.0,
        "text_fill": "Solid", "text_shadow": False, "bg_color": "#ECFCCB", "bg_alpha": 0.94,
        "bg_fill": "Dots", "bg_shape": "Leaf", "border_enabled": True, "border_color": "#84CC16", "border_pattern": "Dashed",
    },
    "Porcelain Blue": {
        "font_name": "Georgia", "font_size": 20, "text_color": "#1E40AF", "text_alpha": 1.0,
        "text_fill": "Solid", "text_shadow": False, "bg_color": "#EFF6FF", "bg_alpha": 0.95,
        "bg_fill": "Solid", "bg_shape": "Ticket", "border_enabled": True, "border_color": "#3B82F6", "border_pattern": "Double",
    },
}

PRESET_CATEGORIES = {
    "Dark Mode Presets": DARK_MODE_PRESETS,
    "Light Mode Presets": LIGHT_MODE_PRESETS,
}

DARK_PRESETS_LIST = list(DARK_MODE_PRESETS.values())
LIGHT_PRESETS_LIST = list(LIGHT_MODE_PRESETS.values())


class QuoteWidget:
    def __init__(self, mutex_handle=None):
        self.mutex_handle = mutex_handle
        self.root = tk.Tk()
        self.root.withdraw()
        self.available_fonts = discover_windows_fonts()

        # Check if already installed (via .installed marker or %LOCALAPPDATA% pointer)
        saved_install_dir = load_saved_install_dir()
        if saved_install_dir:
            self.install_dir = saved_install_dir
            self.config_file = os.path.join(self.install_dir, CONFIG_FILENAME)
            self.config = self.load_config(self.config_file)
            self.config["disclaimer_accepted"] = True
            self.config["installed"] = True
        else:
            self.install_dir = BASE_DIR
            self.config_file = CONFIG_FILE
            self.config = self.load_config(self.config_file)

        # ------------------------------------------------------------------
        # FIRST-RUN INSTALLATION FLOW (Runs ONCE; completely skipped on restarts)
        # ------------------------------------------------------------------
        if not self.config.get("installed", False) and not self.config.get("disclaimer_accepted", False):
            if not self.prompt_disclaimer():
                self.root.destroy()
                return

            chosen_dir = self.prompt_install_folder()
            if not chosen_dir:
                self.root.destroy()
                return

            self.install_dir = chosen_dir
            self.config_file = os.path.join(self.install_dir, CONFIG_FILENAME)
            excel_path = os.path.join(self.install_dir, EXCEL_FILENAME)

            os.makedirs(self.install_dir, exist_ok=True)
            if is_valid_excel_format(excel_path):
                retain_existing = messagebox.askyesno(
                    "Existing Quotes File Found",
                    f"A properly formatted quotes.xlsx file was found in:\n{self.install_dir}\n\n"
                    "• Click 'Yes' to RETAIN your existing quotes.xlsx file.\n"
                    "• Click 'No' to OVERWRITE it with the default quotes.",
                    parent=self.root,
                )
                if not retain_existing:
                    create_default_excel_file(excel_path)
            else:
                create_default_excel_file(excel_path)

            self.config["disclaimer_accepted"] = True
            self.config["installed"] = True
            self.config["install_dir"] = self.install_dir
            self.config["excel_file"] = excel_path
            self.save_config()
            save_install_state(self.install_dir)

            if getattr(sys, "frozen", False):
                current_exe = os.path.abspath(sys.executable)
                dest_exe = os.path.join(self.install_dir, os.path.basename(current_exe))
                if os.path.abspath(dest_exe) != current_exe:
                    shutil.copy2(current_exe, dest_exe)

                src_icon = get_bundled_resource(ICON_FILENAME)
                dest_icon = None
                if src_icon:
                    dest_icon = os.path.join(self.install_dir, ICON_FILENAME)
                    if os.path.abspath(src_icon) != os.path.abspath(dest_icon):
                        shutil.copy2(src_icon, dest_icon)

                create_all_shortcuts(dest_exe, dest_icon or dest_exe)
                set_windows_startup(self.config.get("run_at_startup", True), target_exe=dest_exe)

                if os.path.abspath(dest_exe) != current_exe:
                    self.root.destroy()
                    release_single_instance_lock(self.mutex_handle)
                    self.mutex_handle = None
                    subprocess.Popen([dest_exe], cwd=self.install_dir)
                    return
        else:
            self.install_dir = self.config.get("install_dir", self.install_dir) or self.install_dir
            excel_path = self.config.get("excel_file", os.path.join(self.install_dir, EXCEL_FILENAME))
            if not is_valid_excel_format(excel_path):
                os.makedirs(os.path.dirname(excel_path) or self.install_dir, exist_ok=True)
                create_default_excel_file(excel_path)
            self.config["excel_file"] = excel_path
            self.config["installed"] = True
            self.save_config()
            save_install_state(self.install_dir)
            set_windows_startup(self.config.get("run_at_startup", True))

        # ------------------------------------------------------------------
        # Load categorized quotes. Prompt for categories ONLY on initial setup.
        # ------------------------------------------------------------------
        self.quote_entries = self.load_quote_entries()
        self.all_categories = sorted({cat for _, cat in self.quote_entries})

        if not self.config.get("categories_initialized", False):
            chosen_cats = self.prompt_category_selection(self.all_categories)
            if not chosen_cats:
                self.root.destroy()
                return
            self.config["selected_categories"] = chosen_cats
            self.config["categories_initialized"] = True
            self.save_config()
        else:
            saved_cats = self.config.get("selected_categories", [])
            valid_saved = [c for c in saved_cats if c in self.all_categories]
            chosen_cats = valid_saved if valid_saved else list(self.all_categories)

        self.category_vars = {
            cat: tk.BooleanVar(value=(cat in chosen_cats)) for cat in self.all_categories
        }

        self.quotes = self.get_filtered_quotes()
        self.current_quote = ""
        self.quote_history = []
        self.history_pos = -1
        self.undo_stack = []
        self.is_paused = False
        self._paused_remaining_secs = None

        self.photo_ref = None
        self.exit_icon = create_exit_menu_icon()
        self._last_final_rgba = None
        self._last_chroma_img = None
        self._render_seed = 42
        self._after_id = None
        self._anim_id = None
        self._next_change_ts = None

        self._button_rects = []
        self._top_bar_h = 26
        self._controls_visible = False
        self._hide_controls_after_id = None
        self._tooltip_win = None
        self._hovered_btn = None
        self._toast_win = None
        self._toast_after_id = None

        # Frameless shaped window using chroma-key transparency
        self.root.overrideredirect(True)
        self.root.configure(bg=CHROMA_KEY)
        self.root.attributes("-transparentcolor", CHROMA_KEY)
        self.root.geometry(f"+{self.config['x']}+{self.config['y']}")
        self.root.protocol("WM_DELETE_WINDOW", lambda: None)

        self.display = tk.Label(self.root, bg=CHROMA_KEY, bd=0, highlightthickness=0)
        self.display.pack()

        # Click-to-next-quote, Drag-to-move, and Drag-to-resize state
        self._interaction_mode = "move"
        self._did_drag = False
        self._drag_x = 0
        self._drag_y = 0
        self._resize_start_x = 0
        self._resize_start_y = 0
        self._resize_orig_w = 560
        self._resize_orig_fs = 20
        self._last_resize_ts = 0.0

        self.display.bind("<Motion>", self.on_hover)
        self.display.bind("<Leave>", self.on_mouse_leave)
        self.display.bind("<Button-1>", self.on_mouse_down)
        self.display.bind("<B1-Motion>", self.on_mouse_drag)
        self.display.bind("<ButtonRelease-1>", self.on_mouse_up)
        self.root.bind("<FocusIn>", self.on_focus_in)

        # Menu state variables
        self.order_var = tk.StringVar(value=self.config["order_mode"])
        self.window_layer_var = tk.StringVar(value=self.config["window_layer"])
        self.text_align_var = tk.StringVar(value=self.config["text_align"])
        self.theme_mode_var = tk.StringVar(value=self.config.get("theme_mode", "light"))
        self.startup_var = tk.BooleanVar(value=self.config["run_at_startup"])
        self.randomize_style_var = tk.BooleanVar(value=self.config["randomize_style"])
        self.anim_enabled_var = tk.BooleanVar(value=self.config["anim_enabled"])
        self.anim_mode_var = tk.StringVar(value=self.config["anim_mode"])
        self.shadow_var = tk.BooleanVar(value=self.config["text_shadow"])
        self.border_var = tk.BooleanVar(value=self.config["border_enabled"])

        # Track all menus for dynamic Light/Dark theming
        self._all_menus = []

        # ==================================================================
        # GROUPED RIGHT-CLICK CONTEXT MENU
        # ==================================================================
        self.menu = self._make_menu(self.root)

        # 1st Option: Countdown showing remaining time
        self.menu.add_command(label=self.get_interval_display_text(), state="disabled")
        self.menu.add_separator()

        # Standalone quick action
        self.menu.add_command(label="Next Quote", command=lambda: self.show_next_quote(animate=True, show_toast=True))
        self.menu.add_separator()

        # --- Group 1: Quote Settings ---
        quote_settings_menu = self._make_menu(self.menu)
        quote_settings_menu.add_command(
            label="Set Quote Change Interval...", command=self.open_interval_dialog
        )
        quote_settings_menu.add_command(
            label="Set Undo History Limit (1-20)...", command=self.open_undo_limit_dialog
        )

        self.category_menu = self._make_menu(quote_settings_menu)
        self.rebuild_category_menu()
        quote_settings_menu.add_cascade(label="Quote Categories", menu=self.category_menu)

        order_menu = self._make_menu(quote_settings_menu)
        order_menu.add_radiobutton(
            label="Random", variable=self.order_var, value="random", command=self.set_order_mode
        )
        order_menu.add_radiobutton(
            label="In Sequence",
            variable=self.order_var,
            value="sequential",
            command=self.set_order_mode,
        )
        quote_settings_menu.add_cascade(label="Quote Order", menu=order_menu)

        anim_menu = self._make_menu(quote_settings_menu)
        anim_menu.add_checkbutton(
            label="Enable Animation",
            variable=self.anim_enabled_var,
            command=self.toggle_anim_enabled,
        )
        anim_menu.add_separator()
        anim_menu.add_radiobutton(
            label="Randomize",
            variable=self.anim_mode_var,
            value="Random",
            command=self.set_anim_mode,
        )
        for anim_name in ANIMATION_TYPES:
            anim_menu.add_radiobutton(
                label=anim_name,
                variable=self.anim_mode_var,
                value=anim_name,
                command=self.set_anim_mode,
            )
        quote_settings_menu.add_cascade(label="Quote Animation", menu=anim_menu)

        quote_settings_menu.add_separator()
        quote_settings_menu.add_command(label="Add New Quote...", command=self.open_add_quote_dialog)
        quote_settings_menu.add_command(label="Edit Excel File...", command=self.edit_excel_file)
        self.menu.add_cascade(label="Quote Settings", menu=quote_settings_menu)

        # --- Group 2: Text Settings ---
        text_settings_menu = self._make_menu(self.menu)
        text_settings_menu.add_command(
            label="Text Style (Font, Color, Fill, Opacity)...",
            command=self.open_text_dialog,
        )

        align_menu = self._make_menu(text_settings_menu)
        align_menu.add_radiobutton(
            label="Center", variable=self.text_align_var, value="center", command=self.set_text_align
        )
        align_menu.add_radiobutton(
            label="Left", variable=self.text_align_var, value="left", command=self.set_text_align
        )
        align_menu.add_radiobutton(
            label="Right", variable=self.text_align_var, value="right", command=self.set_text_align
        )
        text_settings_menu.add_cascade(label="Text Alignment", menu=align_menu)

        text_settings_menu.add_checkbutton(
            label="Text Shadow", variable=self.shadow_var, command=self.toggle_shadow
        )
        self.menu.add_cascade(label="Text Settings", menu=text_settings_menu)

        # --- Group 3: Background & Border ---
        bg_settings_menu = self._make_menu(self.menu)
        bg_settings_menu.add_command(
            label="Background & Border (Shape, Color, Fill, Opacity)...",
            command=self.open_bg_dialog,
        )
        bg_settings_menu.add_checkbutton(
            label="Show Border", variable=self.border_var, command=self.toggle_border
        )
        bg_settings_menu.add_separator()

        # Mode-categorized Presets submenu (Dark Mode Presets & Light Mode Presets)
        preset_root_menu = self._make_menu(bg_settings_menu)
        for cat_name, presets_dict in PRESET_CATEGORIES.items():
            cat_menu = self._make_menu(preset_root_menu)
            for preset_name, preset_data in presets_dict.items():
                cat_menu.add_command(
                    label=preset_name,
                    command=lambda data=preset_data: self.apply_preset(data),
                )
            preset_root_menu.add_cascade(label=cat_name, menu=cat_menu)
        bg_settings_menu.add_cascade(label="Presets", menu=preset_root_menu)

        bg_settings_menu.add_checkbutton(
            label="Randomize Style on Quote Change",
            variable=self.randomize_style_var,
            command=self.toggle_randomize_style,
        )
        self.menu.add_cascade(label="Background & Border", menu=bg_settings_menu)

        # --- Appearance, Display Position & Startup ---
        self.menu.add_separator()

        theme_mode_menu = self._make_menu(self.menu)
        theme_mode_menu.add_radiobutton(
            label="System (Follow Windows)",
            variable=self.theme_mode_var,
            value="system",
            command=self.set_theme_mode,
        )
        theme_mode_menu.add_radiobutton(
            label="Dark Mode",
            variable=self.theme_mode_var,
            value="dark",
            command=self.set_theme_mode,
        )
        theme_mode_menu.add_radiobutton(
            label="Light Mode",
            variable=self.theme_mode_var,
            value="light",
            command=self.set_theme_mode,
        )
        self.menu.add_cascade(label="Appearance Mode", menu=theme_mode_menu)

        layer_menu = self._make_menu(self.menu)
        layer_menu.add_radiobutton(
            label="Show Only on Desktop",
            variable=self.window_layer_var,
            value="desktop",
            command=self.set_window_layer,
        )
        layer_menu.add_radiobutton(
            label="Show Above All Windows",
            variable=self.window_layer_var,
            value="topmost",
            command=self.set_window_layer,
        )
        self.menu.add_cascade(label="Display Position", menu=layer_menu)

        self.menu.add_checkbutton(
            label="Run at Windows Startup",
            variable=self.startup_var,
            command=self.toggle_startup,
        )
        self.menu.add_command(
            label="Reset Settings to Default...",
            command=self.reset_settings_to_default,
        )

        # --- Exit (Bold) with Red Circle Cross Icon on Right ---
        self.menu.add_separator()
        bold_menu_font = ("Segoe UI", 9, "bold")
        self.menu.add_command(
            label="Exit   ",
            font=bold_menu_font,
            image=self.exit_icon,
            compound="right",
            command=self.exit_app,
        )

        # --- Footer: GitHub Link & Version 2 ---
        self.menu.add_separator()
        link_font = ("Segoe UI", 9, "underline")
        self.menu.add_command(
            label="GitHub: github.com/ManishKashyapIndia",
            font=link_font,
            command=self.open_github_page,
        )
        self._github_idx = self.menu.index("end")
        footer_font = ("Courier New", 9, "bold")
        self.menu.add_command(
            label="Version 2",
            font=footer_font,
            command=lambda: None,
        )
        self._version_idx = self.menu.index("end")

        # --- Last Option: Uninstall App (Bold, default color, separated after copyright) ---
        self.menu.add_separator()
        self.menu.add_command(
            label="Uninstall App...",
            font=bold_menu_font,
            command=self.uninstall_app,
        )

        # Apply initial Light/Dark context menu styling
        self.apply_menu_theme()

        self.display.bind("<Button-3>", self.show_menu)

        # Hide from Windows Taskbar before showing widget
        self.hide_from_taskbar()
        self.root.deiconify()
        self.apply_window_layer()
        self.show_next_quote(animate=False)
        self.start_schedule_heartbeat()
        self.root.mainloop()

    # -------------------- Context Menu Theme (Dark / Light / System) --------------------

    def _make_menu(self, parent):
        m = tk.Menu(parent, tearoff=0)
        self._all_menus.append(m)
        return m

    def get_effective_dark_mode(self):
        """Return True if Dark Mode is currently active (explicit or via Windows System setting)."""
        mode = self.config.get("theme_mode", "light")
        if mode == "dark":
            return True
        if mode == "light":
            return False
        return is_windows_dark_mode()

    def get_active_presets_pool(self):
        """Return Dark Mode presets or Light Mode presets based on the active appearance mode."""
        return DARK_PRESETS_LIST if self.get_effective_dark_mode() else LIGHT_PRESETS_LIST

    def apply_menu_theme(self):
        """Style the right-click context menu and all submenus for Dark or Light mode."""
        dark = self.get_effective_dark_mode()
        if dark:
            bg = "#1F2937"
            fg = "#F9FAFB"
            active_bg = "#374151"
            active_fg = "#FFFFFF"
            disabled_fg = "#9CA3AF"
            select_col = "#38BDF8"
            link_fg = "#60A5FA"
            link_active = "#93C5FD"
            footer_fg = "#93C5FD"
        else:
            bg = "#FFFFFF"
            fg = "#111827"
            active_bg = "#E5E7EB"
            active_fg = "#000000"
            disabled_fg = "#6B7280"
            select_col = "#2563EB"
            link_fg = "#0066CC"
            link_active = "#004499"
            footer_fg = "#1E40AF"

        for m in self._all_menus:
            try:
                m.configure(
                    bg=bg,
                    fg=fg,
                    activebackground=active_bg,
                    activeforeground=active_fg,
                    disabledforeground=disabled_fg,
                    selectcolor=select_col,
                )
            except Exception:
                pass

        try:
            self.menu.entryconfig(self._github_idx, foreground=link_fg, activeforeground=link_active)
            self.menu.entryconfig(self._version_idx, foreground=footer_fg, activeforeground=footer_fg)
        except Exception:
            pass

    def set_theme_mode(self):
        self.config["theme_mode"] = self.theme_mode_var.get()
        self.save_config()
        self.apply_menu_theme()
        # Immediately apply a matching preset if randomize_style is enabled
        if self.config.get("randomize_style", True):
            chosen_preset = random.choice(self.get_active_presets_pool())
            self.apply_preset(chosen_preset, save=True, render=True)
        mode_lbl = {"system": "System", "dark": "Dark", "light": "Light"}.get(
            self.config["theme_mode"], "System"
        )
        self.show_toast(f"Appearance set to {mode_lbl} Mode")

    def reset_settings_to_default(self):
        """Reset all visual and behavioral settings to defaults while preserving installation & quotes."""
        confirm = messagebox.askyesno(
            "Reset Settings to Default",
            "Are you sure you want to reset all widget settings to their default values?\n\n"
            "Your quotes.xlsx file and installation folder will remain untouched.",
            parent=self.root,
        )
        if not confirm:
            return

        # Preserve installation and file paths
        preserved = {
            "installed": self.config.get("installed", True),
            "disclaimer_accepted": self.config.get("disclaimer_accepted", True),
            "categories_initialized": self.config.get("categories_initialized", True),
            "install_dir": self.install_dir,
            "excel_file": self.config.get("excel_file", DEFAULT_EXCEL_FILE),
        }

        # Build fresh defaults
        fresh = self.load_config("__nonexistent_reset_path__")
        fresh.update(preserved)
        fresh["selected_categories"] = list(self.all_categories)
        self.config = fresh

        # Sync all Tk menu variables
        self.order_var.set(self.config["order_mode"])
        self.window_layer_var.set(self.config["window_layer"])
        self.text_align_var.set(self.config["text_align"])
        self.theme_mode_var.set(self.config["theme_mode"])
        self.startup_var.set(self.config["run_at_startup"])
        self.randomize_style_var.set(self.config["randomize_style"])
        self.anim_enabled_var.set(self.config["anim_enabled"])
        self.anim_mode_var.set(self.config["anim_mode"])
        self.shadow_var.set(self.config["text_shadow"])
        self.border_var.set(self.config["border_enabled"])

        for cat in self.all_categories:
            self.category_vars[cat].set(True)

        # Trim undo stack if needed & unpause timer
        if len(self.undo_stack) > DEFAULT_UNDO_LIMIT:
            self.undo_stack = self.undo_stack[-DEFAULT_UNDO_LIMIT:]
        self.is_paused = False
        self._paused_remaining_secs = None

        set_windows_startup(self.config["run_at_startup"])
        self.save_config()
        self.apply_menu_theme()
        self.apply_window_layer()
        self.root.geometry(f"+{self.config['x']}+{self.config['y']}")
        self.reset_schedule_timer()
        self.render_widget()
        self.show_toast("All settings reset to default")

    # -------------------- Undo Limit Dialog (1 to 20, default 5) --------------------

    def open_undo_limit_dialog(self):
        dialog = tk.Toplevel(self.root)
        dialog.title("Set Undo History Limit")
        dialog.resizable(False, False)
        dialog.attributes("-topmost", True)

        current_limit = int(self.config.get("undo_limit", DEFAULT_UNDO_LIMIT))
        limit_var = tk.IntVar(value=current_limit)

        ttk.Label(
            dialog,
            text=f"Set maximum number of Add/Delete actions remembered by Undo\n"
                 f"(Allowed range: {MIN_UNDO_LIMIT} to {MAX_UNDO_LIMIT}, Default: {DEFAULT_UNDO_LIMIT}).",
            justify="left",
        ).grid(row=0, column=0, columnspan=2, padx=14, pady=(12, 8), sticky="w")

        ttk.Label(dialog, text="Undo Limit:").grid(row=1, column=0, padx=14, pady=6, sticky="w")
        ttk.Spinbox(
            dialog, from_=MIN_UNDO_LIMIT, to=MAX_UNDO_LIMIT, textvariable=limit_var, width=8
        ).grid(row=1, column=1, padx=14, pady=6, sticky="w")

        def apply_limit():
            try:
                val = int(limit_var.get())
            except tk.TclError:
                val = DEFAULT_UNDO_LIMIT
            val = max(MIN_UNDO_LIMIT, min(MAX_UNDO_LIMIT, val))
            self.config["undo_limit"] = val
            # Trim existing undo stack if it exceeds the new limit
            if len(self.undo_stack) > val:
                self.undo_stack = self.undo_stack[-val:]
            self.save_config()
            self.render_widget()
            dialog.destroy()
            self.show_toast(f"Undo limit set to {val}")

        btns = ttk.Frame(dialog)
        btns.grid(row=2, column=0, columnspan=2, pady=(8, 12))
        ttk.Button(btns, text="Apply", command=apply_limit).pack(side="left", padx=6)
        ttk.Button(btns, text="Cancel", command=dialog.destroy).pack(side="left", padx=6)

        dialog.update_idletasks()
        dw, dh = dialog.winfo_width(), dialog.winfo_height()
        sx = (dialog.winfo_screenwidth() - dw) // 2
        sy = (dialog.winfo_screenheight() - dh) // 2
        dialog.geometry(f"+{sx}+{sy}")

    # -------------------- 3-Second On-Screen Toast Notification --------------------

    def show_toast(self, message, duration_ms=3000):
        self.hide_toast()
        try:
            tw = tk.Toplevel(self.root)
            tw.overrideredirect(True)
            tw.attributes("-topmost", True)
            dark = self.get_effective_dark_mode()
            bg_c = "#0F172A" if dark else "#F8FAFC"
            fg_c = "#F8FAFC" if dark else "#0F172A"
            tw.configure(bg=bg_c)

            lbl = tk.Label(
                tw,
                text=message,
                bg=bg_c,
                fg=fg_c,
                font=("Segoe UI", 9, "bold"),
                padx=12,
                pady=6,
                relief="solid",
                bd=1,
            )
            lbl.pack()

            tw.update_idletasks()
            tw_w = tw.winfo_width()
            tw_h = tw.winfo_height()

            wx = self.root.winfo_x()
            wy = self.root.winfo_y()
            ww = self.root.winfo_width()
            wh = self.root.winfo_height()

            tx = wx + max(0, (ww - tw_w) // 2)
            ty = wy + wh + 6
            sh = self.root.winfo_screenheight()
            if ty + tw_h > sh - 40:
                ty = max(10, wy - tw_h - 6)

            tw.geometry(f"+{tx}+{ty}")
            self._toast_win = tw
            self._toast_after_id = self.root.after(duration_ms, self.hide_toast)
        except Exception:
            pass

    def hide_toast(self):
        if self._toast_after_id is not None:
            try:
                self.root.after_cancel(self._toast_after_id)
            except Exception:
                pass
            self._toast_after_id = None
        if self._toast_win is not None:
            try:
                self._toast_win.destroy()
            except Exception:
                pass
            self._toast_win = None

    # -------------------- Top Circular Buttons (Auto-Wrap & Hover-Reveal) --------------------

    def compute_top_controls_layout(self, widget_w):
        """
        Compute multi-row layout for the 6 top-left buttons and the top-right close button.
        Returns (button_rects, top_bar_h).
        """
        r = 10
        gap = 6
        step = r * 2 + gap
        row_step = r * 2 + 4
        margin_x = 14
        top_y = 14

        left_buttons = [
            ("prev", "Previous Quote"),
            ("next", "Next Quote"),
            ("pause", "Resume Timer" if self.is_paused else "Pause Timer"),
            ("add", "Add New Quote"),
            ("delete", "Delete Current Quote"),
            ("undo", "Undo (Add/Delete)"),
        ]

        # Top-right Close App button (red circle with cross)
        close_cx = max(margin_x + r, widget_w - margin_x - r)
        close_cy = top_y

        # Right boundary on row 0 leaves room for the close button
        row0_max_x = close_cx - r - gap
        other_row_max_x = widget_w - margin_x

        rects = []
        row = 0
        col = 0
        for name, tip in left_buttons:
            cx = margin_x + r + col * step
            max_x = row0_max_x if row == 0 else other_row_max_x
            if col > 0 and (cx + r) > max_x:
                row += 1
                col = 0
                cx = margin_x + r
            cy = top_y + row * row_step
            rects.append((name, tip, cx, cy, r))
            col += 1

        # Add the top-right close button
        rects.append(("close", "Close App", close_cx, close_cy, r))

        num_rows = row + 1
        top_bar_h = top_y + (num_rows - 1) * row_step + r + 2
        return rects, top_bar_h

    def draw_top_controls(self, canvas_img, rects):
        """Draw all circular control buttons when hovered."""
        draw = ImageDraw.Draw(canvas_img)

        for name, _tip, cx, cy, r in rects:
            if name == "close":
                fill_col = (220, 38, 38, 240)
                outline_col = (254, 202, 202, 245)
                icon_col = (255, 255, 255, 255)
            elif name == "delete":
                fill_col = (185, 28, 28, 230)
                outline_col = (254, 202, 202, 240)
                icon_col = (255, 255, 255, 255)
            elif name == "pause" and self.is_paused:
                fill_col = (22, 163, 74, 235)
                outline_col = (187, 247, 208, 245)
                icon_col = (255, 255, 255, 255)
            elif name == "undo" and not self.undo_stack:
                fill_col = (71, 85, 105, 125)
                outline_col = (148, 163, 184, 125)
                icon_col = (203, 213, 225, 140)
            elif name == "prev" and self.history_pos <= 0:
                fill_col = (71, 85, 105, 125)
                outline_col = (148, 163, 184, 125)
                icon_col = (203, 213, 225, 140)
            elif name == "add":
                fill_col = (37, 99, 235, 230)
                outline_col = (191, 219, 254, 240)
                icon_col = (255, 255, 255, 255)
            else:
                fill_col = (15, 23, 42, 225)
                outline_col = (226, 232, 240, 230)
                icon_col = (255, 255, 255, 255)

            draw.ellipse([cx - r, cy - r, cx + r, cy + r], fill=fill_col, outline=outline_col, width=1)

            if name == "prev":
                draw.polygon([(cx + 3, cy - 4), (cx + 3, cy + 4), (cx - 4, cy)], fill=icon_col)
            elif name == "next":
                draw.polygon([(cx - 3, cy - 4), (cx - 3, cy + 4), (cx + 4, cy)], fill=icon_col)
            elif name == "pause":
                if self.is_paused:
                    draw.polygon([(cx - 3, cy - 4), (cx - 3, cy + 4), (cx + 4, cy)], fill=icon_col)
                else:
                    draw.rectangle([cx - 4, cy - 4, cx - 1, cy + 4], fill=icon_col)
                    draw.rectangle([cx + 1, cy - 4, cx + 4, cy + 4], fill=icon_col)
            elif name == "add":
                draw.line([(cx - 4, cy), (cx + 4, cy)], fill=icon_col, width=2)
                draw.line([(cx, cy - 4), (cx, cy + 4)], fill=icon_col, width=2)
            elif name == "delete":
                # Minus sign for deleting current quote
                draw.line([(cx - 4, cy), (cx + 4, cy)], fill=icon_col, width=2)
            elif name == "undo":
                draw.arc([cx - 4, cy - 4, cx + 5, cy + 5], start=210, end=90, fill=icon_col, width=2)
                draw.polygon([(cx - 5, cy - 2), (cx - 1, cy - 5), (cx - 1, cy + 1)], fill=icon_col)
            elif name == "close":
                # Cross sign for closing the app
                draw.line([(cx - 4, cy - 4), (cx + 4, cy + 4)], fill=icon_col, width=2)
                draw.line([(cx + 4, cy - 4), (cx - 4, cy + 4)], fill=icon_col, width=2)

        return canvas_img

    def _cancel_hide_controls(self):
        if self._hide_controls_after_id is not None:
            try:
                self.root.after_cancel(self._hide_controls_after_id)
            except Exception:
                pass
            self._hide_controls_after_id = None

    def _schedule_hide_controls(self):
        self._cancel_hide_controls()
        if self._controls_visible:
            self._hide_controls_after_id = self.root.after(260, self._do_hide_controls)

    def _do_hide_controls(self):
        self._hide_controls_after_id = None
        if self._controls_visible:
            self._controls_visible = False
            self.hide_tooltip()
            self.render_widget()

    def show_tooltip(self, text, x_root, y_root):
        self.hide_tooltip()
        tw = tk.Toplevel(self.root)
        tw.overrideredirect(True)
        tw.attributes("-topmost", True)
        tw.geometry(f"+{x_root + 10}+{y_root + 16}")
        lbl = tk.Label(
            tw,
            text=text,
            bg="#1E293B",
            fg="#F8FAFC",
            font=("Segoe UI", 8),
            padx=6,
            pady=2,
            relief="solid",
            bd=1,
        )
        lbl.pack()
        self._tooltip_win = tw

    def hide_tooltip(self):
        if self._tooltip_win is not None:
            try:
                self._tooltip_win.destroy()
            except Exception:
                pass
            self._tooltip_win = None
        self._hovered_btn = None

    def handle_control_button(self, btn_name):
        self.hide_tooltip()
        if btn_name == "prev":
            self.show_prev_quote()
        elif btn_name == "next":
            self.show_next_quote(animate=True, show_toast=True)
        elif btn_name == "pause":
            self.toggle_pause()
        elif btn_name == "add":
            self.open_add_quote_dialog()
        elif btn_name == "delete":
            self.delete_current_quote()
        elif btn_name == "undo":
            self.undo_last_action()
        elif btn_name == "close":
            self.exit_app()

    # -------------------- Quote Actions: Prev, Pause, Delete, Undo, Add --------------------

    def _push_quote_history(self, quote_text):
        self.quote_history = self.quote_history[: self.history_pos + 1]
        self.quote_history.append(quote_text)
        if len(self.quote_history) > MAX_PREV_HISTORY + 1:
            self.quote_history = self.quote_history[-(MAX_PREV_HISTORY + 1) :]
        self.history_pos = len(self.quote_history) - 1

    def _push_undo_action(self, action_type, quote_text, category_name):
        limit = max(
            MIN_UNDO_LIMIT,
            min(MAX_UNDO_LIMIT, int(self.config.get("undo_limit", DEFAULT_UNDO_LIMIT))),
        )
        self.undo_stack.append((action_type, quote_text, category_name))
        if len(self.undo_stack) > limit:
            self.undo_stack = self.undo_stack[-limit:]

    def show_prev_quote(self):
        if self.history_pos <= 0:
            self.show_toast("No earlier quote remembered")
            return
        self.cancel_animation()
        self.history_pos -= 1
        self.current_quote = self.quote_history[self.history_pos]
        self._render_seed = random.randint(1, 999999)

        if self.config.get("randomize_style", True):
            chosen_preset = random.choice(self.get_active_presets_pool())
            self.apply_preset(chosen_preset, save=False, render=False)

        self.render_widget()
        if self.config.get("anim_enabled", True):
            self.play_change_animation()
        self.show_toast("Previous quote displayed")

    def toggle_pause(self):
        if not self.is_paused:
            self.is_paused = True
            if self._next_change_ts is not None:
                self._paused_remaining_secs = max(1, int(self._next_change_ts - time.time()))
            else:
                self._paused_remaining_secs = None
            self.render_widget()
            self.show_toast("Quote timer paused")
        else:
            self.is_paused = False
            if self._paused_remaining_secs is not None:
                self._next_change_ts = time.time() + self._paused_remaining_secs
            else:
                self.reset_schedule_timer()
            self._paused_remaining_secs = None
            self.render_widget()
            self.show_toast("Quote timer resumed")

    def delete_current_quote(self):
        if not self.current_quote:
            return
        excel_path = self.config.get("excel_file", DEFAULT_EXCEL_FILE)
        if not os.path.isfile(excel_path):
            return

        target_quote = self.current_quote
        deleted_entry = None
        try:
            wb = openpyxl.load_workbook(excel_path)
            ws = wb.active
            row_to_delete = None
            for row_idx in range(2, ws.max_row + 1):
                val = ws.cell(row=row_idx, column=1).value
                if val is not None and str(val).strip() == target_quote:
                    cat_val = ws.cell(row=row_idx, column=2).value
                    cat_str = str(cat_val).strip() if cat_val and str(cat_val).strip() else "Uncategorized"
                    deleted_entry = (target_quote, cat_str)
                    row_to_delete = row_idx
                    break
            if row_to_delete is not None:
                ws.delete_rows(row_to_delete, 1)
                wb.save(excel_path)
            wb.close()
        except Exception as exc:
            show_windows_popup(
                "Delete Failed",
                f"Could not delete quote from quotes.xlsx.\nMake sure the file is not open in Excel.\n\n{exc}",
                is_error=True,
            )
            return

        if deleted_entry:
            self._push_undo_action("delete", deleted_entry[0], deleted_entry[1])

            if 0 <= self.history_pos < len(self.quote_history):
                if self.quote_history[self.history_pos] == target_quote:
                    self.quote_history.pop(self.history_pos)
                    self.history_pos = max(-1, self.history_pos - 1)
            self.show_next_quote(animate=True, show_toast=False)
            self.show_toast("Quote deleted successfully")

    def undo_last_action(self):
        if not self.undo_stack:
            self.show_toast("Nothing to undo")
            return

        action_type, quote_text, category_name = self.undo_stack.pop()
        excel_path = self.config.get("excel_file", DEFAULT_EXCEL_FILE)

        if action_type == "delete":
            try:
                wb = openpyxl.load_workbook(excel_path)
                ws = wb.active
                next_row = ws.max_row + 1
                ws.cell(row=next_row, column=1, value=quote_text).alignment = Alignment(
                    wrap_text=True, vertical="center"
                )
                ws.cell(row=next_row, column=2, value=category_name).alignment = Alignment(
                    vertical="center"
                )
                wb.save(excel_path)
                wb.close()
            except Exception as exc:
                self.undo_stack.append((action_type, quote_text, category_name))
                show_windows_popup(
                    "Undo Failed",
                    f"Could not restore quote to quotes.xlsx.\nMake sure the file is not open in Excel.\n\n{exc}",
                    is_error=True,
                )
                return

            self.quote_entries = self.load_quote_entries()
            self.all_categories = sorted({cat for _, cat in self.quote_entries})
            if category_name not in self.category_vars:
                self.category_vars[category_name] = tk.BooleanVar(value=True)
            else:
                self.category_vars[category_name].set(True)
            self.rebuild_category_menu()
            self.quotes = self.get_filtered_quotes()

            self.current_quote = quote_text
            self._push_quote_history(quote_text)
            self.render_widget()
            if self.config.get("anim_enabled", True):
                self.play_change_animation()
            self.show_toast("Undo successful — deleted quote restored")

        elif action_type == "add":
            try:
                wb = openpyxl.load_workbook(excel_path)
                ws = wb.active
                row_to_remove = None
                for row_idx in range(ws.max_row, 1, -1):
                    val = ws.cell(row=row_idx, column=1).value
                    if val is not None and str(val).strip() == quote_text:
                        row_to_remove = row_idx
                        break
                if row_to_remove is not None:
                    ws.delete_rows(row_to_remove, 1)
                    wb.save(excel_path)
                wb.close()
            except Exception as exc:
                self.undo_stack.append((action_type, quote_text, category_name))
                show_windows_popup(
                    "Undo Failed",
                    f"Could not remove added quote from quotes.xlsx.\nMake sure the file is not open in Excel.\n\n{exc}",
                    is_error=True,
                )
                return

            self.quote_entries = self.load_quote_entries()
            self.all_categories = sorted({cat for _, cat in self.quote_entries})
            self.rebuild_category_menu()
            self.quotes = self.get_filtered_quotes()

            if 0 <= self.history_pos < len(self.quote_history) and self.quote_history[self.history_pos] == quote_text:
                self.quote_history.pop(self.history_pos)
                self.history_pos = max(-1, self.history_pos - 1)

            if 0 <= self.history_pos < len(self.quote_history):
                self.current_quote = self.quote_history[self.history_pos]
                self.render_widget()
            else:
                self.show_next_quote(animate=True, show_toast=False)
            self.show_toast("Undo successful — added quote removed")

    def open_add_quote_dialog(self):
        dialog = tk.Toplevel(self.root)
        dialog.title("Add New Quote")
        dialog.resizable(False, False)
        dialog.attributes("-topmost", True)

        icon_path = get_bundled_resource(ICON_FILENAME)
        if icon_path:
            try:
                dialog.iconbitmap(icon_path)
            except Exception:
                pass

        ttk.Label(dialog, text="Quote: (Any Language)", font=("Segoe UI", 9, "bold")).grid(
            row=0, column=0, padx=14, pady=(12, 4), sticky="w"
        )
        quote_text_widget = tk.Text(dialog, width=48, height=4, wrap="word", font=("Segoe UI", 10))
        quote_text_widget.grid(row=1, column=0, columnspan=2, padx=14, pady=(0, 10))

        ttk.Label(dialog, text="Category:", font=("Segoe UI", 9, "bold")).grid(
            row=2, column=0, padx=14, pady=(2, 4), sticky="w"
        )

        CREATE_NEW_OPTION = "+ Create New Category..."
        cat_options = sorted(set(self.all_categories) | {"Uncategorized"})
        cat_options.append(CREATE_NEW_OPTION)

        cat_var = tk.StringVar(value="")
        cb_cat = ttk.Combobox(dialog, textvariable=cat_var, values=cat_options, width=32, state="readonly")
        cb_cat.grid(row=2, column=1, padx=(0, 14), pady=(2, 4), sticky="ew")

        new_cat_lbl = ttk.Label(dialog, text="New Category Name:")
        new_cat_var = tk.StringVar(value="")
        new_cat_entry = ttk.Entry(dialog, textvariable=new_cat_var, width=34)

        def on_cat_change(*_args):
            if cat_var.get() == CREATE_NEW_OPTION:
                new_cat_lbl.grid(row=3, column=0, padx=14, pady=6, sticky="w")
                new_cat_entry.grid(row=3, column=1, padx=(0, 14), pady=6, sticky="ew")
                new_cat_entry.focus_set()
            else:
                new_cat_lbl.grid_forget()
                new_cat_entry.grid_forget()

        cb_cat.bind("<<ComboboxSelected>>", on_cat_change)

        def save_new_quote():
            q_str = quote_text_widget.get("1.0", "end").strip()
            if not q_str:
                messagebox.showwarning("Empty Quote", "Please enter a quote before saving.", parent=dialog)
                return

            selected = cat_var.get().strip()
            if selected == CREATE_NEW_OPTION:
                cat_str = new_cat_var.get().strip() or "Uncategorized"
            elif not selected:
                cat_str = "Uncategorized"
            else:
                cat_str = selected

            excel_path = self.config.get("excel_file", DEFAULT_EXCEL_FILE)
            try:
                wb = openpyxl.load_workbook(excel_path)
                ws = wb.active
                next_row = ws.max_row + 1
                ws.cell(row=next_row, column=1, value=q_str).alignment = Alignment(
                    wrap_text=True, vertical="center"
                )
                ws.cell(row=next_row, column=2, value=cat_str).alignment = Alignment(
                    vertical="center"
                )
                wb.save(excel_path)
                wb.close()
            except Exception as exc:
                messagebox.showerror(
                    "Save Failed",
                    f"Could not write to quotes.xlsx.\nMake sure it is not open in Excel.\n\n{exc}",
                    parent=dialog,
                )
                return

            self._push_undo_action("add", q_str, cat_str)

            self.quote_entries = self.load_quote_entries()
            self.all_categories = sorted({cat for _, cat in self.quote_entries})
            if cat_str not in self.category_vars:
                self.category_vars[cat_str] = tk.BooleanVar(value=True)
            else:
                self.category_vars[cat_str].set(True)
            self.config["selected_categories"] = [
                c for c, v in self.category_vars.items() if v.get()
            ]
            self.rebuild_category_menu()
            self.save_config()

            self.quotes = self.get_filtered_quotes()
            self.current_quote = q_str
            self._push_quote_history(q_str)
            self.render_widget()
            dialog.destroy()
            self.show_toast("Quote added successfully")

        btn_frame = ttk.Frame(dialog)
        btn_frame.grid(row=4, column=0, columnspan=2, pady=(10, 14))
        ttk.Button(btn_frame, text="OK", command=save_new_quote, width=12).pack(side="left", padx=8)
        ttk.Button(btn_frame, text="Cancel", command=dialog.destroy, width=12).pack(side="left", padx=8)

        dialog.update_idletasks()
        dw, dh = dialog.winfo_width(), dialog.winfo_height()
        sx = (dialog.winfo_screenwidth() - dw) // 2
        sy = (dialog.winfo_screenheight() - dh) // 2
        dialog.geometry(f"+{sx}+{sy}")
        quote_text_widget.focus_set()

    # -------------------- Edit Excel File --------------------

    def edit_excel_file(self):
        excel_path = self.config.get("excel_file", DEFAULT_EXCEL_FILE)
        if not os.path.isfile(excel_path):
            os.makedirs(os.path.dirname(excel_path) or self.install_dir, exist_ok=True)
            create_default_excel_file(excel_path)
        try:
            os.startfile(excel_path)
            self.show_toast("Opening quotes.xlsx for editing...")
        except Exception as exc:
            show_windows_popup(
                "Unable to Open Excel File",
                f"Could not open:\n{excel_path}\n\n{exc}",
                is_error=True,
            )

    # -------------------- Uninstall App --------------------

    def uninstall_app(self):
        confirm = messagebox.askyesno(
            f"Uninstall {APP_DISPLAY_NAME}",
            f"Are you sure you want to uninstall {APP_DISPLAY_NAME}?\n\n"
            "• All application files, settings, Desktop & Start Menu shortcuts, and startup entries will be removed.\n"
            "• Your quotes.xlsx file will be PRESERVED so you do not lose your custom quotes.",
            parent=self.root,
        )
        if not confirm:
            return

        set_windows_startup(False)
        clear_install_state()

        for sc_path in (get_desktop_shortcut_path(), get_start_menu_shortcut_path()):
            if os.path.isfile(sc_path):
                try:
                    os.remove(sc_path)
                except Exception:
                    pass

        excel_path = os.path.abspath(self.config.get("excel_file", DEFAULT_EXCEL_FILE))
        install_dir = os.path.abspath(self.install_dir)
        current_exe = os.path.abspath(sys.executable) if getattr(sys, "frozen", False) else None
        needs_delayed_exe_delete = False

        if os.path.isdir(install_dir):
            for entry in os.listdir(install_dir):
                full_path = os.path.abspath(os.path.join(install_dir, entry))
                if os.path.normcase(full_path) == os.path.normcase(excel_path):
                    continue
                if entry.lower() == EXCEL_FILENAME.lower():
                    continue
                if current_exe and os.path.normcase(full_path) == os.path.normcase(current_exe):
                    needs_delayed_exe_delete = True
                    continue
                try:
                    if os.path.isdir(full_path):
                        shutil.rmtree(full_path, ignore_errors=True)
                    else:
                        os.remove(full_path)
                except Exception:
                    pass

        self.show_uninstall_preserved_notice(excel_path)

        if needs_delayed_exe_delete and current_exe:
            try:
                subprocess.Popen(
                    ["cmd", "/c", f'timeout /t 2 /nobreak >nul & del /f /q "{current_exe}"'],
                    creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
                )
            except Exception:
                pass

        self.exit_app(save_cfg=False)

    def show_uninstall_preserved_notice(self, excel_path):
        dialog = tk.Toplevel(self.root)
        dialog.title(f"{APP_DISPLAY_NAME} - Uninstalled")
        dialog.resizable(False, False)
        dialog.attributes("-topmost", True)

        ttk.Label(
            dialog,
            text=f"{APP_DISPLAY_NAME} has been uninstalled.",
            font=("Segoe UI", 11, "bold"),
        ).pack(padx=18, pady=(14, 6), anchor="w")

        msg = (
            "Your quotes file (quotes.xlsx) has been preserved so you can keep any\n"
            "quotes you added or modified while using the app for future use.\n\n"
            "Click the path below to open its folder and back up or view your file:"
        )
        ttk.Label(dialog, text=msg, justify="left", font=("Segoe UI", 9)).pack(
            padx=18, pady=(0, 8), anchor="w"
        )

        def open_preserved_location(_event=None):
            try:
                if os.path.isfile(excel_path):
                    subprocess.Popen(["explorer", "/select,", excel_path])
                else:
                    folder = os.path.dirname(excel_path)
                    if os.path.isdir(folder):
                        os.startfile(folder)
            except Exception:
                pass

        link_lbl = tk.Label(
            dialog,
            text=excel_path,
            fg="#0066CC",
            cursor="hand2",
            font=("Segoe UI", 9, "underline"),
            wraplength=460,
            justify="left",
        )
        link_lbl.pack(padx=18, pady=(0, 14), anchor="w")
        link_lbl.bind("<Button-1>", open_preserved_location)

        ttk.Button(dialog, text="Close", command=dialog.destroy, width=14).pack(pady=(0, 14))

        dialog.update_idletasks()
        dw, dh = dialog.winfo_width(), dialog.winfo_height()
        sx = (dialog.winfo_screenwidth() - dw) // 2
        sy = (dialog.winfo_screenheight() - dh) // 2
        dialog.geometry(f"+{sx}+{sy}")

        dialog.grab_set()
        self.root.wait_window(dialog)

    # -------------------- Screen 2: Installation Folder Dialog --------------------

    def prompt_install_folder(self):
        result = {"dir": None}
        dialog = tk.Toplevel(self.root)
        dialog.title(f"{APP_DISPLAY_NAME} - Select Installation Folder")
        dialog.resizable(False, False)
        dialog.attributes("-topmost", True)

        icon_path = get_bundled_resource(ICON_FILENAME)
        if icon_path:
            try:
                dialog.iconbitmap(icon_path)
            except Exception:
                pass

        path_var = tk.StringVar(value=DEFAULT_INSTALL_DIR)

        info_text = (
            f"Choose the installation folder for {APP_DISPLAY_NAME}.\n\n"
            "All application files—including widget_config.json and quotes.xlsx—will be\n"
            "stored here. If a valid quotes.xlsx already exists in this folder, it will\n"
            "be preserved without overwriting your changes."
        )
        ttk.Label(dialog, text=info_text, justify="left").grid(
            row=0, column=0, columnspan=3, padx=16, pady=(14, 10), sticky="w"
        )

        ttk.Label(dialog, text="Install Folder:").grid(
            row=1, column=0, padx=(16, 6), pady=6, sticky="w"
        )
        entry = ttk.Entry(dialog, textvariable=path_var, width=48)
        entry.grid(row=1, column=1, padx=4, pady=6, sticky="ew")

        def browse_folder():
            current = path_var.get().strip()
            init_dir = current if os.path.isdir(current) else "C:\\"
            chosen = filedialog.askdirectory(
                parent=dialog,
                title=f"Select Installation Folder for {APP_DISPLAY_NAME}",
                initialdir=init_dir,
            )
            if chosen:
                path_var.set(os.path.normpath(chosen))

        ttk.Button(dialog, text="Browse...", command=browse_folder).grid(
            row=1, column=2, padx=(6, 16), pady=6
        )

        def confirm():
            target = path_var.get().strip().strip('"')
            if not target:
                messagebox.showerror(
                    "Invalid Folder",
                    "Please enter or browse to a valid folder path.",
                    parent=dialog,
                )
                return
            try:
                os.makedirs(target, exist_ok=True)
                test_file = os.path.join(target, ".write_test_tmp")
                with open(test_file, "w", encoding="utf-8") as f:
                    f.write("ok")
                os.remove(test_file)
            except PermissionError:
                messagebox.showerror(
                    "Permission Denied",
                    f"Cannot write to:\n{target}\n\n"
                    "Please run the application as Administrator, or click 'Browse...' "
                    "to select a writable folder.",
                    parent=dialog,
                )
                return
            except Exception as exc:
                messagebox.showerror(
                    "Folder Error",
                    f"Could not access or create folder:\n{target}\n\n{exc}",
                    parent=dialog,
                )
                return

            result["dir"] = os.path.normpath(target)
            dialog.destroy()

        def cancel():
            result["dir"] = None
            dialog.destroy()

        dialog.protocol("WM_DELETE_WINDOW", cancel)

        btn_frame = ttk.Frame(dialog)
        btn_frame.grid(row=2, column=0, columnspan=3, pady=(10, 14))
        ttk.Button(btn_frame, text="Install & Continue", command=confirm).pack(side="left", padx=8)
        ttk.Button(btn_frame, text="Exit", command=cancel).pack(side="left", padx=8)

        dialog.update_idletasks()
        dw, dh = dialog.winfo_width(), dialog.winfo_height()
        sx = (dialog.winfo_screenwidth() - dw) // 2
        sy = (dialog.winfo_screenheight() - dh) // 2
        dialog.geometry(f"+{sx}+{sy}")

        dialog.grab_set()
        self.root.wait_window(dialog)
        return result["dir"]

    # -------------------- Category Selection & Menu --------------------

    def prompt_category_selection(self, categories):
        result = {"selected": None}
        dialog = tk.Toplevel(self.root)
        dialog.title("Select Quote Categories")
        dialog.resizable(False, False)
        dialog.attributes("-topmost", True)

        icon_path = get_bundled_resource(ICON_FILENAME)
        if icon_path:
            try:
                dialog.iconbitmap(icon_path)
            except Exception:
                pass

        ttk.Label(
            dialog,
            text="Select the quote categories you want to display:",
            font=("Segoe UI", 10, "bold"),
        ).pack(padx=16, pady=(14, 8), anchor="w")

        chk_frame = ttk.Frame(dialog)
        chk_frame.pack(padx=20, pady=(0, 10), fill="both", expand=True)

        vars_map = {}
        for cat in categories:
            var = tk.BooleanVar(value=True)
            vars_map[cat] = var
            ttk.Checkbutton(chk_frame, text=cat, variable=var).pack(anchor="w", pady=3)

        def select_all():
            for v in vars_map.values():
                v.set(True)

        def confirm():
            chosen = [c for c, v in vars_map.items() if v.get()]
            if not chosen:
                messagebox.showwarning(
                    "No Category Selected",
                    "Please select at least one category to continue.",
                    parent=dialog,
                )
                return
            result["selected"] = chosen
            dialog.destroy()

        def cancel():
            result["selected"] = None
            dialog.destroy()

        dialog.protocol("WM_DELETE_WINDOW", cancel)

        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(padx=16, pady=(4, 14), fill="x")
        ttk.Button(btn_frame, text="Select All", command=select_all).pack(side="left")
        ttk.Button(btn_frame, text="Start Widget", command=confirm).pack(side="right", padx=(6, 0))
        ttk.Button(btn_frame, text="Exit", command=cancel).pack(side="right")

        dialog.update_idletasks()
        dw, dh = dialog.winfo_width(), dialog.winfo_height()
        sx = (dialog.winfo_screenwidth() - dw) // 2
        sy = (dialog.winfo_screenheight() - dh) // 2
        dialog.geometry(f"+{sx}+{sy}")

        dialog.grab_set()
        self.root.wait_window(dialog)
        return result["selected"]

    def rebuild_category_menu(self):
        self.category_menu.delete(0, "end")
        for cat in self.all_categories:
            self.category_menu.add_checkbutton(
                label=cat,
                variable=self.category_vars[cat],
                command=lambda c=cat: self.on_category_toggled(c),
            )

    def on_category_toggled(self, toggled_cat):
        active = [c for c, v in self.category_vars.items() if v.get()]
        if not active:
            self.category_vars[toggled_cat].set(True)
            show_windows_popup(
                "Category Required",
                "At least one quote category must remain selected.",
                is_error=False,
            )
            return
        self.config["selected_categories"] = active
        self.save_config()
        self.show_next_quote(animate=True)

    # -------------------- Taskbar Hiding & Window Layer --------------------

    def hide_from_taskbar(self):
        try:
            self.root.update_idletasks()
            hwnd = ctypes.windll.user32.GetParent(self.root.winfo_id())
            if not hwnd:
                hwnd = self.root.winfo_id()
            style = ctypes.windll.user32.GetWindowLongW(hwnd, GWL_EXSTYLE)
            style = (style & ~WS_EX_APPWINDOW) | WS_EX_TOOLWINDOW
            ctypes.windll.user32.SetWindowLongW(hwnd, GWL_EXSTYLE, style)
            ctypes.windll.user32.SetWindowPos(
                hwnd,
                0,
                0,
                0,
                0,
                0,
                SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE | SWP_FRAMECHANGED,
            )
        except Exception:
            pass

    def send_to_desktop_bottom(self):
        try:
            self.root.update_idletasks()
            hwnd = ctypes.windll.user32.GetParent(self.root.winfo_id())
            if not hwnd:
                hwnd = self.root.winfo_id()
            ctypes.windll.user32.SetWindowPos(
                hwnd, HWND_BOTTOM, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE
            )
            self.root.lower()
        except Exception:
            pass

    def apply_window_layer(self):
        if self.config.get("window_layer", "desktop") == "topmost":
            self.root.attributes("-topmost", True)
            self.root.lift()
        else:
            self.root.attributes("-topmost", False)
            self.send_to_desktop_bottom()

    def set_window_layer(self):
        self.config["window_layer"] = self.window_layer_var.get()
        self.save_config()
        self.apply_window_layer()

    def on_focus_in(self, _event):
        if self.config.get("window_layer", "desktop") == "desktop":
            self.root.after(10, self.send_to_desktop_bottom)

    # -------------------- Screen 1: Startup Disclaimer Dialog --------------------

    def prompt_disclaimer(self):
        accepted = {"ok": False}
        dialog = tk.Toplevel(self.root)
        dialog.title("IMPORTANT NOTICE — Disclaimer & Terms of Use")
        dialog.resizable(False, False)
        dialog.attributes("-topmost", True)

        header = tk.Label(
            dialog,
            text="WARNING & DISCLAIMER OF WARRANTY",
            font=("Segoe UI", 11, "bold"),
            fg="#B91C1C",
        )
        header.pack(padx=16, pady=(14, 6), anchor="w")

        disclaimer_text = (
            'THIS SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, '
            "EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF "
            "MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE, AND NONINFRINGEMENT.\n\n"
            "IN NO EVENT SHALL THE CREATOR OR COPYRIGHT HOLDER BE LIABLE FOR ANY "
            "CLAIM, DAMAGES, DATA LOSS, SYSTEM MALFUNCTION, OR OTHER LIABILITY, "
            "WHETHER IN AN ACTION OF CONTRACT, TORT, OR OTHERWISE, ARISING FROM, "
            "OUT OF, OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS "
            "IN THE SOFTWARE—CAUSED TO ANYTHING OR ANYONE.\n\n"
            "Proceed and use this software ONLY if you fully understand these terms "
            "and voluntarily wish to use it at your own risk."
        )
        body = tk.Label(
            dialog,
            text=disclaimer_text,
            justify="left",
            wraplength=460,
            font=("Segoe UI", 9),
        )
        body.pack(padx=16, pady=(0, 12), anchor="w")

        agree_var = tk.BooleanVar(value=False)

        def on_toggle():
            ok_btn.config(state="normal" if agree_var.get() else "disabled")

        chk = ttk.Checkbutton(
            dialog,
            text="I have read, understood, and agree to the above disclaimer and wish to proceed.",
            variable=agree_var,
            command=on_toggle,
        )
        chk.pack(padx=16, pady=(0, 12), anchor="w")

        def on_ok():
            if agree_var.get():
                accepted["ok"] = True
                dialog.destroy()

        def on_close():
            accepted["ok"] = False
            dialog.destroy()

        dialog.protocol("WM_DELETE_WINDOW", on_close)

        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(pady=(0, 14))
        ok_btn = ttk.Button(btn_frame, text="OK", command=on_ok, state="disabled", width=12)
        ok_btn.pack(side="left", padx=8)
        close_btn = ttk.Button(btn_frame, text="Close", command=on_close, width=12)
        close_btn.pack(side="left", padx=8)

        dialog.update_idletasks()
        dw, dh = dialog.winfo_width(), dialog.winfo_height()
        sx = (dialog.winfo_screenwidth() - dw) // 2
        sy = (dialog.winfo_screenheight() - dh) // 2
        dialog.geometry(f"+{sx}+{sy}")

        dialog.grab_set()
        self.root.wait_window(dialog)
        return accepted["ok"]

    # -------------------- GitHub Link --------------------

    def open_github_page(self):
        try:
            webbrowser.open(GITHUB_URL)
        except Exception as exc:
            show_windows_popup(
                "Unable to Open Link",
                f"Could not open the GitHub page in your browser:\n\n{exc}",
                is_error=True,
            )

    # -------------------- Clean Exit --------------------

    def cancel_animation(self):
        if self._anim_id is not None:
            try:
                self.root.after_cancel(self._anim_id)
            except Exception:
                pass
            self._anim_id = None
        try:
            self.root.attributes("-alpha", 1.0)
            self.root.geometry(f"+{self.config['x']}+{self.config['y']}")
        except Exception:
            pass

    def exit_app(self, save_cfg=True):
        self._cancel_hide_controls()
        self.hide_tooltip()
        self.hide_toast()
        self.cancel_animation()
        if self._after_id is not None:
            try:
                self.root.after_cancel(self._after_id)
            except Exception:
                pass
            self._after_id = None
        if save_cfg:
            try:
                self.save_config()
            except Exception:
                pass
        try:
            self.root.quit()
            self.root.destroy()
        except Exception:
            pass

    # -------------------- Interval & Scheduling --------------------

    def get_interval_seconds(self):
        val = int(self.config.get("interval_value", 1))
        unit = self.config.get("interval_unit", "Hours")
        if val <= 0 or unit not in INTERVAL_UNITS:
            return 0
        unit_secs, max_val = INTERVAL_UNITS[unit]
        val = min(val, max_val)
        return val * unit_secs

    def get_interval_display_text(self):
        if self.is_paused:
            return "Quote changes in: Paused"
        val = int(self.config.get("interval_value", 1))
        if val <= 0 or self._next_change_ts is None:
            return "Quote changes in: Never (Manual)"

        remaining_secs = max(0, int(self._next_change_ts - time.time()))
        if remaining_secs < 60:
            return "Quote changes in: < 1 minute"

        total_mins = remaining_secs // 60
        days, rem_mins = divmod(total_mins, 1440)
        hours, mins = divmod(rem_mins, 60)

        parts = []
        if days:
            parts.append(f"{days} {'day' if days == 1 else 'days'}")
        if hours:
            parts.append(f"{hours} {'hour' if hours == 1 else 'hours'}")
        if mins or not parts:
            parts.append(f"{mins} {'minute' if mins == 1 else 'minutes'}")

        return f"Quote changes in: {', '.join(parts)}"

    def refresh_interval_menu_label(self):
        try:
            self.menu.entryconfig(0, label=self.get_interval_display_text())
        except Exception:
            pass

    def reset_schedule_timer(self):
        secs = self.get_interval_seconds()
        self._next_change_ts = (time.time() + secs) if secs > 0 else None

    def start_schedule_heartbeat(self):
        if not self.is_paused and self._next_change_ts is not None and time.time() >= self._next_change_ts:
            self.show_next_quote(animate=True, show_toast=False)
        self._after_id = self.root.after(1000, self.start_schedule_heartbeat)

    # -------------------- Config & Quotes --------------------

    def load_config(self, cfg_path):
        stencil_font = next(
            (fn for fn in self.available_fonts if "stencil" in fn.lower()), None
        )
        default_font = (
            stencil_font
            or ("Georgia" if "Georgia" in self.available_fonts else next(iter(self.available_fonts), "Arial"))
        )
        defaults = {
            "installed": False,
            "disclaimer_accepted": False,
            "categories_initialized": False,
            "install_dir": BASE_DIR,
            "run_at_startup": True,
            "excel_file": DEFAULT_EXCEL_FILE,
            "selected_categories": [],
            "undo_limit": DEFAULT_UNDO_LIMIT,
            "theme_mode": "light",  # 'light', 'dark', or 'system'
            "font_name": default_font,
            "font_size": 20,
            "wrap_width": 560,
            "text_align": "center",
            "text_color": "#FFFFFF",
            "text_alpha": 1.0,
            "text_fill": "Solid",
            "text_shadow": True,
            "bg_color": "#141414",
            "bg_alpha": 0.85,
            "bg_fill": "Solid",
            "bg_shape": "Rounded",
            "border_enabled": True,
            "border_color": "#FFFFFF",
            "border_pattern": "Solid",
            "order_mode": "random",
            "window_layer": "desktop",
            "seq_index": 0,
            "interval_value": 1,
            "interval_unit": "Hours",
            "randomize_style": True,
            "anim_enabled": True,
            "anim_mode": "Random",
            "x": 100,
            "y": 100,
        }
        if os.path.exists(cfg_path):
            try:
                with open(cfg_path, "r", encoding="utf-8") as f:
                    defaults.update(json.load(f))
            except Exception:
                pass
        return defaults

    def save_config(self):
        try:
            with open(self.config_file, "w", encoding="utf-8") as f:
                json.dump(self.config, f, indent=2)
        except Exception:
            pass

    def load_quote_entries(self):
        excel_path = self.config.get("excel_file", DEFAULT_EXCEL_FILE)
        if not excel_path or not os.path.exists(excel_path):
            return [(f"Excel file not found:\n{excel_path}", "Uncategorized")]
        wb = openpyxl.load_workbook(excel_path, read_only=True, data_only=True)
        ws = wb.active
        entries = []
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row or row[0] is None:
                continue
            q_text = str(row[0]).strip()
            c_text = str(row[1]).strip() if len(row) > 1 and row[1] is not None and str(row[1]).strip() else "Uncategorized"
            if q_text:
                entries.append((q_text, c_text))
        wb.close()
        return entries if entries else [("No quotes found in quotes.xlsx.", "Uncategorized")]

    def get_filtered_quotes(self):
        active_cats = {c for c, v in self.category_vars.items() if v.get()}
        filtered = [q for q, c in self.quote_entries if c in active_cats]
        return filtered if filtered else [q for q, _ in self.quote_entries]

    def find_font_path(self, preferred_names):
        for name in preferred_names:
            for k, v in self.available_fonts.items():
                if name.lower() in k.lower() and os.path.exists(v):
                    return v
        return None

    def get_pil_font(self, font_name, size, text=""):
        script = detect_script_family(text)
        path = None

        if script == "indic":
            path = self.find_font_path(
                ["Nirmala", "Mangal", "Kokila", "Aparajita", "Vrinda", "Latha", "Gautami", "Tunga", "Shruti", "Raavi", "Kartika"]
            )
        elif script == "rtl":
            path = self.find_font_path(["Segoeui", "Arial", "Tahoma", "Leelawad"])
        elif script == "cjk":
            path = self.find_font_path(["Msyh", "Yugoth", "Malgun", "Simsun", "Meiryo"])

        if not path:
            path = self.available_fonts.get(font_name)
            if not path:
                path = self.find_font_path([font_name, "Segoeui", "Arial"])

        if path and os.path.exists(path):
            try:
                return ImageFont.truetype(path, size)
            except Exception:
                pass
        return ImageFont.load_default()

    # -------------------- Menu Actions --------------------

    def toggle_startup(self):
        enabled = self.startup_var.get()
        self.config["run_at_startup"] = enabled
        self.save_config()
        set_windows_startup(enabled)

    def set_order_mode(self):
        self.config["order_mode"] = self.order_var.get()
        self.save_config()

    def set_text_align(self):
        self.config["text_align"] = self.text_align_var.get()
        self.save_config()
        self.render_widget()

    def toggle_anim_enabled(self):
        self.config["anim_enabled"] = self.anim_enabled_var.get()
        self.save_config()

    def set_anim_mode(self):
        self.config["anim_mode"] = self.anim_mode_var.get()
        self.save_config()

    def toggle_randomize_style(self):
        self.config["randomize_style"] = self.randomize_style_var.get()
        self.save_config()

    def toggle_shadow(self):
        self.config["text_shadow"] = self.shadow_var.get()
        self.save_config()
        self.render_widget()

    def toggle_border(self):
        self.config["border_enabled"] = self.border_var.get()
        self.save_config()
        self.render_widget()

    def apply_preset(self, preset, save=True, render=True):
        for k, v in preset.items():
            if k == "font_name":
                matched = next(
                    (fn for fn in self.available_fonts if v.lower() in fn.lower()),
                    self.config["font_name"],
                )
                self.config["font_name"] = matched
            else:
                self.config[k] = v
        self.shadow_var.set(self.config["text_shadow"])
        self.border_var.set(self.config["border_enabled"])
        if save:
            self.save_config()
        if render:
            self.render_widget()

    def show_next_quote(self, animate=True, show_toast=False):
        self.cancel_animation()
        self.quote_entries = self.load_quote_entries()

        updated_cats = sorted({cat for _, cat in self.quote_entries})
        if updated_cats != self.all_categories:
            for c in updated_cats:
                if c not in self.category_vars:
                    self.category_vars[c] = tk.BooleanVar(value=True)
            self.all_categories = updated_cats
            self.rebuild_category_menu()

        self.quotes = self.get_filtered_quotes()
        self._render_seed = random.randint(1, 999999)

        if self.history_pos < len(self.quote_history) - 1:
            self.history_pos += 1
            self.current_quote = self.quote_history[self.history_pos]
        else:
            if self.config["order_mode"] == "sequential":
                idx = self.config["seq_index"] % len(self.quotes)
                self.current_quote = self.quotes[idx]
                self.config["seq_index"] = (idx + 1) % len(self.quotes)
            else:
                if len(self.quotes) > 1 and self.current_quote in self.quotes:
                    choices = [q for q in self.quotes if q != self.current_quote]
                    self.current_quote = random.choice(choices)
                else:
                    self.current_quote = random.choice(self.quotes)

            self._push_quote_history(self.current_quote)

        if self.config.get("randomize_style", True):
            chosen_preset = random.choice(self.get_active_presets_pool())
            self.apply_preset(chosen_preset, save=False, render=False)

        self.save_config()
        if not self.is_paused:
            self.reset_schedule_timer()
        self.render_widget()

        if animate and self.config.get("anim_enabled", True):
            self.play_change_animation()

        if show_toast:
            self.show_toast("Next quote displayed")

    # -------------------- Animations --------------------

    def play_change_animation(self):
        mode = self.config.get("anim_mode", "Random")
        if mode == "Random" or mode not in ANIMATION_TYPES:
            chosen = random.choice(ANIMATION_TYPES)
        else:
            chosen = mode

        if chosen == "Jitter":
            self._anim_jitter()
        elif chosen == "Fade In":
            self._anim_fade_in()
        elif chosen == "Pop Scale":
            self._anim_pop_scale()
        elif chosen == "Border Pulse":
            self._anim_border_pulse()
        elif chosen == "Color Flash":
            self._anim_color_flash()

    def _anim_jitter(self):
        offsets = [(-7, -2), (7, 2), (-5, 2), (5, -2), (-3, 1), (3, -1), (0, 0)]
        base_x, base_y = self.config["x"], self.config["y"]

        def step(idx=0):
            if idx >= len(offsets):
                self._anim_id = None
                return
            dx, dy = offsets[idx]
            self.root.geometry(f"+{base_x + dx}+{base_y + dy}")
            self._anim_id = self.root.after(32, lambda: step(idx + 1))

        step()

    def _anim_fade_in(self):
        alphas = [0.15, 0.28, 0.42, 0.56, 0.70, 0.84, 0.94, 1.0]

        def step(idx=0):
            if idx >= len(alphas):
                self._anim_id = None
                return
            self.root.attributes("-alpha", alphas[idx])
            self._anim_id = self.root.after(35, lambda: step(idx + 1))

        step()

    def _anim_pop_scale(self):
        if self._last_chroma_img is None:
            return
        base_img = self._last_chroma_img
        w, h = base_img.size
        scales = [0.84, 0.90, 0.97, 1.04, 1.02, 1.0]

        def step(idx=0):
            if idx >= len(scales):
                self._anim_id = None
                return
            sc = scales[idx]
            if abs(sc - 1.0) < 0.001:
                frame = base_img
            else:
                nw, nh = max(1, int(w * sc)), max(1, int(h * sc))
                resized = base_img.resize((nw, nh), Image.BILINEAR)
                frame = Image.new("RGB", (w, h), CHROMA_RGB)
                frame.paste(resized, ((w - nw) // 2, (h - nh) // 2))
            self.photo_ref = ImageTk.PhotoImage(frame)
            self.display.config(image=self.photo_ref)
            self._anim_id = self.root.after(38, lambda: step(idx + 1))

        step()

    def _anim_border_pulse(self):
        boosts = [2, 6, 10, 6, 2, 0]

        def step(idx=0):
            if idx >= len(boosts):
                self._anim_id = None
                return
            self.render_widget(border_boost=boosts[idx])
            self._anim_id = self.root.after(48, lambda: step(idx + 1))

        step()

    def _anim_color_flash(self):
        if self._last_final_rgba is None:
            return
        base_rgba = self._last_final_rgba
        w, h = base_rgba.size
        grad = self.make_rainbow_image(w, h, alpha=255)
        alpha_mask = base_rgba.split()[3]
        mix_levels = [0.18, 0.42, 0.65, 0.48, 0.28, 0.12, 0.0]

        def step(idx=0):
            if idx >= len(mix_levels):
                self._anim_id = None
                return
            mix = mix_levels[idx]
            if mix <= 0.001:
                blended_rgba = base_rgba
            else:
                blended_rgba = Image.blend(base_rgba, grad, mix)
                blended_rgba.putalpha(alpha_mask)
            self._push_rgba_to_display(blended_rgba)
            self._anim_id = self.root.after(45, lambda: step(idx + 1))

        step()

    # -------------------- Pattern & Shape Generators --------------------

    def make_rainbow_image(self, w, h, alpha=255):
        img = Image.new("RGBA", (w, h), (0, 0, 0, alpha))
        draw = ImageDraw.Draw(img)
        for x in range(w):
            hue = x / max(1, w)
            rf, gf, bf = colorsys.hsv_to_rgb(hue, 0.85, 1.0)
            draw.line([(x, 0), (x, h)], fill=(int(rf * 255), int(gf * 255), int(bf * 255), alpha))
        return img

    def make_fill_texture(self, w, h, base_hex, style, alpha_float, seed=42):
        r, g, b = hex_to_rgb(base_hex)
        a = int(max(0, min(1, alpha_float)) * 255)
        img = Image.new("RGBA", (w, h), (r, g, b, a))
        draw = ImageDraw.Draw(img)
        rng = random.Random(seed)

        if style == "Gradient":
            r2, g2, b2 = min(255, r + 90), max(0, g - 40), min(255, b + 110)
            for y in range(h):
                t = y / max(1, h - 1)
                cr = int(r * (1 - t) + r2 * t)
                cg = int(g * (1 - t) + g2 * t)
                cb = int(b * (1 - t) + b2 * t)
                draw.line([(0, y), (w, y)], fill=(cr, cg, cb, a))
        elif style == "Glitter":
            count = max(120, (w * h) // 35)
            for _ in range(count):
                gx, gy = rng.randint(0, w - 1), rng.randint(0, h - 1)
                sz = rng.choice([1, 1, 2])
                sparkle = rng.choice(
                    [
                        (255, 255, 255, a),
                        (min(255, r + 120), min(255, g + 120), min(255, b + 120), a),
                        (max(0, r - 60), max(0, g - 60), max(0, b - 60), a),
                    ]
                )
                draw.rectangle([gx, gy, gx + sz, gy + sz], fill=sparkle)
        elif style == "Stripes":
            stripe_col = (min(255, r + 35), min(255, g + 35), min(255, b + 35), a)
            for i in range(-h, w + h, 14):
                draw.line([(i, 0), (i + h, h)], fill=stripe_col, width=4)
        elif style == "Dots":
            dot_col = (min(255, r + 45), min(255, g + 45), min(255, b + 45), a)
            for y in range(6, h, 18):
                for x in range(6, w, 18):
                    draw.ellipse([x - 2, y - 2, x + 2, y + 2], fill=dot_col)

        return img

    def make_shape_mask(self, w, h, shape, inset=10, top_offset=0):
        mask = Image.new("L", (w, h), 0)
        draw = ImageDraw.Draw(mask)
        x1, y1, x2, y2 = inset, inset + top_offset, w - inset, h - inset

        if shape == "Rectangle":
            draw.rectangle([x1, y1, x2, y2], fill=255)
        elif shape == "Rounded":
            draw.rounded_rectangle([x1, y1, x2, y2], radius=24, fill=255)
        elif shape == "Pill":
            draw.rounded_rectangle([x1, y1, x2, y2], radius=(y2 - y1) // 2, fill=255)
        elif shape == "Chamfered":
            c = min(22, (y2 - y1) // 3)
            pts = [
                (x1 + c, y1), (x2 - c, y1), (x2, y1 + c), (x2, y2 - c),
                (x2 - c, y2), (x1 + c, y2), (x1, y2 - c), (x1, y1 + c),
            ]
            draw.polygon(pts, fill=255)
        elif shape == "Ribbon":
            notch = min(24, (y2 - y1) // 2)
            ym = (y1 + y2) // 2
            pts = [
                (x1, y1), (x2, y1), (x2 - notch, ym), (x2, y2),
                (x1, y2), (x1 + notch, ym),
            ]
            draw.polygon(pts, fill=255)
        elif shape == "Banner":
            tip = min(26, (y2 - y1) // 2)
            ym = (y1 + y2) // 2
            pts = [
                (x1 + tip, y1), (x2 - tip, y1), (x2, ym),
                (x2 - tip, y2), (x1 + tip, y2), (x1, ym),
            ]
            draw.polygon(pts, fill=255)
        elif shape == "Leaf":
            rad = min(38, (y2 - y1) // 2)
            draw.rounded_rectangle([x1, y1, x2, y2], radius=rad, fill=255)
            draw.rectangle([x1, y1, x1 + rad, y1 + rad], fill=255)
            draw.rectangle([x2 - rad, y2 - rad, x2, y2], fill=255)
        elif shape == "Ticket":
            draw.rectangle([x1, y1, x2, y2], fill=255)
            nr = 15
            for cx, cy in [(x1, y1), (x2, y1), (x1, y2), (x2, y2)]:
                draw.ellipse([cx - nr, cy - nr, cx + nr, cy + nr], fill=0)
        elif shape == "Scalloped":
            draw.rounded_rectangle([x1 + 8, y1 + 8, x2 - 8, y2 - 8], radius=16, fill=255)
            r_sc = 11
            step = 18
            for x in range(x1 + 14, x2 - 10, step):
                draw.ellipse([x - r_sc, y1, x + r_sc, y1 + r_sc * 2], fill=255)
                draw.ellipse([x - r_sc, y2 - r_sc * 2, x + r_sc, y2], fill=255)
            for y in range(y1 + 14, y2 - 10, step):
                draw.ellipse([x1, y - r_sc, x1 + r_sc * 2, y + r_sc], fill=255)
                draw.ellipse([x2 - r_sc * 2, y - r_sc, x2, y + r_sc], fill=255)
        elif shape == "Cloud":
            draw.rounded_rectangle([x1 + 12, y1 + 12, x2 - 12, y2 - 12], radius=24, fill=255)
            step = 26
            for x in range(x1 + 18, x2 - 12, step):
                draw.ellipse([x - 16, y1, x + 16, y1 + 32], fill=255)
                draw.ellipse([x - 16, y2 - 32, x + 16, y2], fill=255)
            for y in range(y1 + 18, y2 - 12, step):
                draw.ellipse([x1, y - 16, x1 + 32, y + 16], fill=255)
                draw.ellipse([x2 - 32, y - 16, x2, y + 16], fill=255)
        elif shape == "Speech":
            tail_h = 16
            draw.rounded_rectangle([x1, y1, x2, y2 - tail_h], radius=22, fill=255)
            tx = x1 + 36
            draw.polygon(
                [(tx, y2 - tail_h - 2), (tx + 24, y2 - tail_h - 2), (tx - 6, y2)],
                fill=255,
            )
        else:
            draw.rounded_rectangle([x1, y1, x2, y2], radius=22, fill=255)

        return mask

    def draw_border(self, canvas_img, shape_mask, w, h, color_hex, pattern, boost=0):
        r, g, b = hex_to_rgb(color_hex)
        solid_col = Image.new("RGBA", (w, h), (r, g, b, 255))
        border_layer = Image.new("RGBA", (w, h), (0, 0, 0, 0))

        def odd(n):
            n = max(3, int(n))
            return n if n % 2 == 1 else n + 1

        def make_ring(mask, thickness_odd):
            inner = mask.filter(ImageFilter.MinFilter(odd(thickness_odd)))
            return ImageChops.subtract(mask, inner)

        t_std = odd(7 + boost)

        if pattern == "Solid":
            ring = make_ring(shape_mask, t_std)
            border_layer.paste(solid_col, (0, 0), ring)
        elif pattern == "Thick":
            ring = make_ring(shape_mask, odd(13 + boost))
            border_layer.paste(solid_col, (0, 0), ring)
        elif pattern == "Double":
            ring1 = make_ring(shape_mask, odd(5 + boost // 2))
            shrunk = shape_mask.filter(ImageFilter.MinFilter(9))
            ring2 = make_ring(shrunk, odd(5 + boost // 2))
            ring = ImageChops.lighter(ring1, ring2)
            border_layer.paste(solid_col, (0, 0), ring)
        elif pattern == "Triple":
            ring1 = make_ring(shape_mask, 3)
            s1 = shape_mask.filter(ImageFilter.MinFilter(7))
            ring2 = make_ring(s1, 3)
            s2 = s1.filter(ImageFilter.MinFilter(7))
            ring3 = make_ring(s2, odd(3 + boost // 2))
            ring = ImageChops.lighter(ImageChops.lighter(ring1, ring2), ring3)
            border_layer.paste(solid_col, (0, 0), ring)
        elif pattern == "Glow":
            ring = make_ring(shape_mask, t_std)
            border_layer.paste(solid_col, (0, 0), ring)
            glow = border_layer.filter(ImageFilter.GaussianBlur(radius=4 + boost // 2))
            border_layer = Image.alpha_composite(glow, border_layer)
        elif pattern == "Neon":
            outer_ring = make_ring(shape_mask, odd(9 + boost))
            border_layer.paste(solid_col, (0, 0), outer_ring)
            glow = border_layer.filter(ImageFilter.GaussianBlur(radius=5 + boost // 2))
            core_ring = make_ring(shape_mask.filter(ImageFilter.MinFilter(3)), 3)
            white_core = Image.new("RGBA", (w, h), (255, 255, 255, 255))
            core_layer = Image.new("RGBA", (w, h), (0, 0, 0, 0))
            core_layer.paste(white_core, (0, 0), core_ring)
            border_layer = Image.alpha_composite(Image.alpha_composite(glow, border_layer), core_layer)
        elif pattern == "Gradient":
            ring = make_ring(shape_mask, odd(9 + boost))
            grad_tex = self.make_fill_texture(w, h, color_hex, "Gradient", 1.0, seed=self._render_seed)
            border_layer.paste(grad_tex, (0, 0), ring)
        elif pattern == "Rainbow":
            ring = make_ring(shape_mask, odd(9 + boost))
            rb_tex = self.make_rainbow_image(w, h, alpha=255)
            border_layer.paste(rb_tex, (0, 0), ring)
        elif pattern == "Glitter":
            ring = make_ring(shape_mask, odd(9 + boost))
            glit_tex = self.make_fill_texture(w, h, color_hex, "Glitter", 1.0, seed=self._render_seed + 3)
            border_layer.paste(glit_tex, (0, 0), ring)
        elif pattern in ("Dashed", "Dotted"):
            ring = make_ring(shape_mask, t_std)
            dash_mask = Image.new("L", (w, h), 0)
            dm_draw = ImageDraw.Draw(dash_mask)
            step = 18 if pattern == "Dashed" else 8
            seg = 10 if pattern == "Dashed" else 3
            for i in range(-h, w + h, step):
                dm_draw.line([(i, 0), (i + h, h)], fill=255, width=seg)
            dashed_ring = ImageChops.multiply(ring, dash_mask)
            border_layer.paste(solid_col, (0, 0), dashed_ring)

        return Image.alpha_composite(canvas_img, border_layer)

    # -------------------- Main Render Pipeline --------------------

    def _push_rgba_to_display(self, final_rgba):
        w, h = final_rgba.size
        black_base = Image.new("RGBA", (w, h), (0, 0, 0, 255))
        blended = Image.alpha_composite(black_base, final_rgba).convert("RGB")
        chroma_img = Image.new("RGB", (w, h), CHROMA_RGB)
        visible_mask = final_rgba.split()[3].point(lambda p: 255 if p > 8 else 0)
        chroma_img.paste(blended, (0, 0), visible_mask)
        self._last_chroma_img = chroma_img
        self.photo_ref = ImageTk.PhotoImage(chroma_img)
        self.display.config(image=self.photo_ref)

    def wrap_text_lines(self, text, fnt, max_w, dummy_draw):
        lines = []
        for para in text.splitlines() or [""]:
            if not para.strip():
                lines.append("")
                continue
            words = para.split(" ")
            cur_line = ""
            for word in words:
                candidate = f"{cur_line} {word}".strip() if cur_line else word
                if cur_line and dummy_draw.textlength(candidate, font=fnt) > max_w:
                    lines.append(cur_line)
                    cur_line = word
                else:
                    cur_line = candidate
            if cur_line:
                lines.append(cur_line)
        return lines

    def render_widget(self, border_boost=0):
        cfg = self.config
        fnt = self.get_pil_font(cfg["font_name"], cfg["font_size"], self.current_quote)
        dummy = ImageDraw.Draw(Image.new("L", (10, 10)))
        max_w = cfg["wrap_width"]

        lines = self.wrap_text_lines(self.current_quote, fnt, max_w, dummy)

        sample_str = self.current_quote[:10] if self.current_quote else "Ag"
        bbox = fnt.getbbox(sample_str + "Ag")
        line_h = int(max(cfg["font_size"] + 6, (bbox[3] - bbox[1]) * 1.45))

        wide_shapes = ("Ribbon", "Banner", "Cloud", "Pill", "Scalloped")
        pad_x = 46 if cfg["bg_shape"] in wide_shapes else 32

        line_widths = [int(dummy.textlength(ln, font=fnt)) for ln in lines]
        content_w = max(line_widths or [140])
        content_h = max(line_h, len(lines) * line_h)
        w = max(140, content_w + pad_x * 2)

        # Compute multi-row top controls layout dynamically based on widget width
        self._button_rects, top_bar_h = self.compute_top_controls_layout(w)
        self._top_bar_h = top_bar_h

        pad_y_top = top_bar_h + 30
        pad_y_bot = 44 if cfg["bg_shape"] == "Speech" else 30
        h = max(96 + top_bar_h, content_h + pad_y_top + pad_y_bot)

        bg_layer = Image.new("RGBA", (w, h), (0, 0, 0, 0))
        shape_mask = self.make_shape_mask(w, h, cfg["bg_shape"], inset=10, top_offset=top_bar_h)
        if cfg["bg_alpha"] > 0.01:
            bg_tex = self.make_fill_texture(
                w, h, cfg["bg_color"], cfg["bg_fill"], cfg["bg_alpha"], seed=self._render_seed
            )
            bg_tex.putalpha(ImageChops.multiply(bg_tex.split()[3], shape_mask))
            bg_layer = bg_tex

        if cfg["border_enabled"] or border_boost > 0:
            bg_layer = self.draw_border(
                bg_layer,
                shape_mask,
                w,
                h,
                cfg["border_color"],
                cfg["border_pattern"],
                boost=border_boost,
            )

        text_mask = Image.new("L", (w, h), 0)
        tm_draw = ImageDraw.Draw(text_mask)
        align = cfg.get("text_align", "center")
        inner_w = w - pad_x * 2
        y = pad_y_top
        for ln, lw in zip(lines, line_widths):
            if ln:
                if align == "center":
                    tx = pad_x + (inner_w - lw) // 2
                elif align == "right":
                    tx = pad_x + (inner_w - lw)
                else:
                    tx = pad_x
                tm_draw.text((tx, y), ln, font=fnt, fill=255)
            y += line_h

        if cfg["text_shadow"]:
            shadow_mask = ImageChops.offset(text_mask, 3, 3).filter(
                ImageFilter.GaussianBlur(radius=2)
            )
            shadow_alpha = int(190 * cfg["text_alpha"])
            shadow_layer = Image.new("RGBA", (w, h), (0, 0, 0, shadow_alpha))
            shadow_layer.putalpha(
                ImageChops.multiply(shadow_mask, Image.new("L", (w, h), shadow_alpha))
            )
            bg_layer = Image.alpha_composite(bg_layer, shadow_layer)

        inv_text_mask = ImageChops.invert(text_mask)
        bg_layer.putalpha(ImageChops.multiply(bg_layer.split()[3], inv_text_mask))

        text_tex = self.make_fill_texture(
            w, h, cfg["text_color"], cfg["text_fill"], cfg["text_alpha"], seed=self._render_seed + 7
        )
        text_tex.putalpha(ImageChops.multiply(text_tex.split()[3], text_mask))
        final_rgba = Image.alpha_composite(bg_layer, text_tex)

        if self._controls_visible:
            final_rgba = self.draw_top_controls(final_rgba, self._button_rects)

        if border_boost == 0:
            self._last_final_rgba = final_rgba
        self._push_rgba_to_display(final_rgba)

        if self.config.get("window_layer", "desktop") == "desktop":
            self.send_to_desktop_bottom()

    # -------------------- Left-Click Buttons, Next Quote, Move & Resize --------------------

    def _hit_test_button(self, x, y):
        if not self._controls_visible:
            return None, None
        for name, tip, cx, cy, r in self._button_rects:
            if (x - cx) ** 2 + (y - cy) ** 2 <= (r + 2) ** 2:
                return name, tip
        return None, None

    def _hit_test(self, event):
        btn_name, _ = self._hit_test_button(event.x, event.y)
        if btn_name:
            return f"btn:{btn_name}"
        w = self.display.winfo_width()
        h = self.display.winfo_height()
        near_right = event.x >= w - 18
        near_bottom = event.y >= h - 18
        if near_right and near_bottom:
            return "resize_both"
        if near_right:
            return "resize_w"
        return "move"

    def on_hover(self, event):
        # Show all top circular buttons when hovering over the top area (or top edge of quote box)
        in_top_zone = event.y <= (self._top_bar_h + 28)
        if in_top_zone:
            self._cancel_hide_controls()
            if not self._controls_visible:
                self._controls_visible = True
                self.render_widget()
        else:
            self._schedule_hide_controls()

        btn_name, tip = self._hit_test_button(event.x, event.y)
        if btn_name:
            self.display.config(cursor="hand2")
            if self._hovered_btn != btn_name:
                self._hovered_btn = btn_name
                self.show_tooltip(tip, event.x_root, event.y_root)
            return

        if self._hovered_btn is not None:
            self.hide_tooltip()

        mode = self._hit_test(event)
        if mode == "resize_both":
            self.display.config(cursor="size_nw_se")
        elif mode == "resize_w":
            self.display.config(cursor="size_we")
        else:
            self.display.config(cursor="hand2")

    def on_mouse_leave(self, _event):
        self.hide_tooltip()
        self._schedule_hide_controls()

    def on_mouse_down(self, event):
        self.hide_tooltip()
        self.cancel_animation()
        self._interaction_mode = self._hit_test(event)
        self._did_drag = False
        self._drag_x = event.x
        self._drag_y = event.y
        self._resize_start_x = event.x_root
        self._resize_start_y = event.y_root
        self._resize_orig_w = self.config["wrap_width"]
        self._resize_orig_fs = self.config["font_size"]

    def on_mouse_drag(self, event):
        if self._interaction_mode.startswith("btn:"):
            return
        dx = event.x_root - self._resize_start_x
        dy = event.y_root - self._resize_start_y
        if not self._did_drag and (abs(dx) > 4 or abs(dy) > 4):
            self._did_drag = True
        if not self._did_drag:
            return

        if self._interaction_mode == "move":
            x = self.root.winfo_x() + (event.x - self._drag_x)
            y = self.root.winfo_y() + (event.y - self._drag_y)
            self.root.geometry(f"+{x}+{y}")
        else:
            now = time.time()
            if now - self._last_resize_ts < 0.04:
                return
            self._last_resize_ts = now

            new_w = max(220, min(1400, int(self._resize_orig_w + dx)))
            self.config["wrap_width"] = new_w

            if self._interaction_mode == "resize_both":
                new_fs = max(12, min(72, int(self._resize_orig_fs + dy / 5)))
                self.config["font_size"] = new_fs

            self.render_widget()

    def on_mouse_up(self, event):
        if self._interaction_mode.startswith("btn:"):
            btn_name, _ = self._hit_test_button(event.x, event.y)
            if btn_name and f"btn:{btn_name}" == self._interaction_mode:
                self.handle_control_button(btn_name)
            return

        if not self._did_drag:
            self.show_next_quote(animate=True, show_toast=True)
            return
        self.config["x"] = self.root.winfo_x()
        self.config["y"] = self.root.winfo_y()
        self.save_config()
        if self.config.get("window_layer", "desktop") == "desktop":
            self.send_to_desktop_bottom()

    def show_menu(self, event):
        self.hide_tooltip()
        self.refresh_interval_menu_label()
        # Dynamically sync menu colors if following Windows System theme
        if self.config.get("theme_mode", "light") == "system":
            self.apply_menu_theme()
        self.menu.tk_popup(event.x_root, event.y_root)

    # -------------------- Dialog 0: Change Interval --------------------

    def open_interval_dialog(self):
        dialog = tk.Toplevel(self.root)
        dialog.title("Set Quote Change Interval")
        dialog.resizable(False, False)
        dialog.attributes("-topmost", True)

        val_var = tk.IntVar(value=int(self.config.get("interval_value", 1)))
        unit_var = tk.StringVar(value=self.config.get("interval_unit", "Hours"))

        ttk.Label(
            dialog, text="Set 0 for no automatic change (manual only).\nMaximum allowed duration: 1 Year."
        ).grid(row=0, column=0, columnspan=2, padx=12, pady=(10, 8), sticky="w")

        ttk.Label(dialog, text="Duration:").grid(row=1, column=0, padx=12, pady=6, sticky="w")
        spin_val = ttk.Spinbox(dialog, from_=0, to=8760, textvariable=val_var, width=10)
        spin_val.grid(row=1, column=1, padx=12, pady=6, sticky="w")

        ttk.Label(dialog, text="Unit:").grid(row=2, column=0, padx=12, pady=6, sticky="w")
        cb_unit = ttk.Combobox(
            dialog,
            textvariable=unit_var,
            values=list(INTERVAL_UNITS.keys()),
            width=12,
            state="readonly",
        )
        cb_unit.grid(row=2, column=1, padx=12, pady=6, sticky="w")

        def on_unit_change(*_args):
            unit = unit_var.get()
            if unit in INTERVAL_UNITS:
                max_v = INTERVAL_UNITS[unit][1]
                spin_val.config(to=max_v)
                try:
                    if val_var.get() > max_v:
                        val_var.set(max_v)
                except tk.TclError:
                    val_var.set(1)

        cb_unit.bind("<<ComboboxSelected>>", on_unit_change)
        on_unit_change()

        def apply_interval():
            try:
                v = max(0, int(val_var.get()))
            except tk.TclError:
                v = 0
            u = unit_var.get()
            max_v = INTERVAL_UNITS.get(u, (3600, 8760))[1]
            v = min(v, max_v)
            self.config["interval_value"] = v
            self.config["interval_unit"] = u
            self.save_config()
            if not self.is_paused:
                self.reset_schedule_timer()
            dialog.destroy()

        btns = ttk.Frame(dialog)
        btns.grid(row=3, column=0, columnspan=2, pady=12)
        ttk.Button(btns, text="Apply", command=apply_interval).pack(side="left", padx=6)
        ttk.Button(btns, text="Cancel", command=dialog.destroy).pack(side="left", padx=6)

    # -------------------- Dialog 1: Text Style --------------------

    def open_text_dialog(self):
        dialog = tk.Toplevel(self.root)
        dialog.title("Text Style Settings")
        dialog.resizable(False, False)
        dialog.attributes("-topmost", True)

        backup = dict(self.config)
        r0, g0, b0 = hex_to_rgb(self.config["text_color"])

        font_var = tk.StringVar(value=self.config["font_name"])
        size_var = tk.IntVar(value=self.config["font_size"])
        width_var = tk.IntVar(value=self.config["wrap_width"])
        fill_var = tk.StringVar(value=self.config["text_fill"])
        alpha_var = tk.DoubleVar(value=self.config["text_alpha"])
        r_var, g_var, b_var = tk.IntVar(value=r0), tk.IntVar(value=g0), tk.IntVar(value=b0)

        def live_update(*_args):
            try:
                self.config["font_name"] = font_var.get()
                self.config["font_size"] = size_var.get()
                self.config["wrap_width"] = width_var.get()
                self.config["text_fill"] = fill_var.get()
                self.config["text_alpha"] = round(alpha_var.get(), 2)
                self.config["text_color"] = rgb_to_hex(r_var.get(), g_var.get(), b_var.get())
                preview.config(bg=self.config["text_color"])
                pct_lbl.config(text=f"{int(round(alpha_var.get() * 100))}%")
                self.render_widget()
            except tk.TclError:
                pass

        ttk.Label(dialog, text="Font:").grid(row=0, column=0, padx=10, pady=5, sticky="w")
        cb_font = ttk.Combobox(
            dialog,
            textvariable=font_var,
            values=list(self.available_fonts.keys()),
            width=24,
            state="readonly",
        )
        cb_font.grid(row=0, column=1, columnspan=2, padx=10, pady=5, sticky="ew")
        cb_font.bind("<<ComboboxSelected>>", live_update)

        ttk.Label(dialog, text="Font Size:").grid(row=1, column=0, padx=10, pady=5, sticky="w")
        ttk.Spinbox(dialog, from_=12, to=72, textvariable=size_var, width=8, command=live_update).grid(
            row=1, column=1, padx=10, pady=5, sticky="w"
        )

        ttk.Label(dialog, text="Max Width (px):").grid(row=2, column=0, padx=10, pady=5, sticky="w")
        ttk.Spinbox(
            dialog, from_=220, to=1400, increment=40, textvariable=width_var, width=8, command=live_update
        ).grid(row=2, column=1, padx=10, pady=5, sticky="w")

        ttk.Label(dialog, text="Text Fill Style:").grid(row=3, column=0, padx=10, pady=5, sticky="w")
        cb_fill = ttk.Combobox(
            dialog,
            textvariable=fill_var,
            values=["Solid", "Glitter", "Gradient", "Stripes", "Dots"],
            width=16,
            state="readonly",
        )
        cb_fill.grid(row=3, column=1, columnspan=2, padx=10, pady=5, sticky="w")
        cb_fill.bind("<<ComboboxSelected>>", live_update)

        ttk.Label(dialog, text="Text Color:").grid(row=4, column=0, padx=10, pady=5, sticky="w")
        preview = tk.Label(dialog, bg=self.config["text_color"], width=10, height=1, relief="solid", bd=1)
        preview.grid(row=4, column=1, padx=10, pady=5, sticky="w")

        def pick_color():
            rgb_t, hex_c = colorchooser.askcolor(initialcolor=self.config["text_color"], parent=dialog)
            if rgb_t and hex_c:
                r_var.set(int(rgb_t[0]))
                g_var.set(int(rgb_t[1]))
                b_var.set(int(rgb_t[2]))
                live_update()

        ttk.Button(dialog, text="Palette...", command=pick_color).grid(row=4, column=2, padx=10, pady=5)

        rgb_f = ttk.LabelFrame(dialog, text="Text RGB (0 - 255)")
        rgb_f.grid(row=5, column=0, columnspan=3, padx=10, pady=6, sticky="ew")
        for idx, (lbl, var) in enumerate([("R:", r_var), ("G:", g_var), ("B:", b_var)]):
            ttk.Label(rgb_f, text=lbl).grid(row=0, column=idx * 2, padx=(8, 2), pady=5)
            ttk.Spinbox(rgb_f, from_=0, to=255, textvariable=var, width=5, command=live_update).grid(
                row=0, column=idx * 2 + 1, padx=(0, 8), pady=5
            )
            var.trace_add("write", live_update)

        ttk.Label(dialog, text="Text Opacity:").grid(row=6, column=0, padx=10, pady=6, sticky="w")
        ttk.Scale(dialog, from_=0.2, to=1.0, variable=alpha_var, command=live_update).grid(
            row=6, column=1, padx=6, pady=6, sticky="ew"
        )
        pct_lbl = ttk.Label(dialog, text=f"{int(round(alpha_var.get() * 100))}%", width=5)
        pct_lbl.grid(row=6, column=2, padx=(0, 10), pady=6)

        def save_close():
            self.save_config()
            dialog.destroy()

        def cancel_close():
            self.config = backup
            self.render_widget()
            dialog.destroy()

        dialog.protocol("WM_DELETE_WINDOW", cancel_close)
        btns = ttk.Frame(dialog)
        btns.grid(row=7, column=0, columnspan=3, pady=10)
        ttk.Button(btns, text="Apply", command=save_close).pack(side="left", padx=6)
        ttk.Button(btns, text="Cancel", command=cancel_close).pack(side="left", padx=6)

    # -------------------- Dialog 2: Background & Border --------------------

    def open_bg_dialog(self):
        dialog = tk.Toplevel(self.root)
        dialog.title("Background & Border Settings")
        dialog.resizable(False, False)
        dialog.attributes("-topmost", True)

        backup = dict(self.config)
        r0, g0, b0 = hex_to_rgb(self.config["bg_color"])

        shape_var = tk.StringVar(value=self.config["bg_shape"])
        fill_var = tk.StringVar(value=self.config["bg_fill"])
        alpha_var = tk.DoubleVar(value=self.config["bg_alpha"])
        border_pat_var = tk.StringVar(value=self.config["border_pattern"])
        r_var, g_var, b_var = tk.IntVar(value=r0), tk.IntVar(value=g0), tk.IntVar(value=b0)

        def live_update(*_args):
            try:
                self.config["bg_shape"] = shape_var.get()
                self.config["bg_fill"] = fill_var.get()
                self.config["bg_alpha"] = round(alpha_var.get(), 2)
                self.config["border_pattern"] = border_pat_var.get()
                self.config["bg_color"] = rgb_to_hex(r_var.get(), g_var.get(), b_var.get())
                preview.config(bg=self.config["bg_color"])
                pct_lbl.config(text=f"{int(round(alpha_var.get() * 100))}%")
                self.render_widget()
            except tk.TclError:
                pass

        ttk.Label(dialog, text="Shape:").grid(row=0, column=0, padx=10, pady=5, sticky="w")
        cb_shape = ttk.Combobox(
            dialog,
            textvariable=shape_var,
            values=AVAILABLE_SHAPES,
            width=16,
            state="readonly",
        )
        cb_shape.grid(row=0, column=1, columnspan=2, padx=10, pady=5, sticky="w")
        cb_shape.bind("<<ComboboxSelected>>", live_update)

        ttk.Label(dialog, text="Background Fill:").grid(row=1, column=0, padx=10, pady=5, sticky="w")
        cb_fill = ttk.Combobox(
            dialog,
            textvariable=fill_var,
            values=["Solid", "Glitter", "Gradient", "Stripes", "Dots"],
            width=16,
            state="readonly",
        )
        cb_fill.grid(row=1, column=1, columnspan=2, padx=10, pady=5, sticky="w")
        cb_fill.bind("<<ComboboxSelected>>", live_update)

        ttk.Label(dialog, text="Background Color:").grid(row=2, column=0, padx=10, pady=5, sticky="w")
        preview = tk.Label(dialog, bg=self.config["bg_color"], width=10, height=1, relief="solid", bd=1)
        preview.grid(row=2, column=1, padx=10, pady=5, sticky="w")

        def pick_bg_color():
            rgb_t, hex_c = colorchooser.askcolor(initialcolor=self.config["bg_color"], parent=dialog)
            if rgb_t and hex_c:
                r_var.set(int(rgb_t[0]))
                g_var.set(int(rgb_t[1]))
                b_var.set(int(rgb_t[2]))
                live_update()

        ttk.Button(dialog, text="Palette...", command=pick_bg_color).grid(row=2, column=2, padx=10, pady=5)

        rgb_f = ttk.LabelFrame(dialog, text="Background RGB (0 - 255)")
        rgb_f.grid(row=3, column=0, columnspan=3, padx=10, pady=6, sticky="ew")
        for idx, (lbl, var) in enumerate([("R:", r_var), ("G:", g_var), ("B:", b_var)]):
            ttk.Label(rgb_f, text=lbl).grid(row=0, column=idx * 2, padx=(8, 2), pady=5)
            ttk.Spinbox(rgb_f, from_=0, to=255, textvariable=var, width=5, command=live_update).grid(
                row=0, column=idx * 2 + 1, padx=(0, 8), pady=5
            )
            var.trace_add("write", live_update)

        ttk.Label(dialog, text="BG Opacity:").grid(row=4, column=0, padx=10, pady=6, sticky="w")
        ttk.Scale(dialog, from_=0.0, to=1.0, variable=alpha_var, command=live_update).grid(
            row=4, column=1, padx=6, pady=6, sticky="ew"
        )
        pct_lbl = ttk.Label(dialog, text=f"{int(round(alpha_var.get() * 100))}%", width=5)
        pct_lbl.grid(row=4, column=2, padx=(0, 10), pady=6)

        ttk.Label(dialog, text="Border Pattern:").grid(row=5, column=0, padx=10, pady=5, sticky="w")
        cb_bpat = ttk.Combobox(
            dialog,
            textvariable=border_pat_var,
            values=AVAILABLE_BORDER_PATTERNS,
            width=16,
            state="readonly",
        )
        cb_bpat.grid(row=5, column=1, padx=10, pady=5, sticky="w")
        cb_bpat.bind("<<ComboboxSelected>>", live_update)

        def pick_border_color():
            _, hex_c = colorchooser.askcolor(initialcolor=self.config["border_color"], parent=dialog)
            if hex_c:
                self.config["border_color"] = hex_c
                self.render_widget()

        ttk.Button(dialog, text="Border Color...", command=pick_border_color).grid(
            row=5, column=2, padx=10, pady=5
        )

        def save_close():
            self.save_config()
            dialog.destroy()

        def cancel_close():
            self.config = backup
            self.render_widget()
            dialog.destroy()

        dialog.protocol("WM_DELETE_WINDOW", cancel_close)
        btns = ttk.Frame(dialog)
        btns.grid(row=6, column=0, columnspan=3, pady=10)
        ttk.Button(btns, text="Apply", command=save_close).pack(side="left", padx=6)
        ttk.Button(btns, text="Cancel", command=cancel_close).pack(side="left", padx=6)


def main():
    mutex, already_running = acquire_single_instance_lock()
    if already_running:
        show_windows_popup(
            "Quote Widget Already Running",
            "An instance of Quote Widget is already running on your desktop.\n"
            "Only one instance can run at a time.",
            is_error=False,
        )
        sys.exit(0)

    app = None
    try:
        app = QuoteWidget(mutex_handle=mutex)
    except Exception as exc:
        err_details = traceback.format_exc()
        show_windows_popup(
            "Quote Widget - Failed to Start",
            f"Quote Widget failed to start because of the following error:\n\n"
            f"{exc}\n\nDetails:\n{err_details}",
            is_error=True,
        )
        sys.exit(1)
    finally:
        if app is None or app.mutex_handle is not None:
            release_single_instance_lock(mutex)


if __name__ == "__main__":
    main()
