# Windows Quote Widget (Version 2) - VIBE CODED IN GEMINI BY GOOGLE

A lightweight, offline, and customizable desktop quote widget for Windows that displays inspiring quotes directly on your desktop without appearing in the Windows Taskbar. Built with Python, it pairs a frameless, customizable UI with Excel-based quote management, hover-activated controls, and 52 Dark and Light mode presets.

---

## Key Features

### 1. Excel-Based Quote Management (`quotes.xlsx`)

* **Ready Out of the Box:** Automatically generates a formatted `quotes.xlsx` file on first launch containing **100 original quotes** across **5 categories** (*Atomic Habits*, *Health Motivation*, *Study Motivation*, *Finance*, and *Good Moral Values*).


* **Simple Two-Column Structure:** Reads quotes starting from Row 2, with `Quote` in Column A and `Category` in Column B. Any quote with a blank category cell is automatically assigned to `Uncategorized`.


* **Multi-Language & Unicode Support:** Renders Latin, Indic, Right-to-Left (Hindi/Indian Languages/Arabic/Hebrew), and CJK (Chinese/Japanese/Korean) scripts with automatic Windows Unicode font fallback.


* **Direct Spreadsheet Editing:** Includes an **Edit Excel File...** option in the context menu to open `quotes.xlsx` directly in your default spreadsheet editor.


* **Live Category Filtering:** Enable or disable individual categories anytime from the **Quote Categories** submenu, with newly added Excel categories syncing automatically.



### 2. Hover-Activated Quick Controls

* **Distraction-Free Default View:** All top circular control buttons remain hidden while idle and appear simultaneously when hovering over the top area of the widget.


* **Top-Left Control Buttons:**
* **Previous Quote:** Step back through up to 2 previously displayed quotes with style randomization support.


* **Next Quote:** Advance immediately to the next quote.


* **Pause / Resume Timer:** Pause or resume automatic quote rotation while preserving the remaining countdown duration.


* **Add New Quote:** Opens a quick dialog to add a quote in any language to an existing category, `Uncategorized`, or a newly created category.


* **Delete Current Quote (`-`):** Removes the currently displayed quote directly from `quotes.xlsx`.


* **Undo (Add / Delete):** Reverses recent quote additions or deletions, with a user-configurable history limit between **1 and 20** (default: **5**).




* **Top-Right Close Button (`X`):** A red circular button at the top-right corner to exit the widget directly.


* **Responsive Multi-Row Wrapping:** Automatically wraps the control buttons into multiple rows above the quote box when the widget width is narrowed.


* **3-Second Toast Notifications:** Displays a brief, theme-matched confirmation toast near the widget whenever an action is performed.



### 3. Visual Styling & 52 Mode-Categorized Presets

* **Dark & Light Mode Presets:** Includes **52 built-in presets** divided into **28 Dark Mode Presets** and **24 Light Mode Presets**. When *Randomize Style on Quote Change* is enabled, the widget automatically picks from the preset pool matching your active appearance mode.


* **11 Widget Shapes:** *Rounded, Pill, Leaf, Ticket, Scalloped, Cloud, Banner, Ribbon, Chamfered, Speech,* and *Rectangle*.


* **11 Border Patterns:** *Solid, Thick, Dashed, Dotted, Double, Triple, Glow, Neon, Gradient, Rainbow,* and *Glitter*.


* **5 Fill Textures:** Apply *Solid, Glitter, Gradient, Stripes,* or *Dots* independently to text and background.


* **Full Typography & Color Controls:** Automatically scans installed Windows fonts and provides live-preview dialogs for font size (12–72 px), wrap width (220–1400 px), Center/Left/Right alignment, text shadow, RGB/palette color pickers, and independent text and background opacity.



### 4. Appearance, Animations & Scheduling

* **Context Menu Themes:** Supports **Light Mode** (default), **Dark Mode**, and **System (Follow Windows)**, which dynamically reads Windows' `AppsUseLightTheme` registry setting.


* **5 Transition Animations:** *Color Flash, Border Pulse, Jitter, Fade In,* and *Pop Scale* (with a *Randomize* mode or toggle to disable).


* **Flexible Rotation Timer:** Configure automatic quote change intervals in *Minutes, Hours, Days, Weeks, Months,* or *Years* (up to 1 year), or set `0` for manual-only operation.


* **Playback Order:** Supports **Random** (default) and **In Sequence** playback.



### 5. Desktop Integration & Usability

* **Direct Mouse Gestures:** Left-click the quote to advance, drag anywhere to reposition, drag the right edge to resize width, or drag the bottom-right corner to resize both width and font size.


* **Display Layering:** Pin the widget to the desktop background (*Show Only on Desktop*) or keep it floating (*Show Above All Windows*).


* **Taskbar-Free & Single Instance:** Uses `WS_EX_TOOLWINDOW` to stay hidden from the Windows Taskbar and a Windows Named Mutex to prevent multiple instances from running simultaneously.


* **One-Click Reset:** **Reset Settings to Default...** restores all visual and behavioral settings while leaving your `quotes.xlsx` file untouched.



### 6. Built-In Setup & Safe Uninstallation

* **Guided First-Run Setup:** Displays a disclaimer, lets you choose an installation directory (defaulting to `C:\Windows Quote Widget`), creates **Desktop** and **Start Menu** shortcuts, and optionally registers for **Windows Startup**.


* **Smart Reinstall Handling:** If a validly formatted `quotes.xlsx` already exists in the chosen folder, the setup prompts you to either **Retain** your existing file or **Overwrite** it with the defaults.


* **Instant Subsequent Launches:** Tracks installation state (`.installed` marker and `%LOCALAPPDATA%\WindowsQuoteWidget\install_state.json`) so future restarts skip setup dialogs and launch the widget immediately.


* **Safe Uninstaller:** The **Uninstall App...** option removes all application files, shortcuts, startup entries, and state files while **preserving `quotes.xlsx**` and providing a clickable path to locate it in File Explorer.



---

## Installation & Usage

### Option 1: Standalone Executable (Recommended — No Python Required)

* **OS:** Windows 10 / 11


* **No Python Installation Needed:** The `.exe` file is completely self-contained and bundles the Python runtime and all required libraries—end users do not need Python installed on their system to run it.
* **Getting Started:** Double-click the `.exe` file, accept the disclaimer, select your installation folder (defaults to `C:\Windows Quote Widget`), and choose your preferred quote categories. Desktop and Start Menu shortcuts are created automatically.



### Option 2: Running from Source (For Developers)

* **OS:** Windows 10 / 11


* **Python:** Python 3.8+
* **Install Dependencies:**
```bash
pip install openpyxl pillow

```


* **Run:**
```bash
python quote_widget.py

```



### Building the Standalone `.exe` (PyInstaller)

Place `app_icon.ico` in the same directory as `quote_widget.py` and run:

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --icon=app_icon.ico --add-data "app_icon.ico;." quote_widget.py

```

---

## Author

* **GitHub:** [github.com/ManishKashyapIndia](https://github.com/ManishKashyapIndia)


* **Version:** 2
