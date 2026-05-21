import streamlit as st
import os
import sys
import re
import openpyxl
import pandas as pd
import subprocess
import time
import logging
import warnings

# Silence noisy deprecation and user warnings
warnings.filterwarnings("ignore", category=UserWarning)
warnings.filterwarnings("ignore", category=DeprecationWarning)

# Mute Streamlit's noisy local sources watcher tracebacks for missing optional third-party modules
logging.getLogger("streamlit.watcher.local_sources_watcher").setLevel(logging.ERROR)

# Configure page layout and style
# Initialize session state variables
if "theme" not in st.session_state:
    st.session_state["theme"] = "light"

theme = st.session_state["theme"]

if theme == "dark":
    bg_color = "#0f172a"
    text_color = "#f8fafc"
    card_bg = "rgba(30, 41, 59, 0.7)"
    card_border = "rgba(51, 65, 85, 0.8)"
    subtitle_color = "#94a3b8"
    input_border = "#334155"
    shadow_color = "rgba(0, 0, 0, 0.3)"
    popover_bg = "#1e293b"
    placeholder_color = "#94a3b8"
    alert_bg = "#1e293b"
    alert_border = "#334155"
else:
    bg_color = "#f8fafc"
    text_color = "#0f172a"
    card_bg = "rgba(255, 255, 255, 0.8)"
    card_border = "rgba(226, 232, 240, 0.8)"
    subtitle_color = "#64748b"
    input_border = "#e2e8f0"
    shadow_color = "rgba(31, 38, 135, 0.05)"
    popover_bg = "#ffffff"
    placeholder_color = "#64748b"
    alert_bg = "rgba(241, 245, 249, 0.8)"
    alert_border = "rgba(226, 232, 240, 0.8)"

# Configure page layout and style
st.set_page_config(
    page_title="TemplateMachine Control Center",
    page_icon="🚀",
    layout="wide",
    initial_sidebar_state="collapsed"
)



# Custom premium styling
st.markdown(f"""
<style>
    /* Google Fonts */
    @import url('https://fonts.googleapis.com/css2?family=Outfit:wght@300;400;500;600;700&family=JetBrains+Mono:wght@400;700&display=swap');
    
    :root {{
        --background-color: {bg_color};
        --text-color: {text_color};
        --secondary-background-color: {popover_bg};
    }}
    
    html, body, [data-testid="stAppViewContainer"] {{
        background-color: {bg_color} !important;
        color: {text_color} !important;
        transition: background-color 0.3s ease, color 0.3s ease;
    }}
    
    /* Hide Streamlit header, footer, and sidebar & remove top spacing */
    [data-testid="stHeader"], footer, [data-testid="stSidebar"] {{
        display: none !important;
    }}
    
    [data-testid="stMainBlockContainer"], .block-container {{
        padding-top: 0.5rem !important;
        margin-top: 0px !important;
    }}
    
    html, body, [class*="css"] {{
        font-family: 'Outfit', sans-serif;
    }}
    
    h1:not(.main-title), h2, h3, h4, h5, h6, p, span, label, li {{
        color: {text_color} !important;
    }}
    
    code, pre {{
        font-family: 'JetBrains Mono', monospace !important;
        background-color: {popover_bg} !important;
        color: {text_color} !important;
        border: 1px solid {card_border} !important;
        border-radius: 6px !important;
        padding: 0.2rem 0.4rem !important;
    }}
    
    pre {{
        padding: 1rem !important;
        overflow-x: auto !important;
        background-color: {popover_bg} !important;
    }}
    
    pre code {{
        background-color: transparent !important;
        border: none !important;
        padding: 0 !important;
    }}
    
    /* Elegant Title and Badges */
    .main-title {{
        background: linear-gradient(135deg, #4A90E2 0%, #50E3C2 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        font-size: 3rem;
        font-weight: 700;
        margin-bottom: 0.2rem;
        letter-spacing: -0.5px;
    }}
    .subtitle {{
        color: {subtitle_color} !important;
        font-size: 1.15rem;
        margin-bottom: 2rem;
        font-weight: 300;
    }}
    
    /* Premium Styled Card Container */
    .card {{
        background: {card_bg};
        backdrop-filter: blur(10px);
        -webkit-backdrop-filter: blur(10px);
        border-radius: 16px;
        box-shadow: 0 8px 32px 0 {shadow_color};
        padding: 1.8rem;
        margin-bottom: 1.8rem;
        border: 1px solid {card_border};
        color: {text_color};
        transition: all 0.3s ease;
    }}
    
    /* Styled widgets & alerts */
    .stAlert, div[data-testid="stAlert"] {{
        border-radius: 12px !important;
        border: 1px solid {alert_border} !important;
        background-color: {alert_bg} !important;
        color: {text_color} !important;
        box-shadow: 0 4px 12px rgba(0,0,0,0.02) !important;
    }}
    
    /* Button custom hover effects */
    button[kind="primary"] {{
        background: linear-gradient(135deg, #4A90E2 0%, #357ABD 100%) !important;
        border: none !important;
        color: white !important;
        box-shadow: 0 4px 15px rgba(74, 144, 226, 0.3) !important;
        transition: all 0.25s ease-in-out !important;
    }}
    button[kind="primary"]:hover {{
        transform: translateY(-2px) !important;
        box-shadow: 0 6px 20px rgba(74, 144, 226, 0.4) !important;
    }}
    
    button[kind="secondary"] {{
        border-radius: 8px !important;
        border: 1px solid {input_border} !important;
        background-color: {popover_bg} !important;
        color: {text_color} !important;
        transition: all 0.2s ease !important;
    }}
    button[kind="secondary"]:hover {{
        border-color: #4A90E2 !important;
        color: #4A90E2 !important;
        background-color: rgba(74, 144, 226, 0.03) !important;
    }}
    
    /* Code block copy toolbar buttons styled */
    button[kind="elementToolbar"], button[data-testid="stBaseButton-elementToolbar"] {{
        background-color: {popover_bg} !important;
        color: {text_color} !important;
        border: 1px solid {card_border} !important;
        border-radius: 6px !important;
    }}
    button[kind="elementToolbar"] svg, button[data-testid="stBaseButton-elementToolbar"] svg {{
        fill: {text_color} !important;
        color: {text_color} !important;
    }}
    button[kind="elementToolbar"]:hover, button[data-testid="stBaseButton-elementToolbar"]:hover {{
        background-color: #4A90E2 !important;
        color: white !important;
        border-color: #4A90E2 !important;
    }}
    button[kind="elementToolbar"]:hover svg, button[data-testid="stBaseButton-elementToolbar"]:hover svg {{
        fill: white !important;
        color: white !important;
    }}
    
    /* Tooltip popup styled */
    div[data-baseweb="tooltip"],
    div[data-baseweb="tooltip"] *,
    div[data-testid="stTooltipContent"],
    div[data-testid="stTooltipContent"] * {{
        background-color: {popover_bg} !important;
        color: {text_color} !important;
    }}
    div[data-baseweb="tooltip"] {{
        border: 1px solid {card_border} !important;
        border-radius: 8px !important;
        box-shadow: 0 4px 12px {shadow_color} !important;
    }}
    
    .badge-icon {{
        font-size: 1.5rem;
        margin-right: 0.5rem;
    }}
    
    /* Input fields and Selectboxes custom styling */
    .stTextInput input,
    .stSelectbox div[data-baseweb="select"],
    .stSelectbox div[data-baseweb="select"] > div,
    .stSelectbox div[role="combobox"],
    .stSelectbox div[role="combobox"] > div,
    .stNumberInput input,
    .stTextArea textarea,
    .stMultiSelect div[data-baseweb="select"],
    .stMultiSelect div[data-baseweb="select"] > div {{
        border-color: {input_border} !important;
        background-color: {popover_bg} !important;
        color: {text_color} !important;
    }}
    
    /* Input selected text color and control elements */
    .stSelectbox div[data-baseweb="select"] *, 
    .stSelectbox div[role="combobox"] *,
    .stMultiSelect div[data-baseweb="select"] * {{
        background-color: {popover_bg} !important;
        color: {text_color} !important;
    }}
    
    /* Input field placeholders styling */
    input::placeholder, textarea::placeholder {{
        color: {placeholder_color} !important;
        opacity: 0.75 !important;
    }}
    input::-webkit-input-placeholder, textarea::-webkit-input-placeholder {{
        color: {placeholder_color} !important;
        opacity: 0.75 !important;
    }}
    
    /* baseweb popover dropdown lists styled */
    div[data-baseweb="popover"], div[data-baseweb="menu"], ul[role="listbox"], li[role="option"] {{
        background-color: {popover_bg} !important;
        color: {text_color} !important;
    }}
    
    div[data-baseweb="popover"] ul, div[data-baseweb="popover"] li, div[data-baseweb="popover"] span {{
        background-color: {popover_bg} !important;
        color: {text_color} !important;
    }}
    
    /* baseweb hover highlights */
    div[data-baseweb="popover"] li:hover,
    div[data-baseweb="popover"] li[aria-selected="true"],
    div[data-baseweb="popover"] li:hover * {{
        background-color: #4A90E2 !important;
        color: #ffffff !important;
    }}
    
    /* Details & Expanders */
    .streamlit-expanderHeader, details {{
        background-color: {card_bg} !important;
        color: {text_color} !important;
        border-color: {card_border} !important;
    }}
    .streamlit-expanderContent {{
        background-color: {popover_bg} !important;
        color: {text_color} !important;
        border-color: {card_border} !important;
    }}
    
    /* Responsive Theme Toggle */
</style>
""", unsafe_allow_html=True)

# ----------------------------------------------------
# HELPER FUNCTIONS FOR FILE & WINDOWS DIALOGS
# ----------------------------------------------------

def open_folder_picker(title="Оберіть папку", initialdir=None):
    """Opens a native Windows directory selection dialog."""
    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk()
        root.withdraw()
        root.wm_attributes('-topmost', 1)
        if not initialdir:
            initialdir = st.session_state.get("last_opened_folder") or ""
        if not initialdir or not os.path.exists(initialdir) or not os.path.isdir(initialdir):
            initialdir = os.getcwd()
        folder = filedialog.askdirectory(parent=root, title=title, initialdir=initialdir)
        root.destroy()
        return folder
    except Exception as e:
        st.warning("Не вдалося відкрити діалогове вікно Windows. Будь ласка, введіть шлях вручну.")
        return ""

def open_file_picker(filetypes=None, initialdir=None):
    """Opens a native Windows file selection dialog."""
    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk()
        root.withdraw()
        root.wm_attributes('-topmost', 1)
        if filetypes is None:
            filetypes = [("Усі підтримувані", "*.docx;*.xlsx"), ("Word файли", "*.docx"), ("Excel файли", "*.xlsx"), ("Усі файли", "*.*")]
        if not initialdir:
            initialdir = st.session_state.get("last_opened_folder") or ""
        if not initialdir or not os.path.exists(initialdir) or not os.path.isdir(initialdir):
            initialdir = os.getcwd()
        file = filedialog.askopenfilename(parent=root, title="Оберіть файл-зразок", filetypes=filetypes, initialdir=initialdir)
        root.destroy()
        return file
    except Exception as e:
        st.warning("Не вдалося відкрити діалогове вікно Windows. Будь ласка, введіть шлях вручну.")
        return ""

def create_new_sheet(filepath, new_sheet_name):
    """Creates a new worksheet with default metadata and template placeholders."""
    try:
        wb = openpyxl.load_workbook(filepath)
        if new_sheet_name in wb.sheetnames:
            st.error(f"Аркуш з назвою '{new_sheet_name}' вже існує!")
            return False
        ws = wb.create_sheet(title=new_sheet_name)
        # Initialize default values
        ws['A1'] = "template_placeholder.docx"
        ws['A2'] = "output_{{YYYY}}{{MM}}{{DD}}.docx"
        ws.cell(row=4, column=1).value = "field_1"
        ws.cell(row=5, column=1).value = "значення_1"
        wb.save(filepath)
        return True
    except Exception as e:
        st.error(f"Помилка при створенні аркуша: {e}")
        return False

def rename_sheet(filepath, old_sheet_name, new_sheet_name):
    """Safely renames a sheet in the Excel config."""
    try:
        wb = openpyxl.load_workbook(filepath)
        if old_sheet_name not in wb.sheetnames:
            st.error(f"Аркуш '{old_sheet_name}' не знайдено!")
            return False
        if new_sheet_name in wb.sheetnames:
            st.error(f"Аркуш з назвою '{new_sheet_name}' вже існує!")
            return False
        ws = wb[old_sheet_name]
        ws.title = new_sheet_name
        wb.save(filepath)
        return True
    except Exception as e:
        st.error(f"Помилка при перейменуванні аркуша: {e}")
        return False

def delete_sheet(filepath, sheet_name):
    """Deletes a sheet from the Excel config."""
    try:
        wb = openpyxl.load_workbook(filepath)
        if sheet_name not in wb.sheetnames:
            st.error(f"Аркуш '{sheet_name}' не знайдено!")
            return False
        if len(wb.sheetnames) <= 1:
            st.error("Неможливо видалити єдиний аркуш у книзі!")
            return False
        wb.remove(wb[sheet_name])
        wb.save(filepath)
        return True
    except Exception as e:
        st.error(f"Помилка при видаленні аркуша: {e}")
        return False

def update_config_template_path(filepath, sheet_name, new_template_path):
    """Safely updates only the template path (cell A1) of a specific sheet in Excel config."""
    try:
        # Convert template path to relative path if possible
        config_dir = os.path.dirname(os.path.abspath(filepath))
        try:
            target_abs = os.path.abspath(new_template_path)
            if os.path.splitdrive(target_abs)[0].lower() == os.path.splitdrive(config_dir)[0].lower():
                new_template_path = os.path.relpath(target_abs, config_dir).replace('\\', '/')
        except Exception:
            pass
            
        wb = openpyxl.load_workbook(filepath)
        if sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            ws.cell(row=1, column=1).value = new_template_path
            wb.save(filepath)
            return True
        return False
    except Exception as e:
        st.error(f"Помилка при оновленні шляху шаблону в конфігу: {e}")
        return False

def get_formatted_documentation_markdown():
    """Returns the full technical guide content formatted as premium Markdown, loaded dynamically from _templates_machine_.txt."""
    try:
        filepath = "_templates_machine_.txt"
        if os.path.exists(filepath):
            with open(filepath, "r", encoding="utf-8") as f:
                return f.read()
    except Exception as e:
        st.warning(f"Не вдалося завантажити файл довідки: {e}")
    
    # Fallback to absolute bare-bones in case the file gets deleted
    return "### 📖 Документація\nНе вдалося завантажити файл довідки `_templates_machine_.txt`. Будь ласка, переконайтеся, що файл існує у робочій папці."

# ----------------------------------------------------
# PROJECT MANAGER HELPERS & SCANNERS
# ----------------------------------------------------

def is_excel_config(file_path):
    """Robust algorithm to distinguish config Excel files from regular spreadsheets."""
    if not file_path.endswith('.xlsx') or os.path.basename(file_path).startswith('~$'):
        return False
    try:
        wb = openpyxl.load_workbook(file_path, read_only=True)
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            # Read cells A1 and A2
            a1 = ws.cell(row=1, column=1).value
            a2 = ws.cell(row=2, column=1).value
            if a1 and isinstance(a1, str):
                a1_lower = a1.lower()
                if any(ext in a1_lower for ext in ['.docx', '.xlsx']) or 'template' in a1_lower:
                    return True
            if a2 and isinstance(a2, str):
                a2_lower = a2.lower()
                if any(ext in a2_lower for ext in ['.docx', '.xlsx']) or '{{' in a2_lower:
                    return True
        return False
    except Exception:
        return False

def scan_recursive_configs(root_folder):
    """Walks the folder structure recursively and returns a list of config files."""
    config_files = []
    if not os.path.exists(root_folder) or not os.path.isdir(root_folder):
        return []
    for dirpath, _, filenames in os.walk(root_folder):
        for f in filenames:
            if f.endswith('.xlsx') and not f.startswith('~$'):
                full_path = os.path.abspath(os.path.join(dirpath, f))
                if is_excel_config(full_path):
                    config_files.append(full_path)
    return sorted(config_files)

def build_dir_tree(config_files, root_path):
    """Builds a nested directory tree dictionary from config paths."""
    tree = {}
    for path in config_files:
        rel_path = os.path.relpath(path, root_path)
        parts = rel_path.split(os.sep)
        current = tree
        for part in parts[:-1]:
            if part not in current:
                current[part] = {}
            current = current[part]
        current[parts[-1]] = path
    return tree

def resolve_virtual_doc_name(pattern, row_data, template_path):
    """Resolves output document name using date/time variables and row data."""
    from datetime import datetime
    now = datetime.now()
    now_vars = {
        "YYYY": now.strftime("%Y"),
        "MM": now.strftime("%m"),
        "DD": now.strftime("%d"),
        "hh": now.strftime("%H"),
        "mm": now.strftime("%M"),
        "ss": now.strftime("%S")
    }
    variables = {**now_vars, **row_data}
    
    result = str(pattern)
    for key, val in variables.items():
        pattern_re = r"\{\{\s*" + re.escape(key) + r"\s*\}\}"
        result = re.sub(pattern_re, str(val), result)
        
    if template_path:
        ext = os.path.splitext(template_path)[1].lower()
        if not result.lower().endswith(ext):
            result += ext
    return result

def save_generated_document_dialog(template_path, variables, config_path, name_pattern=None):
    """Generates the document and opens a native Windows dialog to save it, with fallback."""
    import tkinter as tk
    from tkinter import filedialog
    from _templates_machine_ import process_word, process_excel
    import tempfile
    
    from _templates_machine_ import resolve_path
    cfg_dir = os.path.dirname(os.path.abspath(config_path))
    actual_t_path = resolve_path(cfg_dir, template_path)
        
    if not os.path.exists(actual_t_path):
        st.error(f"Шаблон не знайдено: {template_path}")
        return
        
    ext = os.path.splitext(actual_t_path)[1].lower()
    
    if not name_pattern:
        name_pattern = variables.get("name_pattern", "document")
    proposed_filename = os.path.basename(resolve_virtual_doc_name(name_pattern, variables, template_path))
    
    # Attempt native save dialog
    saved = False
    try:
        root = tk.Tk()
        root.withdraw()
        root.wm_attributes('-topmost', 1)
        
        filetypes = [("Word Document", "*.docx")] if ext == ".docx" else [("Excel Workbook", "*.xlsx")]
        default_ext = ext
        
        save_path = filedialog.asksaveasfilename(
            parent=root,
            title="Зберегти згенерований документ",
            filetypes=filetypes,
            defaultextension=default_ext,
            initialfile=proposed_filename
        )
        root.destroy()
        
        if save_path:
            if ext == ".docx":
                process_word(actual_t_path, save_path, variables)
            elif ext == ".xlsx":
                process_excel(actual_t_path, save_path, variables)
            st.success(f"🎉 Документ успішно збережено: {save_path}")
            saved = True
    except Exception as e:
        st.warning(f"Не вдалося відкрити діалогове вікно збереження Windows: {e}. Використовуємо завантаження через браузер.")
        
    if not saved:
        # Fallback to browser download button
        temp_dir = tempfile.gettempdir()
        temp_file = os.path.join(temp_dir, proposed_filename)
        try:
            if ext == ".docx":
                process_word(actual_t_path, temp_file, variables)
            elif ext == ".xlsx":
                process_excel(actual_t_path, temp_file, variables)
                
            with open(temp_file, "rb") as f:
                file_bytes = f.read()
                
            st.download_button(
                label="⬇️ Завантажити згенерований документ через браузер",
                data=file_bytes,
                file_name=proposed_filename,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document" if ext == ".docx" else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="pm_download_fallback_btn"
            )
        except Exception as e_gen:
            st.error(f"Помилка генерації документа: {e_gen}")

# ----------------------------------------------------
# SYSTEM STATE & WORKSPACE SCANNING
# ----------------------------------------------------

def scan_workspace():
    """Scans the directory for configs and templates."""
    files = os.listdir('.')
    configs = [f for f in files if f.endswith('.xlsx') and not f.startswith('~$') and 'template' not in f.lower()]
    templates = [f for f in files if (f.endswith('.docx') or f.endswith('.xlsx')) and f.startswith('template_')]
    
    # Check inside example directory as well
    if os.path.exists('example'):
        for f in os.listdir('example'):
            if f.endswith('.xlsx') and not f.startswith('~$') and 'template' not in f.lower():
                configs.append(os.path.join('example', f))
            if (f.endswith('.docx') or f.endswith('.xlsx')) and f.startswith('template_'):
                templates.append(os.path.join('example', f))
                
    return sorted(configs), sorted(templates)

# ----------------------------------------------------
# EXCEL CONFIG READ/WRITE WITH OPENPYXL
# ----------------------------------------------------

def get_cached_config(config_path):
    """Retrieves config data from cache or loads it if stale/missing."""
    if "pm_cached_configs" not in st.session_state:
        st.session_state["pm_cached_configs"] = {}
        
    try:
        mtime = os.path.getmtime(config_path)
    except Exception:
        mtime = 0
        
    cache_entry = st.session_state["pm_cached_configs"].get(config_path)
    if cache_entry and cache_entry.get("mtime") == mtime:
        return cache_entry["data"]
        
    data = load_excel_config(config_path)
    if data:
        st.session_state["pm_cached_configs"][config_path] = {
            "mtime": mtime,
            "data": data
        }
    return data

def load_excel_config(filepath):
    """Loads all worksheets from Excel config safely parsing metadata (A1 and A2), imported from core."""
    from _templates_machine_ import load_excel_config as core_load
    try:
        return core_load(filepath)
    except Exception as e:
        st.error(f"Помилка при завантаженні конфігу {filepath}: {e}")
        return None

def save_excel_config(filepath, sheet_name, template_path, name_pattern, headers, df_data):
    """Saves changes back to Excel config safely preserving other sheets, imported from core."""
    from _templates_machine_ import save_excel_config as core_save
    try:
        return core_save(filepath, sheet_name, template_path, name_pattern, headers, df_data)
    except Exception as e:
        st.error(f"Помилка при збереженні змін: {e}")
        return False

# ----------------------------------------------------
# LIVE SUBPROCESS LOG STREAMING
# ----------------------------------------------------

def run_subprocess_and_stream(args):
    """Runs the python automation script and streams console outputs to Streamlit."""
    python_exe = sys.executable
    script_path = "_templates_machine_.py"
    
    cmd = [python_exe, script_path] + args
    
    log_area = st.empty()
    progress_bar = st.progress(0.0)
    
    st.info(f"🚀 Запущено команду: `python _templates_machine_.py {' '.join(f'\"{a}\"' for a in args)}`")
    
    # Initialize last operation state
    st.session_state["last_operation_logs"] = []
    st.session_state["last_operation_status"] = "running"
    st.session_state["last_operation_cmd"] = f"python _templates_machine_.py {' '.join(f'\"{a}\"' for a in args)}"
    
    # Force Python stdout to UTF-8 using environment variables
    env = os.environ.copy()
    env["PYTHONIOENCODING"] = "utf-8"
    
    # Initialize process in binary mode to safely handle dynamic decoding fallbacks
    process = subprocess.Popen(
        cmd,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=False,
        env=env
    )
    
    logs = []
    
    while True:
        line_bytes = process.stdout.readline()
        if not line_bytes:
            break
            
        # Bulletproof byte decoding fallback strategy (UTF-8 -> CP1251 -> UTF-8 with replacement)
        line = ""
        try:
            line = line_bytes.decode('utf-8')
        except UnicodeDecodeError:
            try:
                line = line_bytes.decode('cp1251')
            except Exception:
                line = line_bytes.decode('utf-8', errors='replace')
                
        cleaned_line = line.strip()
        if cleaned_line:
            logs.append(cleaned_line)
            st.session_state["last_operation_logs"] = logs
            # Display logs in real-time
            log_area.code("\n".join(logs[-15:])) # Show last 15 lines of progress
            
            # Simple progress heuristics
            if "Створено" in cleaned_line or "Обробка" in cleaned_line:
                progress_bar.progress(0.5)
            elif "Готово!" in cleaned_line or "Успішно!" in cleaned_line:
                progress_bar.progress(1.0)
                
    process.wait()
    
    if process.returncode == 0:
        progress_bar.progress(1.0)
        st.session_state["last_operation_status"] = "success"
        st.success("🎉 Завдання успішно виконано!")
        st.balloons()
    else:
        st.session_state["last_operation_status"] = "error"
        st.error(f"❌ Помилка виконання! Код завершення: {process.returncode}")
        
    return process.returncode, logs

# ----------------------------------------------------
# HELPER FOR PERSISTENT CONSOLE LOGS
# ----------------------------------------------------

def show_last_operation_logs():
    """Displays the persistent logs of the last executed operation from session state."""
    if "last_operation_logs" in st.session_state and st.session_state["last_operation_logs"]:
        status = st.session_state.get("last_operation_status", "")
        cmd = st.session_state.get("last_operation_cmd", "")
        
        st.markdown("---")
        if status == "success":
            st.markdown("### 🟢 Результат останньої операції: Успішно")
        elif status == "error":
            st.markdown("### 🔴 Результат останньої операції: Помилка")
        else:
            st.markdown("### 🟡 Результат останньої операції: Виконується")
            
        if cmd:
            st.caption(f"Команда: `{cmd}`")
            
        with st.expander("📋 Показати повний лог консолі", expanded=True):
            st.code("\n".join(st.session_state["last_operation_logs"]))

def recreate_bat_file(bat_path, config_path):
    """Recreates the execution .bat file at the destination with correct relative paths."""
    try:
        s_dir = os.path.dirname(os.path.abspath(__file__))
        m_py = os.path.join(s_dir, "_templates_machine_.py")
        v_py = os.path.join(s_dir, ".venv", "Scripts", "python.exe")
        if not os.path.exists(v_py):
            v_py = "python"
            
        bat_dir = os.path.dirname(os.path.abspath(bat_path))
        
        def get_rel(target):
            try:
                target_abs = os.path.abspath(target)
                if os.path.splitdrive(target_abs)[0].lower() == os.path.splitdrive(bat_dir)[0].lower():
                    return os.path.relpath(target_abs, bat_dir)
            except Exception:
                pass
            return os.path.abspath(target)
            
        rel_m_py = get_rel(m_py)
        rel_config = get_rel(config_path)
        
        if v_py == "python":
            rel_v_py = "python"
        else:
            rel_v_py = get_rel(v_py)
            
        cnt = f'@echo off\ncd /d "%~dp0"\n"{rel_v_py}" "{rel_m_py}" "{rel_config}" all all\npause\n'
        with open(bat_path, "w", encoding="cp1251") as f:
            f.write(cnt)
    except Exception as e:
        st.error(f"Не вдалося оновити виконавчий .bat файл: {e}")

def move_autopilot_outputs(dest_dir):
    """Moves generated templates and configs of autopilot mode to a custom destination directory."""
    if not dest_dir:
        return
    try:
        import shutil
        dest_dir = os.path.abspath(dest_dir)
        os.makedirs(dest_dir, exist_ok=True)
        # Move Auto_Config.xlsx and Auto_Run_All.bat
        for filename in ["Auto_Config.xlsx", "Auto_Run_All.bat"]:
            src = os.path.join(os.getcwd(), filename)
            if os.path.exists(src):
                shutil.move(src, os.path.join(dest_dir, filename))
        # Move all template_* files from current directory
        moved_count = 0
        for filename in os.listdir(os.getcwd()):
            if filename.startswith("template_") and os.path.isfile(filename):
                shutil.move(filename, os.path.join(dest_dir, filename))
                moved_count += 1
                
        # Recreate the moved BAT file with relative paths inside the destination folder!
        dest_bat = os.path.join(dest_dir, "Auto_Run_All.bat")
        dest_cfg = os.path.join(dest_dir, "Auto_Config.xlsx")
        if os.path.exists(dest_bat) and os.path.exists(dest_cfg):
            recreate_bat_file(dest_bat, dest_cfg)
            
        st.toast(f"✅ Результати аналізу перенесено в: {dest_dir} (переміщено {moved_count} шаблонів)", icon="📁")
    except Exception as e:
        st.error(f"Помилка при перенесенні файлів результатів: {e}")

def move_batch_outputs(sample_path, dest_dir):
    """Moves generated templates and configs of batch mode to a custom destination directory."""
    if not dest_dir or not sample_path:
        return
    try:
        import shutil
        dest_dir = os.path.abspath(dest_dir)
        os.makedirs(dest_dir, exist_ok=True)
        f_dir = os.path.dirname(os.path.abspath(sample_path))
        base_name, ext = os.path.splitext(os.path.basename(sample_path))
        
        files_to_move = [
            f"template_{base_name}{ext}",
            f"{base_name}_config.xlsx",
            f"{base_name}_run_all.bat"
        ]
        moved_count = 0
        for filename in files_to_move:
            src = os.path.join(f_dir, filename)
            if os.path.exists(src):
                shutil.move(src, os.path.join(dest_dir, filename))
                moved_count += 1
                
        # Recreate the moved BAT file with relative paths inside the destination folder!
        dest_bat = os.path.join(dest_dir, f"{base_name}_run_all.bat")
        dest_cfg = os.path.join(dest_dir, f"{base_name}_config.xlsx")
        if os.path.exists(dest_bat) and os.path.exists(dest_cfg):
            recreate_bat_file(dest_bat, dest_cfg)
            
        st.toast(f"✅ Результати пакетного аналізу перенесено в: {dest_dir} (переміщено {moved_count} файлів)", icon="📁")
    except Exception as e:
        st.error(f"Помилка при перенесенні файлів результатів: {e}")

def move_pairwise_outputs(file1_path, dest_dir):
    """Moves generated templates and configs of pairwise mode to a custom destination directory."""
    if not dest_dir or not file1_path:
        return
    try:
        import shutil
        dest_dir = os.path.abspath(dest_dir)
        os.makedirs(dest_dir, exist_ok=True)
        f_dir = os.path.dirname(os.path.abspath(file1_path))
        base_name, ext = os.path.splitext(os.path.basename(file1_path))
        
        files_to_move = [
            f"template_{base_name}{ext}",
            f"{base_name}_config.xlsx",
            f"{base_name}_run_all.bat"
        ]
        moved_count = 0
        for filename in files_to_move:
            src = os.path.join(f_dir, filename)
            if os.path.exists(src):
                shutil.move(src, os.path.join(dest_dir, filename))
                moved_count += 1
                
        # Recreate the moved BAT file with relative paths inside the destination folder!
        dest_bat = os.path.join(dest_dir, f"{base_name}_run_all.bat")
        dest_cfg = os.path.join(dest_dir, f"{base_name}_config.xlsx")
        if os.path.exists(dest_bat) and os.path.exists(dest_cfg):
            recreate_bat_file(dest_bat, dest_cfg)
            
        st.toast(f"✅ Результати попарного порівняння перенесено в: {dest_dir} (переміщено {moved_count} файлів)", icon="📁")
    except Exception as e:
        st.error(f"Помилка при перенесенні файлів результатів: {e}")

# ----------------------------------------------------
# CORE DASHBOARD INTERFACE LAYOUT
# ----------------------------------------------------

# Initialize session state variables for file and folder picker widgets
for key in [
    "txt_auto_folder", "txt_batch_sample", "txt_batch_folder", 
    "txt_pair_file1", "txt_pair_file2", "editor_config_path", 
    "gen_config_path", "gen_output_dir", "analysis_output_dir",
    "last_operation_logs", "last_operation_status", "last_operation_cmd",
    "last_opened_folder", "last_opened_config", "last_opened_template",
    "pm_folder_path"
]:
    if key not in st.session_state:
        if key == "last_operation_logs":
            st.session_state[key] = []
        elif key in ["last_operation_status", "last_operation_cmd"]:
            st.session_state[key] = None if key == "last_operation_status" else ""
        else:
            st.session_state[key] = ""

# Load/Save persistent state to maintain context across restarts
import json
STATE_FILE = ".last_state.json"

def load_persistent_state():
    if os.path.exists(STATE_FILE):
        try:
            with open(STATE_FILE, "r", encoding="utf-8") as f:
                state = json.load(f)
                for k, v in state.items():
                    if v is not None:
                        st.session_state[k] = v
        except Exception:
            pass

def save_persistent_state():
    state = {
        "last_opened_folder": st.session_state.get("last_opened_folder", ""),
        "last_opened_config": st.session_state.get("last_opened_config", ""),
        "last_opened_template": st.session_state.get("last_opened_template", ""),
        "pm_folder_path": st.session_state.get("pm_folder_path", ""),
        "editor_config_path": st.session_state.get("editor_config_path", ""),
        "editor_template_path": st.session_state.get("editor_template_path", "")
    }
    try:
        with open(STATE_FILE, "w", encoding="utf-8") as f:
            json.dump(state, f, ensure_ascii=False, indent=4)
    except Exception:
        pass

load_persistent_state()

def generate_docx_preview(template_path, variables, config_path=None):
    """Generates a temporary Word document and extracts its content for quick preview in high-fidelity HTML."""
    import tempfile
    from docx import Document
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from _templates_machine_ import process_word
    
    # Resolve relative template paths relative to the current Excel config directory if needed
    from _templates_machine_ import resolve_path
    c_path = config_path or st.session_state.get("editor_config_path")
    cfg_dir = os.path.dirname(os.path.abspath(c_path)) if c_path else os.getcwd()
    actual_path = resolve_path(cfg_dir, template_path)
            
    if not os.path.exists(actual_path):
        return f"<div style='color: red; font-weight: bold;'>Шаблон не знайдено за шляхом: {template_path}</div>"
        
    temp_dir = tempfile.gettempdir()
    temp_out = os.path.join(temp_dir, "temp_preview.docx")
    
    try:
        process_word(actual_path, temp_out, variables)
        doc = Document(temp_out)
        
        html = []
        html.append("<div style='font-family: \"Times New Roman\", Times, serif; color: #1a202c; max-width: 800px; margin: 0 auto; line-height: 1.5; font-size: 14px;'>")
        
        # Traverse paragraphs and tables in order
        for element in doc.element.body:
            if element.tag.endswith('p'):
                # It's a paragraph
                paragraph = None
                for p in doc.paragraphs:
                    if p._p == element:
                        paragraph = p
                        break
                if paragraph:
                    style_str = []
                    
                    # 1. Text Alignment
                    align = paragraph.alignment
                    if align == WD_ALIGN_PARAGRAPH.CENTER:
                        style_str.append("text-align: center;")
                    elif align == WD_ALIGN_PARAGRAPH.RIGHT:
                        style_str.append("text-align: right;")
                    elif align == WD_ALIGN_PARAGRAPH.JUSTIFY:
                        style_str.append("text-align: justify;")
                    else:
                        style_str.append("text-align: left;")
                        
                    # 2. Spacing Before / After & Indents
                    space_before = paragraph.paragraph_format.space_before.pt if paragraph.paragraph_format.space_before else 0
                    space_after = paragraph.paragraph_format.space_after.pt if paragraph.paragraph_format.space_after else 6
                    left_indent = paragraph.paragraph_format.left_indent.pt if paragraph.paragraph_format.left_indent else 0
                    right_indent = paragraph.paragraph_format.right_indent.pt if paragraph.paragraph_format.right_indent else 0
                    
                    style_str.append(f"margin: 0; margin-top: {space_before}pt; margin-bottom: {space_after}pt; margin-left: {left_indent}pt; margin-right: {right_indent}pt;")
                    
                    # 3. First Line Indent (paragraph tab spacing)
                    if paragraph.paragraph_format.first_line_indent:
                        fl_val = paragraph.paragraph_format.first_line_indent.pt
                        style_str.append(f"text-indent: {fl_val}pt;")
                        
                    # 4. Line Spacing
                    line_spacing = paragraph.paragraph_format.line_spacing
                    if line_spacing:
                        if isinstance(line_spacing, float):
                            style_str.append(f"line-height: {line_spacing};")
                        else:
                            style_str.append(f"line-height: {line_spacing.pt}pt;")
                    else:
                        style_str.append("line-height: 1.15;")
                        
                    # Write <p> tag
                    html.append(f"<p style='{' '.join(style_str)}'>")
                    for run in paragraph.runs:
                        r_style = []
                        if run.bold:
                            r_style.append("font-weight: bold;")
                        if run.italic:
                            r_style.append("font-style: italic;")
                        if run.underline:
                            r_style.append("text-decoration: underline;")
                        if run.font.size:
                            r_style.append(f"font-size: {run.font.size.pt}pt;")
                        if run.font.name:
                            r_style.append(f"font-family: '{run.font.name}', 'Times New Roman', serif;")
                        if run.font.color and run.font.color.rgb:
                            rgb = run.font.color.rgb
                            r_style.append(f"color: #{rgb[0]:02x}{rgb[1]:02x}{rgb[2]:02x};")
                            
                        text_html = run.text.replace("\n", "<br>")
                        html.append(f"<span style='{' '.join(r_style)}'>{text_html}</span>")
                    
                    # Handle empty paragraphs (preserve spacing)
                    if not paragraph.runs:
                        html.append("&nbsp;")
                        
                    html.append("</p>")
            elif element.tag.endswith('tbl'):
                # It's a table
                table = None
                for t in doc.tables:
                    if t._tbl == element:
                        table = t
                        break
                if table:
                    html.append("<table style='border-collapse: collapse; width: 100%; margin: 15px 0; border: 1px solid #cbd5e1;'>")
                    for row in table.rows:
                        html.append("<tr>")
                        for cell in row.cells:
                            html.append("<td style='border: 1px solid #cbd5e1; padding: 8px; vertical-align: top;'>")
                            for cell_p in cell.paragraphs:
                                cell_p_style = []
                                align = cell_p.alignment
                                if align == WD_ALIGN_PARAGRAPH.CENTER:
                                    cell_p_style.append("text-align: center;")
                                elif align == WD_ALIGN_PARAGRAPH.RIGHT:
                                    cell_p_style.append("text-align: right;")
                                elif align == WD_ALIGN_PARAGRAPH.JUSTIFY:
                                    cell_p_style.append("text-align: justify;")
                                else:
                                    cell_p_style.append("text-align: left;")
                                    
                                # Cell paragraph spacing should be compact
                                space_before = cell_p.paragraph_format.space_before.pt if cell_p.paragraph_format.space_before else 0
                                space_after = cell_p.paragraph_format.space_after.pt if cell_p.paragraph_format.space_after else 2
                                left_indent = cell_p.paragraph_format.left_indent.pt if cell_p.paragraph_format.left_indent else 0
                                right_indent = cell_p.paragraph_format.right_indent.pt if cell_p.paragraph_format.right_indent else 0
                                
                                cell_p_style.append(f"margin: 0; margin-top: {space_before}pt; margin-bottom: {space_after}pt; margin-left: {left_indent}pt; margin-right: {right_indent}pt;")
                                
                                if cell_p.paragraph_format.first_line_indent:
                                    fl_val = cell_p.paragraph_format.first_line_indent.pt
                                    cell_p_style.append(f"text-indent: {fl_val}pt;")
                                    
                                line_spacing = cell_p.paragraph_format.line_spacing
                                if line_spacing:
                                    if isinstance(line_spacing, float):
                                        cell_p_style.append(f"line-height: {line_spacing};")
                                    else:
                                        cell_p_style.append(f"line-height: {line_spacing.pt}pt;")
                                else:
                                    cell_p_style.append("line-height: 1.15;")
                                    
                                html.append(f"<p style='{' '.join(cell_p_style)}'>")
                                for run in cell_p.runs:
                                    r_style = []
                                    if run.bold:
                                        r_style.append("font-weight: bold;")
                                    if run.italic:
                                        r_style.append("font-style: italic;")
                                    if run.underline:
                                        r_style.append("text-decoration: underline;")
                                    if run.font.size:
                                        r_style.append(f"font-size: {run.font.size.pt}pt;")
                                    if run.font.color and run.font.color.rgb:
                                        rgb = run.font.color.rgb
                                        r_style.append(f"color: #{rgb[0]:02x}{rgb[1]:02x}{rgb[2]:02x};")
                                    text_html = run.text.replace("\n", "<br>")
                                    html.append(f"<span style='{' '.join(r_style)}'>{text_html}</span>")
                                if not cell_p.runs:
                                    html.append("&nbsp;")
                                html.append("</p>")
                            html.append("</td>")
                        html.append("</tr>")
                    html.append("</table>")
                    
        html.append("</div>")
        
        if os.path.exists(temp_out):
            try: os.remove(temp_out)
            except Exception: pass
            
        return "\n".join(html)
    except Exception as e:
        return f"<div style='color: red; font-weight: bold;'>Помилка попереднього перегляду Word: {e}</div>"

def generate_xlsx_preview(template_path, variables, config_path=None):
    """Generates a temporary Excel document and extracts its content for quick preview in high-fidelity HTML."""
    import tempfile
    import openpyxl
    from openpyxl.styles import PatternFill
    from _templates_machine_ import process_excel
    import html as py_html
    
    # Resolve relative template paths relative to the current Excel config directory if needed
    from _templates_machine_ import resolve_path
    c_path = config_path or st.session_state.get("editor_config_path")
    cfg_dir = os.path.dirname(os.path.abspath(c_path)) if c_path else os.getcwd()
    actual_path = resolve_path(cfg_dir, template_path)
            
    if not os.path.exists(actual_path):
        return f"<div style='color: red; font-weight: bold;'>Шаблон не знайдено за шляхом: {template_path}</div>"
        
    temp_dir = tempfile.gettempdir()
    temp_out = os.path.join(temp_dir, "temp_preview.xlsx")
    
    try:
        process_excel(actual_path, temp_out, variables)
        # Load with data_only=False to retrieve Excel formula strings instead of blank/None cells
        wb = openpyxl.load_workbook(temp_out, data_only=False)
        
        html = []
        html.append("<div style='font-family: \"Segoe UI\", Tahoma, Geneva, Verdana, sans-serif; color: #1a202c; overflow-x: auto;'>")
        
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            html.append(f"<h4 style='color: #2b6cb0; margin-top: 15px; border-bottom: 2px solid #2b6cb0; padding-bottom: 5px;'>Аркуш: {sheet_name}</h4>")
            html.append("<table style='border-collapse: collapse; border: 1px solid #cbd5e1; font-size: 13px; min-width: 100%;'>")
            
            rows = list(ws.iter_rows())
            if not rows:
                html.append("<tr><td style='padding: 10px; color: #718096;'>Аркуш порожній</td></tr>")
                html.append("</table>")
                continue
                
            for row in rows:
                html.append("<tr>")
                for cell in row:
                    val = cell.value
                    if val is None:
                        val = ""
                        
                    val_str = str(val).strip()
                    if val_str.startswith('='):
                        # Excel formula formatting
                        escaped_formula = py_html.escape(val_str)
                        val_display = f"<span style='color: #2b6cb0; font-family: \"JetBrains Mono\", monospace; font-size: 11px; font-weight: 500; background-color: #ebf8ff; border: 1px solid #bee3f8; border-radius: 4px; padding: 2px 6px; display: inline-block;'>fx {escaped_formula}</span>"
                    else:
                        val_display = py_html.escape(str(val) if val is not None else "")
                        
                    styles = []
                    if cell.alignment:
                        h_align = cell.alignment.horizontal or "left"
                        v_align = cell.alignment.vertical or "center"
                        styles.append(f"text-align: {h_align};")
                        styles.append(f"vertical-align: {v_align};")
                    else:
                        styles.append("text-align: left; vertical-align: center;")
                        
                    if cell.font:
                        if cell.font.bold:
                            styles.append("font-weight: bold;")
                        if cell.font.italic:
                            styles.append("font-style: italic;")
                        if cell.font.size:
                            styles.append(f"font-size: {cell.font.size}pt;")
                        if cell.font.color and cell.font.color.rgb:
                            rgb = str(cell.font.color.rgb)
                            if len(rgb) == 8:
                                rgb = rgb[2:]
                            if rgb != "00000000":
                                styles.append(f"color: #{rgb};")
                                
                    if cell.fill and isinstance(cell.fill, PatternFill) and cell.fill.fill_type == "solid":
                        if cell.fill.fgColor and cell.fill.fgColor.rgb:
                            rgb = str(cell.fill.fgColor.rgb)
                            if len(rgb) == 8:
                                rgb = rgb[2:]
                            if rgb != "00000000":
                                styles.append(f"background-color: #{rgb};")
                                
                    styles_str = " ".join(styles)
                    html.append(f"<td style='border: 1px solid #cbd5e1; padding: 6px 12px; min-width: 80px; {styles_str}'>{val_display}</td>")
                html.append("</tr>")
            html.append("</table>")
            
        html.append("</div>")
        
        if os.path.exists(temp_out):
            try: os.remove(temp_out)
            except Exception: pass
            
        return "\n".join(html)
    except Exception as e:
        return f"<div style='color: red; font-weight: bold;'>Помилка попереднього перегляду Excel: {e}</div>"

def extract_placeholders_with_context(template_path):
    """Extracts jinja2 placeholders like {{var}} from a docx or xlsx template, aggregating all unique contexts."""
    import re
    import openpyxl
    from docx import Document
    
    placeholders_list = {}
    pattern = re.compile(r"\{\{\s*([^{}\s]+)\s*\}\}")
    
    if template_path.endswith(".docx"):
        try:
            doc = Document(template_path)
            full_text = []
            for p in doc.paragraphs:
                full_text.append(p.text)
            for t in doc.tables:
                for row in t.rows:
                    for cell in row.cells:
                        full_text.append(cell.text)
            
            combined_text = "\n".join(full_text)
            matches = pattern.finditer(combined_text)
            for match in matches:
                var_name = match.group(1)
                start_idx = max(0, match.start() - 100)
                end_idx = min(len(combined_text), match.end() + 100)
                context = combined_text[start_idx:end_idx].strip().replace("\n", " ")
                
                if var_name not in placeholders_list:
                    placeholders_list[var_name] = []
                if context not in placeholders_list[var_name]:
                    placeholders_list[var_name].append(context)
        except Exception as e:
            st.error(f"Помилка зчитування Word: {e}")
            
    elif template_path.endswith(".xlsx"):
        try:
            wb = openpyxl.load_workbook(template_path, data_only=False)
            for sheet_name in wb.sheetnames:
                ws = wb[sheet_name]
                for r in ws.iter_rows():
                    for cell in r:
                        val = cell.value
                        if isinstance(val, str):
                            matches = pattern.finditer(val)
                            for match in matches:
                                var_name = match.group(1)
                                start_idx = max(0, match.start() - 100)
                                end_idx = min(len(val), match.end() + 100)
                                context = f"Комірка {cell.coordinate}: {val[start_idx:end_idx].strip()}"
                                
                                if var_name not in placeholders_list:
                                    placeholders_list[var_name] = []
                                if context not in placeholders_list[var_name]:
                                    placeholders_list[var_name].append(context)
        except Exception as e:
            st.error(f"Помилка зчитування Excel: {e}")
            
    # Merge all unique contexts into a single descriptive string per variable
    placeholders = {}
    for var, contexts in placeholders_list.items():
        placeholders[var] = " | ".join(contexts)
        
    return placeholders

def rename_placeholder_in_template(template_path, old_name, new_name):
    """Automatically renames placeholders inside the Word (.docx) or Excel (.xlsx) template file."""
    import os
    if not os.path.exists(template_path):
        return False
    ext = os.path.splitext(template_path)[1].lower()
    if ext == ".docx":
        from docx import Document
        try:
            doc = Document(template_path)
            from _templates_machine_ import consolidate_jinja_tags
            consolidate_jinja_tags(doc)
            old_tag = f"{{{{{old_name}}}}}"
            new_tag = f"{{{{{new_name}}}}}"
            modified = False
            for p in doc.paragraphs:
                if old_tag in p.text:
                    for run in p.runs:
                        if old_tag in run.text:
                            run.text = run.text.replace(old_tag, new_tag)
                            modified = True
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for p in cell.paragraphs:
                            if old_tag in p.text:
                                for run in p.runs:
                                    if old_tag in run.text:
                                        run.text = run.text.replace(old_tag, new_tag)
                                        modified = True
            if modified:
                doc.save(template_path)
                return True
        except Exception:
            pass
    elif ext == ".xlsx":
        import openpyxl
        try:
            wb = openpyxl.load_workbook(template_path)
            old_tag = f"{{{{{old_name}}}}}"
            new_tag = f"{{{{{new_name}}}}}"
            modified = False
            for sheet in wb.worksheets:
                for row in sheet.iter_rows():
                    for cell in row:
                        if cell.value and isinstance(cell.value, str) and old_tag in cell.value:
                            cell.value = cell.value.replace(old_tag, new_tag)
                            modified = True
            if modified:
                wb.save(template_path)
                return True
        except Exception:
            pass
    return False

configs, templates = scan_workspace()

# MAIN PAGE HEADER
col_title, col_theme = st.columns([11, 1])
with col_title:
    st.markdown('<div class="main-title">🚀 Панель керування TemplateMachine</div>', unsafe_allow_html=True)
    st.markdown('<div class="subtitle">Універсальний комбайн для автоматизації документів та аналізу архівів</div>', unsafe_allow_html=True)
with col_theme:
    st.write(" ") # spacer to push down button slightly
    theme_emoji = "🌙" if theme == "light" else "☀️"
    if st.button(theme_emoji, key="theme_toggle_btn", use_container_width=True):
        st.session_state["theme"] = "dark" if theme == "light" else "light"
        st.rerun()

st.markdown("---")

# Initialize session state for active view retention
views_list = [
    "Менеджер Проектів",
    "✈️ Аналіз та Створення Шаблонів",
    "📝 Редактор Excel Конфігів",
    "⚡ Генерація Документів",
    "📖 Повна Довідка"
]

if "app_view_selector" not in st.session_state or st.session_state["app_view_selector"] not in views_list:
    if "current_view" in st.session_state and st.session_state["current_view"] in views_list:
        st.session_state["app_view_selector"] = st.session_state["current_view"]
    else:
        st.session_state["app_view_selector"] = views_list[0]

# Determine default index from session state to avoid key parameter conflicts in Streamlit selectbox
try:
    default_idx = views_list.index(st.session_state["app_view_selector"])
except Exception:
    default_idx = 0

selected_view = st.selectbox(
    "Оберіть розділ роботи:",
    views_list,
    index=default_idx
)
st.session_state["app_view_selector"] = selected_view
st.session_state["current_view"] = selected_view

# Track view transitions
prev_view = st.session_state.get("prev_view")
if prev_view is None:
    st.session_state["prev_view"] = selected_view
    prev_view = selected_view

if prev_view != selected_view:
    # Transitioning
    if prev_view == "📝 Редактор Excel Конфігів":
        # Leaving Config Editor: Save state
        if st.session_state.get("editor_config_path"):
            st.session_state["last_opened_config"] = st.session_state["editor_config_path"]
            st.session_state["last_opened_folder"] = os.path.dirname(os.path.abspath(st.session_state["editor_config_path"]))
            st.session_state["pm_folder_path"] = st.session_state["last_opened_folder"]
        if st.session_state.get("editor_template_path"):
            st.session_state["last_opened_template"] = st.session_state["editor_template_path"]
        save_persistent_state()
        
    elif selected_view == "📝 Редактор Excel Конфігів":
        # Entering Config Editor: Restore state and force reload
        if st.session_state.get("last_opened_config"):
            st.session_state["editor_config_path"] = st.session_state["last_opened_config"]
        if st.session_state.get("last_opened_template"):
            st.session_state["editor_template_path"] = st.session_state["last_opened_template"]
        # Force reload config sheet data to get latest values
        st.session_state["loaded_config_sheet"] = ""
        save_persistent_state()
        
    elif selected_view == "Менеджер Проектів":
        # Entering Project Manager
        if st.session_state.get("last_opened_folder"):
            st.session_state["pm_folder_path"] = st.session_state["last_opened_folder"]
        save_persistent_state()
        
    elif selected_view == "⚡ Генерація Документів":
        # Entering Document Generation
        if st.session_state.get("last_opened_config"):
            st.session_state["gen_config_path"] = st.session_state["last_opened_config"]
        save_persistent_state()
        
    st.session_state["prev_view"] = selected_view

st.markdown(" ")

# ----------------------------------------------------
# VIEW 0: PROJECT MANAGER & DOCUMENT TREE
# ----------------------------------------------------
if selected_view == "Менеджер Проектів":
    st.header("📁 Менеджер Проектів та Віртуальне Дерево Документів")
    st.write("Вкажіть шлях до папки, щоб автоматично просканувати всі конфіги та побудувати дерево віртуальних документів.")
    
    # 1. Folder picker / text input
    if "pm_folder_path" not in st.session_state:
        st.session_state["pm_folder_path"] = ""
        
    col_p1, col_p2 = st.columns([3, 1])
    with col_p1:
        folder_input = st.text_input(
            "📁 Шлях до папки з конфігами та шаблонами:",
            placeholder="Введіть або оберіть шлях до папки (наприклад, example)...",
            key="pm_folder_path"
        )
    with col_p2:
        st.write(" ")
        st.write(" ")
        def select_pm_folder():
            selected = open_folder_picker("Оберіть папку з конфігами")
            if selected:
                st.session_state["pm_folder_path"] = selected
        st.button("📁 Обрати папку", key="btn_pm_folder", on_click=select_pm_folder)
        
    pm_path = st.session_state["pm_folder_path"]
    if pm_path and pm_path != st.session_state.get("last_opened_folder"):
        st.session_state["last_opened_folder"] = pm_path
        save_persistent_state()
    
    if not pm_path:
        st.info("Будь ласка, оберіть або введіть шлях до папки для сканування.")
    elif not os.path.exists(pm_path):
        st.error(f"Вказаний шлях '{pm_path}' не існує!")
    else:
        # Scan and build tree
        with st.spinner("⏳ Сканування папки..."):
            config_files = scan_recursive_configs(pm_path)
            
        if not config_files:
            st.warning("У вказаній папці не знайдено жодного Excel-конфігу.")
        else:
            # Build tree dictionary
            dir_tree = build_dir_tree(config_files, pm_path)
            
            # Setup split layout: Tree on the left, Preview & Edit on the right
            col_tree, col_content = st.columns([5, 7])
            
            with col_tree:
                st.subheader("🌳 Віртуальне дерево документів")
                
                # Recursive tree renderer
                def render_tree_node(node_name, node_value, depth=0):
                    if isinstance(node_value, dict):
                        # Folder node
                        with st.expander("📁 " + node_name, expanded=(depth == 0)):
                            for name, val in node_value.items():
                                render_tree_node(name, val, depth + 1)
                    else:
                        # Config file node
                        config_path = node_value
                        config_name = node_name
                        with st.expander("📊 " + config_name, expanded=False):
                            sheets_data = get_cached_config(config_path)
                            if not sheets_data:
                                st.caption("Не вдалося завантажити або порожній конфіг")
                                return
                            
                            for sheet_name, info in sheets_data.items():
                                rows = info["rows"]
                                template_path = info["template_path"]
                                name_pattern = info["name_pattern"]
                                
                                with st.expander(f"📋 {sheet_name} ({len(rows)} док.)", expanded=False):
                                    if not rows:
                                        st.caption("Немає даних для генерації")
                                        continue
                                    
                                    for idx, row in enumerate(rows):
                                        doc_name = resolve_virtual_doc_name(name_pattern, row, template_path)
                                        if not doc_name.strip():
                                            doc_name = f"document_{idx + 5}"
                                            
                                        is_selected = (
                                            st.session_state.get("pm_selected_doc", {}).get("config_path") == config_path and
                                            st.session_state.get("pm_selected_doc", {}).get("sheet_name") == sheet_name and
                                            st.session_state.get("pm_selected_doc", {}).get("row_idx") == idx
                                        )
                                        
                                        button_type = "primary" if is_selected else "secondary"
                                        
                                        if st.button(
                                            f"📄 {doc_name}",
                                            key=f"pm_btn_{config_path}_{sheet_name}_{idx}",
                                            use_container_width=True,
                                            type=button_type
                                        ):
                                            st.session_state["pm_selected_doc"] = {
                                                "config_path": config_path,
                                                "sheet_name": sheet_name,
                                                "row_idx": idx,
                                                "doc_name": doc_name,
                                                "template_path": template_path,
                                                "name_pattern": name_pattern
                                            }
                                            st.session_state["last_opened_config"] = config_path
                                            st.session_state["last_opened_folder"] = os.path.dirname(os.path.abspath(config_path))
                                            st.session_state["last_opened_template"] = template_path
                                            save_persistent_state()
                                            if "pm_editing_vars" in st.session_state:
                                                del st.session_state["pm_editing_vars"]
                                            st.rerun()

                # Render tree roots
                for name, val in dir_tree.items():
                    render_tree_node(name, val, 0)
                    
            with col_content:
                selected_doc = st.session_state.get("pm_selected_doc")
                if not selected_doc:
                    st.info("👈 Оберіть віртуальний документ у дереві ліворуч, щоб почати перегляд та редагування.")
                else:
                    config_path = selected_doc["config_path"]
                    sheet_name = selected_doc["sheet_name"]
                    row_idx = selected_doc["row_idx"]
                    template_path = selected_doc["template_path"]
                    name_pattern = selected_doc["name_pattern"]
                    
                    st.subheader("🛠️ Перегляд та редагування")
                    
                    # Display metadata paths
                    st.caption(f"**Конфіг:** `{os.path.basename(config_path)}` | **Аркуш:** `{sheet_name}` | **Шаблон:** `{template_path}`")
                    
                    sheets_data = get_cached_config(config_path)
                    if not sheets_data or sheet_name not in sheets_data:
                        st.error("Помилка завантаження даних документа.")
                    else:
                        sheet_info = sheets_data[sheet_name]
                        headers = sheet_info["headers"]
                        row_vars = sheet_info["rows"][row_idx]
                        
                        # Initialize editing variables
                        if "pm_editing_vars" not in st.session_state:
                            st.session_state["pm_editing_vars"] = dict(row_vars)
                            
                        edited_vars = st.session_state["pm_editing_vars"]
                        
                        # Show dynamic output name preview
                        current_resolved_name = resolve_virtual_doc_name(name_pattern, edited_vars, template_path)
                        st.markdown(f"##### 📄 Вихідне ім'я: `{current_resolved_name}`")
                        
                        # Render editing inputs
                        st.markdown("##### ✏️ Змінні цього документа:")
                        var_cols = st.columns(3)
                        for idx, h in enumerate(headers):
                            if not h:
                                continue
                            with var_cols[idx % 3]:
                                val = edited_vars.get(h, "")
                                new_val = st.text_input(f"{h}", value=val, key=f"pm_input_{config_path}_{sheet_name}_{row_idx}_{h}")
                                if new_val != val:
                                    edited_vars[h] = new_val
                                    st.rerun()
                                    
                        # Buttons row
                        st.markdown(" ")
                        btn_col1, btn_col2, btn_col3 = st.columns(3)
                        
                        with btn_col1:
                            if st.button("💾 Зберегти зміни в конфіг", type="primary", use_container_width=True):
                                # Load fresh config data to write
                                full_data = load_excel_config(config_path)
                                if full_data and sheet_name in full_data:
                                    # Update target row
                                    full_data[sheet_name]["rows"][row_idx] = edited_vars
                                    
                                    # Save to file
                                    success = save_excel_config(
                                        config_path,
                                        sheet_name,
                                        full_data[sheet_name]["template_path"],
                                        full_data[sheet_name]["name_pattern"],
                                        full_data[sheet_name]["headers"],
                                        full_data[sheet_name]["rows"]
                                    )
                                    if success:
                                        # Clear cache for this config
                                        if config_path in st.session_state["pm_cached_configs"]:
                                            del st.session_state["pm_cached_configs"][config_path]
                                        st.toast("🎉 Зміни успішно збережено в Excel!", icon="💾")
                                        time.sleep(0.5)
                                        st.rerun()
                                        
                        with btn_col2:
                            if st.button("📄 Зберегти документ", use_container_width=True):
                                # Resolve absolute template path
                                from _templates_machine_ import resolve_path
                                cfg_dir = os.path.dirname(os.path.abspath(config_path))
                                actual_t_path = resolve_path(cfg_dir, template_path)
                                    
                                save_generated_document_dialog(actual_t_path, edited_vars, config_path, name_pattern=name_pattern)
                                
                        with btn_col3:
                            def go_to_config_editor():
                                st.session_state["current_view"] = "📝 Редактор Excel Конфігів"
                                st.session_state["app_view_selector"] = "📝 Редактор Excel Конфігів"
                                st.session_state["editor_config_path"] = config_path
                                st.session_state["editor_selected_sheet"] = sheet_name
                                st.session_state["widget_editor_sheet"] = sheet_name
                                st.session_state["loaded_config_sheet"] = ""  # Force reload headers and data
                                # Update last opened values immediately on transition
                                st.session_state["last_opened_config"] = config_path
                                st.session_state["last_opened_folder"] = os.path.dirname(os.path.abspath(config_path))
                                if template_path:
                                    st.session_state["last_opened_template"] = template_path
                                save_persistent_state()
                            st.button("✏️ Перейти до редактора конфігів", use_container_width=True, on_click=go_to_config_editor)
                                
                        # High fidelity preview
                        st.markdown(" ")
                        st.markdown("##### 🔍 Попередній перегляд документа")
                        
                        # Resolve relative template path
                        from _templates_machine_ import resolve_path
                        cfg_dir = os.path.dirname(os.path.abspath(config_path))
                        actual_template_path = resolve_path(cfg_dir, template_path)
                            
                        if not os.path.exists(actual_template_path):
                            st.warning(f"Шаблон не знайдено за шляхом: {template_path} (враховуючи відносність до конфігу)")
                        else:
                            ext = os.path.splitext(actual_template_path)[1].lower()
                            with st.spinner("⏳ Генерація прев'ю..."):
                                if ext == ".docx":
                                    preview_html = generate_docx_preview(actual_template_path, edited_vars, config_path=config_path)
                                elif ext == ".xlsx":
                                    preview_html = generate_xlsx_preview(actual_template_path, edited_vars, config_path=config_path)
                                else:
                                    preview_html = f"<div style='color: #e53e3e;'>Непідтримуваний тип шаблону: {ext}</div>"
                                    
                            st.markdown(
                                f"""
                                <div style="border: 2px solid #3182ce; border-radius: 8px; padding: 25px; background-color: #ffffff; max-height: 500px; overflow-y: auto; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1); border-left: 8px solid #3182ce;">
                                    {preview_html}
                                </div>
                                """,
                                unsafe_allow_html=True
                            )

# ----------------------------------------------------
# VIEW 1: ARCHIVE ANALYSIS & TEMPLATE CREATION
# ----------------------------------------------------
if selected_view == "✈️ Аналіз та Створення Шаблонів":
    st.header("🔍 Режими аналізу та розпізнавання шаблонів")
    st.write("Скрипт проведе інтелектуальне попарне або групове порівняння документів, виділить змінні і створить конфігураційний файл.")
    
    # Native Streamlit click callbacks to update widget states safely before instantiation
    def select_auto_folder():
        selected = open_folder_picker("Оберіть папку з архівом для аналізу")
        if selected:
            st.session_state["txt_auto_folder"] = selected
            
    def select_batch_sample():
        selected = open_file_picker()
        if selected:
            st.session_state["txt_batch_sample"] = selected
            
    def select_batch_folder():
        selected = open_folder_picker("Оберіть папку для пакетного аналізу")
        if selected:
            st.session_state["txt_batch_folder"] = selected
            
    def select_pair_file1():
        selected = open_file_picker()
        if selected:
            st.session_state["txt_pair_file1"] = selected
            
    def select_pair_file2():
        selected = open_file_picker()
        if selected:
            st.session_state["txt_pair_file2"] = selected
    
    mode = st.radio(
        "Оберіть режим аналізу:",
        [
            "✈️ Повний автопілот (Сканування папки та автоматичне розбиття на групи)",
            "🔍 Пакетний аналіз (Один файл-зразок + папка з іншими файлами)",
            "⚖️ Попарне порівняння (Точне порівняння двох конкретних файлів)"
        ],
        index=0
    )
    
    # --- ALWAYS-VISIBLE ANALYSIS OUTPUT DIRECTORY SELECTION ---
    st.markdown("##### 📁 Папка для збереження результатів (Необов'язково)")
    st.caption("Якщо вказано, створені шаблони (`template_*`), Excel-конфіги та `.bat` файли будуть перенесені в цю папку. У безвіконному (headless) режимі введіть шлях вручну.")
    
    col_ao1, col_ao2 = st.columns([3, 1])
    with col_ao1:
        a_out_dir = st.text_input(
            "📁 Шлях до папки результатів:",
            placeholder="Наприклад: C:/MyTemplates (залиште порожнім для збереження за замовчуванням)",
            key="analysis_output_dir"
        )
    with col_ao2:
        st.write(" ")
        st.write(" ")
        def select_analysis_output_dir():
            selected = open_folder_picker("Оберіть папку для збереження результатів")
            if selected:
                st.session_state["analysis_output_dir"] = selected
        st.button("📁 Обрати папку", key="btn_analysis_output_dir", on_click=select_analysis_output_dir)
        
    st.markdown("---")
    
    if "Повний автопілот" in mode:
        st.subheader("✈️ Режим 1: Повний автопілот")
        st.write("Аналізує вказану папку, групує схожі за структурою документи, створює для кожної групи окремий шаблон та зводить все в єдиний мульти-конфіг `Auto_Config.xlsx`.")
        
        col1, col2 = st.columns([3, 1])
        with col1:
            folder_path = st.text_input(
                "Шлях до папки з архівом документів:",
                placeholder="Наприклад: example/docs",
                key="txt_auto_folder"
            )
        with col2:
            st.write(" ")
            st.write(" ")
            st.button("📁 Провідник Windows", key="btn_auto_folder", on_click=select_auto_folder)
                    
        if st.button("🚀 Запустити повний автопілот", type="primary"):
            if not folder_path:
                st.error("Будь ласка, вкажіть шлях до папки!")
            elif not os.path.exists(folder_path):
                st.error(f"Вказана папка '{folder_path}' не існує!")
            else:
                ret_code, _ = run_subprocess_and_stream([folder_path])
                if ret_code == 0 and st.session_state.get("analysis_output_dir"):
                    move_autopilot_outputs(st.session_state["analysis_output_dir"])
                configs, templates = scan_workspace()
                st.rerun()
                
    elif "Пакетний аналіз" in mode:
        st.subheader("🔍 Режим 2: Пакетний аналіз")
        st.write("Порівнює один обраний файл-зразок зі схожими файлами у папці та формує шаблон і конфігурацію.")
        
        # Row for Sample file
        col1, col2 = st.columns([3, 1])
        with col1:
            sample_file = st.text_input(
                "Шлях до файлу-зразка (.docx або .xlsx):",
                placeholder="Наприклад: example/w.docx",
                key="txt_batch_sample"
            )
        with col2:
            st.write(" ")
            st.write(" ")
            st.button("📄 Обрати зразок", key="btn_batch_sample", on_click=select_batch_sample)
                    
        # Row for Folder
        col1, col2 = st.columns([3, 1])
        with col1:
            folder_path = st.text_input(
                "Шлях до папки порівняння:",
                placeholder="Наприклад: example",
                key="txt_batch_folder"
            )
        with col2:
            st.write(" ")
            st.write(" ")
            st.button("📁 Обрати папку", key="btn_batch_folder", on_click=select_batch_folder)
                    
        if st.button("🚀 Запустити пакетний аналіз", type="primary"):
            if not sample_file or not folder_path:
                st.error("Будь ласка, вкажіть і файл-зразок, і папку порівняння!")
            elif not os.path.exists(sample_file):
                st.error(f"Файл-зразок '{sample_file}' не знайдено!")
            elif not os.path.exists(folder_path):
                st.error(f"Папку порівняння '{folder_path}' не знайдено!")
            else:
                ret_code, _ = run_subprocess_and_stream([sample_file, folder_path])
                if ret_code == 0 and st.session_state.get("analysis_output_dir"):
                    move_batch_outputs(sample_file, st.session_state["analysis_output_dir"])
                configs, templates = scan_workspace()
                st.rerun()
                
    elif "Попарне порівняння" in mode:
        st.subheader("⚖️ Режим 3: Попарне порівняння")
        st.write("Знаходить розбіжності між двома файлами, виділяє змінні та створює шаблон та індивідуальний конфіг.")
        
        # File 1
        col1, col2 = st.columns([3, 1])
        with col1:
            file_1 = st.text_input(
                "Шлях до Першого файлу:",
                placeholder="Наприклад: example/w1.docx",
                key="txt_pair_file1"
            )
        with col2:
            st.write(" ")
            st.write(" ")
            st.button("📄 Обрати файл 1", key="btn_pair_file1", on_click=select_pair_file1)
                    
        # File 2
        col1, col2 = st.columns([3, 1])
        with col1:
            file_2 = st.text_input(
                "Шлях до Другого файлу:",
                placeholder="Наприклад: example/w2.docx",
                key="txt_pair_file2"
            )
        with col2:
            st.write(" ")
            st.write(" ")
            st.button("📄 Обрати файл 2", key="btn_pair_file2", on_click=select_pair_file2)
                    
        if st.button("🚀 Запустити попарне порівняння", type="primary"):
            if not file_1 or not file_2:
                st.error("Будь ласка, вкажіть обидва файли для порівняння!")
            elif not os.path.exists(file_1) or not os.path.exists(file_2):
                st.error("Один або обидва вказані файли не знайдені!")
            else:
                ret_code, _ = run_subprocess_and_stream([file_1, file_2])
                if ret_code == 0 and st.session_state.get("analysis_output_dir"):
                    move_pairwise_outputs(file_1, st.session_state["analysis_output_dir"])
                configs, templates = scan_workspace()
                st.rerun()

    # Persistent log viewer at the bottom of analysis tab
    show_last_operation_logs()

# ----------------------------------------------------
# VIEW 2: INTERACTIVE EXCEL CONFIG EDITOR
# ----------------------------------------------------
elif selected_view == "📝 Редактор Excel Конфігів":
    st.header("📝 Інтерактивний редактор файлів конфігурації")
    st.write("Оберіть будь-який створений конфіг-файл, щоб відредагувати параметри шаблону, імені або змінити дані змінних та аркушів.")
    
    col_c1, col_c2 = st.columns([3, 1])
    with col_c1:
        selected_config = st.text_input(
            "Шлях до файлу конфігурації для редагування:",
            placeholder="Оберіть Excel-файл конфігурації...",
            key="editor_config_path"
        )
    with col_c2:
        st.write(" ")
        st.write(" ")
        def select_editor_config():
            selected = open_file_picker(filetypes=[("Excel конфігурації", "*.xlsx"), ("Усі файли", "*.*")])
            if selected:
                st.session_state["editor_config_path"] = selected
        st.button("📁 Обрати конфіг", key="btn_editor_config", on_click=select_editor_config)
        
    cfg_path = st.session_state["editor_config_path"]
    if cfg_path and (cfg_path != st.session_state.get("last_opened_config") or os.path.dirname(os.path.abspath(cfg_path)) != st.session_state.get("last_opened_folder")):
        st.session_state["last_opened_config"] = cfg_path
        st.session_state["last_opened_folder"] = os.path.dirname(os.path.abspath(cfg_path))
        st.session_state["pm_folder_path"] = st.session_state["last_opened_folder"]
        save_persistent_state()
    if not cfg_path:
        st.info("Будь ласка, оберіть файл конфігурації Excel (`*.xlsx`) для початку роботи!")
    elif not os.path.exists(cfg_path):
        st.error(f"Вказаний файл конфігурації '{cfg_path}' не знайдено!")
    else:
        sheets_data = load_excel_config(cfg_path)
        
        if sheets_data:
            sheet_names = list(sheets_data.keys())
            
            if "editor_selected_sheet" not in st.session_state or st.session_state["editor_selected_sheet"] not in sheet_names:
                st.session_state["editor_selected_sheet"] = sheet_names[0] if sheet_names else ""
                
            selected_sheet = st.selectbox(
                "Оберіть аркуш конфігурації:",
                sheet_names,
                index=sheet_names.index(st.session_state["editor_selected_sheet"]) if st.session_state["editor_selected_sheet"] in sheet_names else 0,
                key="widget_editor_sheet"
            )
            
            if selected_sheet != st.session_state["editor_selected_sheet"]:
                st.session_state["editor_selected_sheet"] = selected_sheet
                
            # --- SHEET CRUD MANAGEMENT ---
            with st.expander("📂 Керування аркушами (Sheets Operations)", expanded=False):
                col_s1, col_s2, col_s3 = st.columns(3)
                
                with col_s1:
                    st.markdown("##### ➕ Створити новий аркуш")
                    new_s_name = st.text_input("Назва нового аркуша:", key="txt_new_sheet_name")
                    if st.button("➕ Створити аркуш", key="btn_create_sheet"):
                        if not new_s_name:
                            st.error("Введіть назву аркуша!")
                        else:
                            clean_s_name = re.sub(r'[\\/*?:\[\]]', "", new_s_name)[:31].strip()
                            if create_new_sheet(cfg_path, clean_s_name):
                                st.session_state["editor_selected_sheet"] = clean_s_name
                                st.success(f"Аркуш '{clean_s_name}' успішно створено!")
                                time.sleep(1)
                                st.rerun()
                                
                with col_s2:
                    st.markdown("##### ✏️ Перейменувати поточний аркуш")
                    new_rename_name = st.text_input("Нова назва аркуша:", key="txt_rename_sheet_name")
                    if st.button("✏️ Перейменувати аркуш", key="btn_rename_sheet"):
                        if not new_rename_name:
                            st.error("Введіть нову назву!")
                        else:
                            clean_rename_name = re.sub(r'[\\/*?:\[\]]', "", new_rename_name)[:31].strip()
                            if rename_sheet(cfg_path, selected_sheet, clean_rename_name):
                                st.session_state["editor_selected_sheet"] = clean_rename_name
                                st.success(f"Аркуш перейменовано на '{clean_rename_name}'!")
                                time.sleep(1)
                                st.rerun()
                                
                with col_s3:
                    st.markdown("##### ❌ Видалити поточний аркуш")
                    st.warning("Ця дія є незворотною!")
                    confirm_delete = st.checkbox("Підтверджую видалення аркуша", key="chk_confirm_delete_sheet")
                    if st.button("❌ Видалити аркуш", key="btn_delete_sheet"):
                        if not confirm_delete:
                            st.error("Будь ласка, підтвердіть видалення чекбоксом!")
                        else:
                            if delete_sheet(cfg_path, selected_sheet):
                                st.success(f"Аркуш '{selected_sheet}' видалено!")
                                remaining_sheets = [s for s in sheet_names if s != selected_sheet]
                                st.session_state["editor_selected_sheet"] = remaining_sheets[0] if remaining_sheets else ""
                                time.sleep(1)
                                st.rerun()
                                
            # --- LOAD SHEET DATA ---
            sheet_info = sheets_data[selected_sheet]
            config_sheet_key = f"{cfg_path}_{selected_sheet}"
            
            # State synchronization on config/sheet change
            if st.session_state.get("loaded_config_sheet") != config_sheet_key:
                st.session_state["loaded_config_sheet"] = config_sheet_key
                st.session_state["current_sheet_headers"] = list(sheet_info["headers"])
                st.session_state["current_sheet_data"] = list(sheet_info["rows"])
                st.session_state["editor_template_path"] = sheet_info["template_path"]
                st.session_state["editor_name_pattern"] = sheet_info["name_pattern"]
                if sheet_info["template_path"]:
                    st.session_state["last_opened_template"] = sheet_info["template_path"]
                save_persistent_state()
                
            # --- METADATA (A1 & A2) ---
            st.markdown("### ⚙️ Налаштування генерації для обраного аркуша")
            col1, col2 = st.columns(2)
            
            with col1:
                col_t1, col_t2 = st.columns([3, 1])
                with col_t1:
                    t_path = st.text_input(
                        "📄 Шлях до шаблону (комірка A1):",
                        placeholder="Оберіть файл шаблону...",
                        key="editor_template_path"
                    )
                with col_t2:
                    st.write(" ")
                    st.write(" ")
                    def select_editor_template():
                        selected = open_file_picker(filetypes=[
                            ("Усі шаблони", "*.docx;*.xlsx"),
                            ("Word шаблони", "*.docx"),
                            ("Excel шаблони", "*.xlsx"),
                            ("Усі файли", "*.*")
                        ])
                        if selected:
                            st.session_state["editor_template_path"] = selected
                    st.button("📁 Шаблон", key="btn_editor_template", on_click=select_editor_template)
                    
            with col2:
                n_pattern = st.text_input(
                    "✍️ Шаблон імені вихідних файлів (комірка A2):",
                    placeholder="Наприклад: output_{{YYYY}}.docx",
                    key="editor_name_pattern"
                )
                
            # Sync template path to last_opened_template immediately if changed
            curr_editor_t_path = st.session_state.get("editor_template_path", "")
            if curr_editor_t_path and curr_editor_t_path != st.session_state.get("last_opened_template", ""):
                st.session_state["last_opened_template"] = curr_editor_t_path
                save_persistent_state()
                
            # --- SLEEK ROW & COLUMN CONTROLS ---
            st.markdown("### 🛠️ Швидкі дії з рядками та стовпчиками (змінними)")
            col_ctrl1, col_ctrl2, col_ctrl3, col_ctrl4 = st.columns(4)
            
            with col_ctrl1:
                with st.popover("➕ Додати стовпчик (змінну)", width="stretch"):
                    new_col = st.text_input("Ім'я нової змінної (напр., client_name):", key="compact_new_col_input")
                    if st.button("➕ Додати стовпчик", key="btn_compact_add_col", width="stretch"):
                        if not new_col:
                            st.error("Введіть ім'я змінної!")
                        elif new_col in st.session_state["current_sheet_headers"]:
                            st.error("Такий стовпчик вже існує!")
                        else:
                            st.session_state["current_sheet_headers"].append(new_col)
                            for r in st.session_state["current_sheet_data"]:
                                r[new_col] = ""
                            st.success(f"Стовпчик '{new_col}' додано!")
                            st.rerun()
                            
            with col_ctrl2:
                with st.popover("❌ Видалити стовпчик", width="stretch"):
                    if st.session_state["current_sheet_headers"]:
                        col_to_del = st.selectbox("Оберіть стовпчик для видалення:", st.session_state["current_sheet_headers"], key="compact_del_col_select")
                        confirm_col = st.checkbox("Підтверджую видалення стовпчика", key="compact_confirm_del_col")
                        if st.button("❌ Видалити стовпчик", key="btn_compact_del_col", type="primary", width="stretch"):
                            if not confirm_col:
                                st.error("Будь ласка, підтвердіть видалення!")
                            else:
                                st.session_state["current_sheet_headers"].remove(col_to_del)
                                for r in st.session_state["current_sheet_data"]:
                                    if col_to_del in r:
                                        r.pop(col_to_del)
                                st.success(f"Стовпчик '{col_to_del}' видалено!")
                                st.rerun()
                    else:
                        st.caption("Немає активних стовпчиків.")
                        
            with col_ctrl3:
                if st.button("➕ Додати рядок", key="btn_compact_add_row", width="stretch"):
                    new_row = {h: "" for h in st.session_state["current_sheet_headers"]}
                    st.session_state["current_sheet_data"].append(new_row)
                    st.success("Рядок додано!")
                    st.rerun()
                    
            with col_ctrl4:
                with st.popover("❌ Видалити рядок", width="stretch"):
                    if st.session_state["current_sheet_data"]:
                        row_to_del = st.number_input("Номер рядка для видалення (з 1):", min_value=1, max_value=len(st.session_state["current_sheet_data"]), step=1, key="compact_del_row_num")
                        confirm_row = st.checkbox("Підтверджую видалення рядка", key="compact_confirm_del_row")
                        if st.button("❌ Видалити рядок", key="btn_compact_del_row", type="primary", width="stretch"):
                            if not confirm_row:
                                st.error("Будь ласка, підтвердіть видалення!")
                            else:
                                idx = int(row_to_del) - 1
                                st.session_state["current_sheet_data"].pop(idx)
                                st.success(f"Рядок {row_to_del} видалено!")
                                st.rerun()
                    else:
                        st.caption("Немає доступних рядків.")
                        
            # --- DATA EDITOR AND SAVE ---
            st.markdown("### 📊 Дані рядків для генерації документів (починаючи з рядка 5)")
            st.caption("Оберіть рядок для перегляду за допомогою прапорця в колонці **«Перегляд»** ліворуч. Клікніть двічі на будь-завгодно клітинку для редагування.")
            
            # Synchronize changes from Streamlit's data editor to session state
            if "config_data_editor" in st.session_state:
                editor_state = st.session_state["config_data_editor"]
                
                # 1. Sync edited cells
                edited_rows = editor_state.get("edited_rows", {})
                for row_idx_str, edits in edited_rows.items():
                    row_idx = int(row_idx_str)
                    if row_idx < len(st.session_state["current_sheet_data"]):
                        for col, val in edits.items():
                            if col == "Перегляд":
                                if val is True:
                                    st.session_state["last_preview_row_idx"] = row_idx
                                elif val is False and st.session_state.get("last_preview_row_idx") == row_idx:
                                    st.session_state["last_preview_row_idx"] = None
                            elif col in st.session_state["current_sheet_headers"]:
                                st.session_state["current_sheet_data"][row_idx][col] = str(val)
                                
                # 2. Sync added rows
                added_rows = editor_state.get("added_rows", [])
                for row in added_rows:
                    new_row = {h: str(row.get(h, "")) for h in st.session_state["current_sheet_headers"]}
                    st.session_state["current_sheet_data"].append(new_row)
                    
                # 3. Sync deleted rows
                deleted_rows = editor_state.get("deleted_rows", [])
                for row_idx in sorted(deleted_rows, reverse=True):
                    if row_idx < len(st.session_state["current_sheet_data"]):
                        st.session_state["current_sheet_data"].pop(row_idx)
            
            # --- INTERACTIVE COLUMN HEADERS EDITOR (EDITABLE BY DOUBLE-CLICK) ---
            st.markdown("##### ✏️ Редагувати назви змінних (подвійний клік на осередки правої колонки):")
            headers_list = st.session_state["current_sheet_headers"]
            headers_df = pd.DataFrame({
                "Поточне ім'я змінної": headers_list,
                "Нова назва змінної": headers_list
            })
            
            edited_headers_df = st.data_editor(
                headers_df,
                column_config={
                    "Поточне ім'я змінної": st.column_config.TextColumn(disabled=True),
                    "Нова назва змінної": st.column_config.TextColumn(disabled=False)
                },
                hide_index=True,
                width="stretch",
                key="headers_data_editor"
            )
            
            # Check for header modifications
            changed_headers = {}
            for idx, r in edited_headers_df.iterrows():
                orig = r["Поточне ім'я змінної"]
                new_val = r["Нова назва змінної"].strip()
                if new_val and new_val != orig:
                    changed_headers[orig] = new_val
                    
            if changed_headers:
                error_occurred = False
                for orig, new_val in changed_headers.items():
                    if new_val in st.session_state["current_sheet_headers"]:
                        st.error(f"Помилка: змінна з ім'ям '{new_val}' вже існує!")
                        error_occurred = True
                    else:
                        h_idx = st.session_state["current_sheet_headers"].index(orig)
                        st.session_state["current_sheet_headers"][h_idx] = new_val
                        
                        # Rename key in row data
                        for row in st.session_state["current_sheet_data"]:
                            if orig in row:
                                row[new_val] = row.pop(orig)
                                
                        # Queue in pending template renames to update template on Save Excel
                        if "pending_template_renames" not in st.session_state:
                            st.session_state["pending_template_renames"] = []
                        st.session_state["pending_template_renames"].append((orig, new_val))
                        
                if not error_occurred:
                    st.success("Назви змінних оновлено!")
                    time.sleep(0.5)
                    st.rerun()

            headers = st.session_state["current_sheet_headers"]
            rows = st.session_state["current_sheet_data"]
            
            if rows:
                df = pd.DataFrame(rows, columns=headers)
            else:
                df = pd.DataFrame(columns=headers)
                
            # Insert the "Перегляд" selection column as boolean at the very beginning (index 0)
            df.insert(0, "Перегляд", False)
            
            # Setup session state tracking for preview index
            if "last_preview_row_idx" not in st.session_state:
                st.session_state["last_preview_row_idx"] = 0
                
            last_idx = st.session_state["last_preview_row_idx"]
            if last_idx is not None and last_idx < len(df):
                df.at[last_idx, "Перегляд"] = True
                
            edited_df = st.data_editor(
                df,
                num_rows="dynamic",
                width="stretch",
                key="config_data_editor"
            )
            
            # Determine selected index, default to the last tracked one
            selected_row_idx = st.session_state["last_preview_row_idx"]
            
            st.write(" ")
            
            if st.button("💾 Зберегти зміни в Excel", type="primary", key="btn_save_config", use_container_width=True):
                clean_rows = []
                for _, r in edited_df.iterrows():
                    row_dict = {}
                    for h in headers:
                        val = r.get(h, "")
                        row_dict[h] = str(val) if pd.notna(val) else ""
                    clean_rows.append(row_dict)
                    
                st.session_state["current_sheet_data"] = clean_rows
                
                success = save_excel_config(
                    cfg_path,
                    selected_sheet,
                    st.session_state["editor_template_path"],
                    st.session_state["editor_name_pattern"],
                    headers,
                    clean_rows
                )
                if success:
                    # Apply any pending template variable renames when saving config
                    pending_renames = st.session_state.get("pending_template_renames", [])
                    if pending_renames:
                        t_path = st.session_state.get("editor_template_path", "")
                        if t_path:
                            from _templates_machine_ import resolve_path
                            cfg_dir = os.path.dirname(os.path.abspath(cfg_path))
                            actual_t_path = resolve_path(cfg_dir, t_path)
                            
                            if os.path.exists(actual_t_path):
                                renamed_count = 0
                                for old_n, new_n in pending_renames:
                                    if rename_placeholder_in_template(actual_t_path, old_n, new_n):
                                        renamed_count += 1
                                if renamed_count > 0:
                                    st.toast(f"✅ Шаблон також оновлено (перейменовано {renamed_count} змінних)", icon="📄")
                        
                        # Clear pending renames list after successfully applying them
                        st.session_state["pending_template_renames"] = []

                    st.success("🎉 Всі зміни успішно записані в Excel файл!")
                    st.balloons()
                    time.sleep(1)
                    st.rerun()
                    
            st.write(" ")  # spacer
            # --- HIGH-FIDELITY LIVE DOCUMENT PREVIEW WINDOW ---
            if selected_row_idx is not None and selected_row_idx < len(edited_df):
                st.write(" ")
                st.write(" ")
                st.markdown(f"### 🔍 Інтелектуальний миттєвий перегляд документа (Рядок {selected_row_idx + 5})")
                
                # Extract variables for this row
                row_vars = {}
                row_data = edited_df.iloc[selected_row_idx]
                for h in headers:
                    val = row_data.get(h, "")
                    row_vars[h] = str(val) if pd.notna(val) else ""
                    
                # Map pending renames to support both old and new names in live preview before saving
                pending_renames = st.session_state.get("pending_template_renames", [])
                for old_n, new_n in pending_renames:
                    if new_n in row_vars:
                        row_vars[old_n] = row_vars[new_n]
                    
                t_path = st.session_state["editor_template_path"]
                if t_path:
                    from _templates_machine_ import resolve_path
                    cfg_dir = os.path.dirname(os.path.abspath(cfg_path))
                    actual_template_path = resolve_path(cfg_dir, t_path)
                            
                    ext = os.path.splitext(t_path)[1].lower()
                    
                    with st.spinner("⏳ Генерація прев'ю..."):
                        if ext == ".docx":
                            preview_html = generate_docx_preview(actual_template_path, row_vars)
                        elif ext == ".xlsx":
                            preview_html = generate_xlsx_preview(actual_template_path, row_vars)
                        else:
                            preview_html = f"<div style='color: #e53e3e;'>Непідтримуваний тип шаблону: {ext}</div>"
                            
                    st.markdown(
                        f"""
                        <div style="border: 2px solid #3182ce; border-radius: 8px; padding: 25px; background-color: #ffffff; max-height: 600px; overflow-y: auto; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1), 0 2px 4px -1px rgba(0,0,0,0.06); border-left: 8px solid #3182ce;">
                            {preview_html}
                        </div>
                        """,
                        unsafe_allow_html=True
                    )
                else:
                    st.warning("Будь ласка, вкажіть шлях до файлу шаблону в полі A1, щоб активувати попередній перегляд!")

# ----------------------------------------------------
# VIEW 3: DOCUMENT GENERATION & PROGRESS TRACKING
# ----------------------------------------------------
elif selected_view == "⚡ Генерація Документів":
    st.header("⚡ Масова генерація документів")
    st.write("Оберіть файл конфігурації, перевірте шлях до шаблону і запустіть автоматичне заповнення.")
    
    col_g1, col_g2 = st.columns([3, 1])
    with col_g1:
        g_config_input = st.text_input(
            "Шлях до файлу конфігурації для генерації:",
            placeholder="Оберіть Excel-файл конфігурації...",
            key="gen_config_path"
        )
    with col_g2:
        st.write(" ")
        st.write(" ")
        def select_gen_config():
            selected = open_file_picker(filetypes=[("Excel конфігурації", "*.xlsx"), ("Усі файли", "*.*")])
            if selected:
                st.session_state["gen_config_path"] = selected
        st.button("📁 Обрати конфіг", key="btn_gen_config", on_click=select_gen_config)
        
    # Sync document generation path to last_opened state if it changes
    g_path = st.session_state["gen_config_path"]
    if g_path:
        try:
            abs_g_path = os.path.abspath(g_path)
            g_folder = os.path.dirname(abs_g_path)
            if g_path != st.session_state.get("last_opened_config") or g_folder != st.session_state.get("last_opened_folder"):
                st.session_state["last_opened_config"] = g_path
                st.session_state["last_opened_folder"] = g_folder
                st.session_state["pm_folder_path"] = g_folder
                save_persistent_state()
        except Exception:
            pass
    
    # --- ALWAYS-VISIBLE OUTPUT DIRECTORY SELECTION ---
    st.markdown("##### 📁 Папка для збереження готових документів (Необов'язково)")
    st.caption("Якщо не вказано, файли зберігатимуться відносно папки Excel-конфігурації. У безвіконному (headless) режимі введіть шлях вручну.")
    
    col_go1, col_go2 = st.columns([3, 1])
    with col_go1:
        g_out_dir = st.text_input(
            "📁 Шлях до папки збереження:",
            placeholder="Наприклад: C:/ProcessedDocs (залиште порожнім за замовчуванням)",
            key="gen_output_dir"
        )
    with col_go2:
        st.write(" ")
        st.write(" ")
        def select_gen_output_dir():
            selected = open_folder_picker("Оберіть папку для збереження готових документів")
            if selected:
                st.session_state["gen_output_dir"] = selected
        st.button("📁 Обрати папку", key="btn_gen_output_dir", on_click=select_gen_output_dir)
    if not g_path:
        st.info("Будь ласка, оберіть Excel-файл конфігурації для генерації!")
    elif not os.path.exists(g_path):
        st.error(f"Вказаний файл конфігурації '{g_path}' не знайдено!")
    else:
        g_sheets_data = load_excel_config(g_path)
        
        if g_sheets_data:
            g_sheet_names = ["all (Всі аркуші)"] + list(g_sheets_data.keys())
            selected_g_sheet = st.selectbox("Оберіть аркуш для обробки:", g_sheet_names, key="gen_sheet_select")
            
            actual_sheet_name = ""
            if selected_g_sheet != "all (Всі аркуші)":
                actual_sheet_name = selected_g_sheet
                
            # State synchronization on config/sheet change for generator
            gen_sheet_key = f"gen_{g_path}_{actual_sheet_name}"
            if st.session_state.get("loaded_gen_sheet") != gen_sheet_key:
                st.session_state["loaded_gen_sheet"] = gen_sheet_key
                if actual_sheet_name:
                    st.session_state["gen_template_path"] = g_sheets_data[actual_sheet_name]["template_path"]
                else:
                    st.session_state["gen_template_path"] = ""
                
            # --- VIEW/CHANGE TEMPLATE FILE VIA DIALOG ---
            if actual_sheet_name:
                sheet_data = g_sheets_data[actual_sheet_name]
                current_template = sheet_data["template_path"]
                
                st.markdown("##### 📄 Перевірка та зміна файлу шаблону (комірка A1)")
                
                col_gt1, col_gt2 = st.columns([3, 1])
                with col_gt1:
                    new_gt_path = st.text_input(
                        "📄 Поточний шлях до шаблону:",
                        placeholder="Оберіть файл шаблону...",
                        key="gen_template_path"
                    )
                with col_gt2:
                    st.write(" ")
                    st.write(" ")
                    def select_gen_template():
                        selected = open_file_picker(filetypes=[
                            ("Усі шаблони", "*.docx;*.xlsx"),
                            ("Word шаблони", "*.docx"),
                            ("Excel шаблони", "*.xlsx"),
                            ("Усі файли", "*.*")
                        ])
                        if selected:
                            st.session_state["gen_template_path"] = selected
                            if update_config_template_path(g_path, actual_sheet_name, selected):
                                st.toast(f"✅ Шаблон оновлено в Excel: {os.path.basename(selected)}", icon="💾")
                    st.button("📁 Змінити шаблон", key="btn_gen_template", on_click=select_gen_template)
                    
                # If path was changed manually
                if st.session_state.get("gen_template_path", "") != current_template:
                    if st.button("💾 Зберегти новий шлях шаблону в Excel", key="btn_save_gen_template"):
                        manually_entered = st.session_state.get("gen_template_path", "")
                        if update_config_template_path(g_path, actual_sheet_name, manually_entered):
                            st.success("Шлях шаблону успішно оновлено в конфігурації!")
                            time.sleep(1)
                            st.rerun()
                            
            # Row selector filtering options
            selected_g_row = "all"
            if selected_g_sheet != "all (Всі аркуші)":
                rows_count = len(sheet_data["rows"])
                
                row_options = ["all (Всі рядки з даними)"]
                for i in range(rows_count):
                    preview_fields = []
                    row_dict = sheet_data["rows"][i]
                    headers_to_show = sheet_data["headers"][:3]
                    for h in headers_to_show:
                        val = row_dict.get(h, "")
                        if val:
                            preview_fields.append(f"{h}: {val}")
                    preview_str = ", ".join(preview_fields) if preview_fields else "Порожній рядок"
                    row_options.append(f"{5 + i} — ({preview_str})")
                    
                selected_g_row_str = st.selectbox("Оберіть рядок для обробки (за номером рядка в Excel):", row_options, key="gen_row_select")
                
                if "all" in selected_g_row_str:
                    selected_g_row = "all"
                else:
                    selected_g_row = selected_g_row_str.split(" — ")[0].strip()
                    
            st.markdown("---")
            st.subheader("🎯 Параметри запуску")
            
            st.markdown(f"""
            - **Файл конфігурації:** `{g_path}`
            - **Аркуш:** `{selected_g_sheet.split(' (')[0]}`
            - **Рядок:** `{selected_g_row}`
            - **Папка збереження:** `{st.session_state["gen_output_dir"] if st.session_state["gen_output_dir"] else "Папка конфігурації (за замовчуванням)"}`
            """)
            
            if st.button("⚡ Запустити генерацію документів", type="primary", key="btn_run_generation"):
                args = [g_path]
                
                sheet_arg = selected_g_sheet.split(' (')[0]
                row_arg = selected_g_row
                
                if sheet_arg != "all":
                    args.append(sheet_arg)
                    if row_arg != "all":
                        args.append(row_arg)
                elif row_arg != "all":
                    args.append("all")
                    args.append(row_arg)
                    
                # Append custom output folder if provided
                if st.session_state["gen_output_dir"]:
                    while len(args) < 3:
                        args.append("all")
                    args.append(st.session_state["gen_output_dir"])
                    
                run_subprocess_and_stream(args)
                st.rerun() # Refresh to display persistent logs immediately

    # Persistent log viewer at the bottom of generator tab
    show_last_operation_logs()

# ----------------------------------------------------
# VIEW 4: HELP & DOCUMENTATION
# ----------------------------------------------------
elif selected_view == "📖 Повна Довідка":
    # Inject giant card container styling for the documentation page
    st.markdown(f"""
    <style>
        [data-testid="stMainBlockContainer"], .block-container {{
            background: {card_bg} !important;
            backdrop-filter: blur(10px) !important;
            -webkit-backdrop-filter: blur(10px) !important;
            border: 1px solid {card_border} !important;
            border-radius: 16px !important;
            box-shadow: 0 8px 32px 0 {shadow_color} !important;
            padding: 3rem !important;
            margin-top: 1.5rem !important;
            margin-bottom: 2.5rem !important;
        }}
    </style>
    """, unsafe_allow_html=True)

    st.header("📖 Повний посібник користувача")
    st.write("Детальний опис можливостей та технічний посібник роботи комбайна (завантажено з _templates_machine_.txt).")
    
    st.markdown("---")
    
    doc_markdown = get_formatted_documentation_markdown()
    st.markdown(doc_markdown)

# End of file