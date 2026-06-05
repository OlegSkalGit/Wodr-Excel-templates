import streamlit as st
import os
import json
import math
from datetime import datetime, date
from core.io_utils import load_excel_config

STATE_FILE = ".last_state.json"
_cached_file_state = None

def get_persisted_state_dict():
    global _cached_file_state
    if _cached_file_state is not None:
        return _cached_file_state
    if os.path.exists(STATE_FILE):
        try:
            with open(STATE_FILE, "r", encoding="utf-8") as f:
                _cached_file_state = json.load(f)
                return _cached_file_state
        except Exception:
            pass
    _cached_file_state = {}
    return _cached_file_state

def init_state_key(key, default_value):
    if key in st.session_state:
        return
    state = get_persisted_state_dict()
    if key in state and state[key] is not None:
        st.session_state[key] = state[key]
    else:
        st.session_state[key] = default_value

def load_persistent_state():
    global _cached_file_state
    _cached_file_state = None
    state = get_persisted_state_dict()
    for k, v in state.items():
        if v is not None:
            st.session_state[k] = v
    st.session_state["pm_cached_configs"] = {}
    sync_pm_editing_vars()

def make_json_serializable(obj):
    if obj is None:
        return None
    if isinstance(obj, (bool, str, int)):
        return obj
    if isinstance(obj, float):
        if math.isnan(obj) or math.isinf(obj):
            return str(obj)
        return obj
    if isinstance(obj, (datetime, date)):
        return obj.isoformat()
    if isinstance(obj, dict):
        return {str(k): make_json_serializable(v) for k, v in obj.items()}
    if isinstance(obj, (list, tuple, set)):
        return [make_json_serializable(v) for v in obj]
    
    try:
        import pandas as pd
        if isinstance(obj, (pd.DataFrame, pd.Series)):
            return make_json_serializable(obj.to_dict())
        if pd.isna(obj):
            return None
    except ImportError:
        pass

    try:
        import numpy as np
        if isinstance(obj, (np.integer, np.floating)):
            return make_json_serializable(obj.item())
        if isinstance(obj, np.ndarray):
            return make_json_serializable(obj.tolist())
    except ImportError:
        pass

    try:
        json.dumps(obj)
        return obj
    except TypeError:
        return str(obj)

def save_persistent_state():
    state = dict(get_persisted_state_dict())
    
    tracked_keys = [
        "theme",
        "current_view",
        "active_selection_type",
        "selected_template_path",
        "selected_folder_path",
        "last_opened_folder",
        "last_opened_config",
        "last_opened_template",
        "pm_folder_path",
        "editor_config_path",
        "editor_template_path",
        "editor_name_pattern",
        "pm_selected_doc",
        "analysis_mode",
        "analysis_output_dir",
        "txt_auto_folder",
        "txt_batch_sample",
        "txt_batch_folder",
        "txt_pair_file1",
        "txt_pair_file2",
        "editor_selected_sheet",
        "gen_config_path",
        "gen_output_dir",
        "pm_only_docs",
        "current_sheet_headers",
        "current_sheet_data",
        "last_preview_row_idx",
        "gen_sheet_select",
        "gen_template_path",
        "gen_row_select",
        "gen_scope_select",
        "pm_editing_vars",
        "pending_template_renames",
        "pm_loaded_doc_key",
        "pm_loaded_doc_mtime",
        "last_gen_params",
        "gen_completion_status",
        "adjustable_columns_widths_sheet_cols",
        "adjustable_columns_widths_doc_cols",
        "adjustable_columns_widths_folder_cols",
        "adjustable_columns_hidden_sheet_cols",
        "adjustable_columns_hidden_doc_cols",
        "adjustable_columns_hidden_folder_cols",
        "adjustable_columns_widths_workspace_cols",
        "adjustable_columns_hidden_workspace_cols",
        "workspace_nav_width"
    ]
    
    for key in tracked_keys:
        if key in st.session_state:
            state[key] = make_json_serializable(st.session_state[key])
            
    try:
        with open(STATE_FILE, "w", encoding="utf-8") as f:
            json.dump(state, f, ensure_ascii=False, indent=4)
        global _cached_file_state
        _cached_file_state = state
    except Exception:
        try:
            import traceback
            with open("save_error.log", "w", encoding="utf-8") as f:
                traceback.print_exc(file=f)
        except Exception:
            pass

def get_cached_config(config_path):
    if "pm_cached_configs" not in st.session_state:
        st.session_state["pm_cached_configs"] = {}
        
    try:
        mtime = os.path.getmtime(config_path)
    except Exception:
        mtime = 0
        
    cache_entry = st.session_state["pm_cached_configs"].get(config_path)
    if cache_entry and cache_entry.get("mtime") == mtime:
        return cache_entry["data"]
        
    try:
        data = load_excel_config(config_path)
    except Exception as e:
        st.error(f"Помилка при завантаженні конфігу {config_path}: {e}")
        data = None

    if data:
        st.session_state["pm_cached_configs"][config_path] = {
            "mtime": mtime,
            "data": data
        }
    return data

def sync_pm_editing_vars():
    selected_doc = st.session_state.get("pm_selected_doc")
    if selected_doc:
        config_path = selected_doc.get("config_path")
        sheet_name = selected_doc.get("sheet_name")
        row_idx = selected_doc.get("row_idx")
        if config_path and sheet_name and row_idx is not None:
            current_doc_key = f"{config_path}_{sheet_name}_{row_idx}"
            loaded_doc_key = st.session_state.get("pm_loaded_doc_key")
            
            try:
                mtime = os.path.getmtime(config_path)
            except Exception:
                mtime = 0
                
            if (current_doc_key != loaded_doc_key 
                or st.session_state.get("pm_editing_vars") is None 
                or st.session_state.get("pm_loaded_doc_mtime") != mtime):
                
                if "pm_cached_configs" not in st.session_state:
                    st.session_state["pm_cached_configs"] = {}
                if config_path in st.session_state["pm_cached_configs"]:
                    del st.session_state["pm_cached_configs"][config_path]
                sheets_data = get_cached_config(config_path)
                if sheets_data and sheet_name in sheets_data:
                    rows = sheets_data[sheet_name].get("rows", [])
                    if row_idx < len(rows):
                        st.session_state["pm_editing_vars"] = dict(rows[row_idx])
                        st.session_state["pm_loaded_doc_key"] = current_doc_key
                        st.session_state["pm_loaded_doc_mtime"] = mtime
                        save_persistent_state()

def sync_data_editor_states():
    if "current_sheet_data" in st.session_state and "current_sheet_headers" in st.session_state:
        cfg_path = st.session_state.get("editor_config_path", "")
        selected_sheet = st.session_state.get("editor_selected_sheet", "")
        if cfg_path and selected_sheet:
            clean_cfg_path = "".join([c if c.isalnum() else "_" for c in cfg_path])
            config_key = f"config_data_editor_{clean_cfg_path}_{selected_sheet}"
            headers_key = f"headers_data_editor_{clean_cfg_path}_{selected_sheet}"
        else:
            config_key = "config_data_editor"
            headers_key = "headers_data_editor"

        if config_key in st.session_state and st.session_state[config_key] is not None:
            editor_state = st.session_state[config_key]
            
            edited_rows = editor_state.get("edited_rows", {})
            for row_idx_str, edits in edited_rows.items():
                row_idx = int(row_idx_str)
                if row_idx < len(st.session_state["current_sheet_data"]):
                    for col, val in edits.items():
                        if col in st.session_state["current_sheet_headers"]:
                            st.session_state["current_sheet_data"][row_idx][col] = str(val)
                            
            added_rows = editor_state.get("added_rows", [])
            for row in added_rows:
                new_row = {h: str(row.get(h, "")) for h in st.session_state["current_sheet_headers"]}
                st.session_state["current_sheet_data"].append(new_row)
                
            deleted_rows = editor_state.get("deleted_rows", [])
            for row_idx in sorted(deleted_rows, reverse=True):
                if row_idx < len(st.session_state["current_sheet_data"]):
                    st.session_state["current_sheet_data"].pop(row_idx)

            if isinstance(st.session_state.get(config_key), dict):
                st.session_state[config_key].setdefault("edited_rows", {}).clear()
                st.session_state[config_key].setdefault("added_rows", []).clear()
                st.session_state[config_key].setdefault("deleted_rows", []).clear()

        if headers_key in st.session_state and st.session_state[headers_key] is not None:
            h_editor_state = st.session_state[headers_key]
            edited_rows = h_editor_state.get("edited_rows", {})
            for row_idx_str, edits in edited_rows.items():
                row_idx = int(row_idx_str)
                if row_idx < len(st.session_state["current_sheet_headers"]):
                    orig = st.session_state["current_sheet_headers"][row_idx]
                    new_val = edits.get("Нова назва змінної", "").strip()
                    if new_val and new_val != orig:
                        if new_val not in st.session_state["current_sheet_headers"]:
                            st.session_state["current_sheet_headers"][row_idx] = new_val
                            for row in st.session_state["current_sheet_data"]:
                                if orig in row:
                                    row[new_val] = row.pop(orig)
                            
                            # Also rename in editor_name_pattern and editor_template_path if present in session state
                            import re
                            pat = re.compile(r'(\{\{\s*)' + re.escape(orig) + r'(\s*\}\})')
                            if st.session_state.get("editor_template_path"):
                                st.session_state["editor_template_path"] = pat.sub(lambda m: m.group(1) + new_val + m.group(2), st.session_state["editor_template_path"])
                            if st.session_state.get("editor_name_pattern"):
                                st.session_state["editor_name_pattern"] = pat.sub(lambda m: m.group(1) + new_val + m.group(2), st.session_state["editor_name_pattern"])

                            if "pending_template_renames" not in st.session_state:
                                st.session_state["pending_template_renames"] = []
                            st.session_state["pending_template_renames"].append((orig, new_val))
                        else:
                            st.toast(f"⚠️ Змінна '{new_val}' вже існує!", icon="⚠️")

            if isinstance(st.session_state.get(headers_key), dict):
                st.session_state[headers_key].setdefault("edited_rows", {}).clear()

def sync_pm_inputs():
    selected_doc = st.session_state.get("pm_selected_doc")
    if selected_doc and isinstance(st.session_state.get("pm_editing_vars"), dict):
        config_path = selected_doc.get("config_path")
        sheet_name = selected_doc.get("sheet_name")
        row_idx = selected_doc.get("row_idx")
        if config_path and sheet_name and row_idx is not None:
            prefix = f"pm_input_{config_path}_{sheet_name}_{row_idx}_"
            edited_vars = st.session_state["pm_editing_vars"]
            changed = False
            for key in list(st.session_state.keys()):
                if key.startswith(prefix):
                    var_name = key[len(prefix):]
                    new_val = st.session_state[key]
                    if edited_vars.get(var_name) != new_val:
                        edited_vars[var_name] = new_val
                        changed = True
            if changed:
                save_persistent_state()

def clear_pm_input_keys():
    keys_to_del = [k for k in st.session_state.keys() if k.startswith("pm_input_")]
    for k in keys_to_del:
        del st.session_state[k]

def initialize_all_states():
    # Load state from file
    state = get_persisted_state_dict()
    
    # Define default values for all tracked keys
    defaults = {
        "theme": "light",
        "current_view": "workspace",
        "active_selection_type": "folder",
        "selected_template_path": "",
        "selected_folder_path": "",
        "last_opened_folder": "",
        "last_opened_config": "",
        "last_opened_template": "",
        "pm_folder_path": "",
        "editor_config_path": "",
        "editor_template_path": "",
        "editor_name_pattern": "",
        "pm_selected_doc": None,
        "analysis_mode": "✈️ Повний автопілот (Групування та створення шаблонів)",
        "analysis_output_dir": "",
        "txt_auto_folder": "",
        "txt_batch_sample": "",
        "txt_batch_folder": "",
        "txt_pair_file1": "",
        "txt_pair_file2": "",
        "editor_selected_sheet": "",
        "gen_config_path": "",
        "gen_output_dir": "",
        "pm_only_docs": False,
        "current_sheet_headers": [],
        "current_sheet_data": [],
        "last_preview_row_idx": 0,
        "gen_sheet_select": "all (Всі аркуші)",
        "gen_template_path": "",
        "gen_row_select": "all (Всі рядки з даними)",
        "gen_scope_select": "Тільки поточний аркуш",
        "pm_editing_vars": None,
        "pending_template_renames": [],
        "pm_loaded_doc_key": "",
        "pm_loaded_doc_mtime": 0,
        "last_gen_params": "",
        "gen_completion_status": None,
        "adjustable_columns_widths_sheet_cols": {},
        "adjustable_columns_widths_doc_cols": {},
        "adjustable_columns_widths_folder_cols": {},
        "adjustable_columns_hidden_sheet_cols": [],
        "adjustable_columns_hidden_doc_cols": [],
        "adjustable_columns_hidden_folder_cols": [],
        "adjustable_columns_widths_workspace_cols": {},
        "adjustable_columns_hidden_workspace_cols": [],
        "workspace_nav_width": 280
    }
    
    # Restore from persisted file or use defaults for ALL tracked keys on every rerun
    for key, def_val in defaults.items():
        if key not in st.session_state:
            if key in state and state[key] is not None:
                st.session_state[key] = state[key]
            else:
                st.session_state[key] = def_val
                
    if "state_loaded" not in st.session_state:
        st.session_state["pm_cached_configs"] = {}
        sync_pm_editing_vars()
        st.session_state["state_loaded"] = True

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

