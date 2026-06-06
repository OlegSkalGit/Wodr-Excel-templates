import streamlit as st
import os
import sys
import time
import shutil
import tkinter as tk
from tkinter import filedialog
import pandas as pd

# Core utilities
from core.io_utils import (
    scan_recursive_configs,
    build_dir_tree,
    build_docs_only_tree,
    load_excel_config,
    save_excel_config,
    update_config_template_path
)
from core.text_processor import resolve_virtual_doc_name, resolve_path

# UI States & CSS
from ui.state_manager import (
    save_persistent_state,
    load_persistent_state,
    init_state_key,
    clear_pm_input_keys,
    open_folder_picker,
    open_file_picker,
    get_cached_config,
    sync_pm_editing_vars,
    sync_data_editor_states,
    sync_pm_inputs,
    initialize_all_states,
    get_persisted_state_dict
)
from ui.styles.css_builder import get_custom_css

# UI Components
from ui.components.file_tree import render_docs_only_tree, render_tree_node, scan_recursive_templates
from ui.components.document_preview import generate_docx_preview, generate_xlsx_preview
from ui.components.config_table import (
    render_config_editor,
    rename_placeholder_in_template,
    extract_placeholders_with_context
)
from ui.components.variables_analyzer import render_variables_analyzer
from ui.components.logger_console import run_subprocess_and_stream, show_last_operation_logs

def _cb_pick_folder(state_key, title):
    selected = open_folder_picker(title)
    if selected:
        st.session_state[state_key] = selected
        save_persistent_state()

def _cb_pick_project_folder(title):
    selected = open_folder_picker(title)
    if selected:
        st.session_state["pm_folder_path"] = selected
        st.session_state["last_opened_folder"] = selected
        st.session_state["active_selection_type"] = "folder"
        st.session_state["selected_folder_path"] = selected
        save_persistent_state()

def _cb_pick_file(state_key):
    selected = open_file_picker()
    if selected:
        st.session_state[state_key] = selected
        save_persistent_state()

def scan_workspace():
    """Scans the directory for configs and templates."""
    files = os.listdir('.')
    configs = [f for f in files if f.endswith('.xlsx') and not f.startswith('~$') and 'template' not in f.lower()]
    templates = [f for f in files if (f.endswith('.docx') or f.endswith('.xlsx')) and f.startswith('template_')]
    
    if os.path.exists('example'):
        for f in os.listdir('example'):
            if f.endswith('.xlsx') and not f.startswith('~$') and 'template' not in f.lower():
                configs.append(os.path.join('example', f))
            if (f.endswith('.docx') or f.endswith('.xlsx')) and f.startswith('template_'):
                templates.append(os.path.join('example', f))
                
    return sorted(configs), sorted(templates)

def recreate_bat_file(bat_path, config_path):
    """Recreates the execution .bat file at the destination with correct relative paths."""
    try:
        s_dir = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
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
        dest_dir = os.path.abspath(dest_dir)
        os.makedirs(dest_dir, exist_ok=True)
        for filename in ["config_Auto.xlsx", "Auto_Run_All.bat", "config__NODublicate_.xlsx"]:
            src = os.path.join(os.getcwd(), filename)
            if os.path.exists(src):
                shutil.move(src, os.path.join(dest_dir, filename))
                
        nodup_src = os.path.join(os.getcwd(), "_NODublicate_")
        if os.path.exists(nodup_src) and os.path.isdir(nodup_src):
            nodup_dest = os.path.join(dest_dir, "_NODublicate_")
            if os.path.exists(nodup_dest):
                shutil.rmtree(nodup_dest)
            shutil.move(nodup_src, nodup_dest)
            
        moved_count = 0
        for filename in os.listdir(os.getcwd()):
            if filename.startswith("template_") and os.path.isfile(filename):
                shutil.move(filename, os.path.join(dest_dir, filename))
                moved_count += 1
                
        dest_bat = os.path.join(dest_dir, "Auto_Run_All.bat")
        dest_cfg = os.path.join(dest_dir, "config_Auto.xlsx")
        if os.path.exists(dest_bat) and os.path.exists(dest_cfg):
            recreate_bat_file(dest_bat, dest_cfg)
            
        toast_msg = f"✅ Результати аналізу перенесено в: {dest_dir}"
        if os.path.exists(os.path.join(dest_dir, "config_Auto.xlsx")):
            toast_msg += f" (переміщено {moved_count} шаблонів)"
        if os.path.exists(os.path.join(dest_dir, "config__NODublicate_.xlsx")):
            toast_msg += " (створено config__NODublicate_.xlsx)"
        st.toast(toast_msg, icon="📁")
    except Exception as e:
        st.error(f"Помилка при перенесенні файлів результатів: {e}")

def move_batch_outputs(sample_path, dest_dir):
    """Moves generated templates and configs of batch mode to a custom destination directory."""
    if not dest_dir or not sample_path:
        return
    try:
        dest_dir = os.path.abspath(dest_dir)
        os.makedirs(dest_dir, exist_ok=True)
        f_dir = os.path.dirname(os.path.abspath(sample_path))
        base_name, ext = os.path.splitext(os.path.basename(sample_path))
        
        files_to_move = [
            f"template_{base_name}{ext}",
            f"config_{base_name}.xlsx",
            f"{base_name}_run_all.bat"
        ]
        moved_count = 0
        for filename in files_to_move:
            src = os.path.join(f_dir, filename)
            if os.path.exists(src):
                shutil.move(src, os.path.join(dest_dir, filename))
                moved_count += 1
                
        dest_bat = os.path.join(dest_dir, f"{base_name}_run_all.bat")
        dest_cfg = os.path.join(dest_dir, f"config_{base_name}.xlsx")
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
        dest_dir = os.path.abspath(dest_dir)
        os.makedirs(dest_dir, exist_ok=True)
        f_dir = os.path.dirname(os.path.abspath(file1_path))
        base_name, ext = os.path.splitext(os.path.basename(file1_path))
        
        files_to_move = [
            f"template_{base_name}{ext}",
            f"config_{base_name}.xlsx",
            f"{base_name}_run_all.bat"
        ]
        moved_count = 0
        for filename in files_to_move:
            src = os.path.join(f_dir, filename)
            if os.path.exists(src):
                shutil.move(src, os.path.join(dest_dir, filename))
                moved_count += 1
                
        dest_bat = os.path.join(dest_dir, f"{base_name}_run_all.bat")
        dest_cfg = os.path.join(dest_dir, f"config_{base_name}.xlsx")
        if os.path.exists(dest_bat) and os.path.exists(dest_cfg):
            recreate_bat_file(dest_bat, dest_cfg)
            
        st.toast(f"✅ Результати попарного порівняння перенесено в: {dest_dir} (переміщено {moved_count} файлів)", icon="📁")
    except Exception as e:
        st.error(f"Помилка при перенесенні файлів результатів: {e}")

def save_generated_document_dialog(template_path, variables, config_path, name_pattern=None):
    """Generates the document and opens a native Windows dialog to save it, with fallback."""
    from core.file_handlers.docx_handler import process_word
    from core.file_handlers.xlsx_handler import process_excel
    import tempfile
    
    cfg_dir = os.path.dirname(os.path.abspath(config_path))
    actual_t_path = resolve_path(cfg_dir, template_path)
        
    if not os.path.exists(actual_t_path):
        st.error(f"Шаблон не знайдено: {template_path}")
        return
        
    ext = os.path.splitext(actual_t_path)[1].lower()
    
    if not name_pattern:
        name_pattern = variables.get("name_pattern", "document")
    proposed_filename = os.path.basename(resolve_virtual_doc_name(name_pattern, variables, template_path))
    
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
        temp_dir = tempfile.gettempdir()
        temp_file = os.path.join(temp_dir, proposed_filename)
        try:
            if ext == ".docx":
                process_word(actual_t_path, temp_file, variables)
            elif ext == ".xlsx":
                process_excel(actual_t_path, temp_file, variables)
                
            with open(temp_file, "rb") as f:
                file_bytes = f.read()
                
            st.session_state["pending_download"] = {
                "bytes": file_bytes,
                "name": proposed_filename,
                "mime": "application/vnd.openxmlformats-officedocument.wordprocessingml.document" if ext == ".docx" else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            }
            save_persistent_state()
            st.rerun()
        except Exception as e_gen:
            st.error(f"Помилка генерації документа: {e_gen}")

@st.dialog("📖 Посібник користувача (TemplateMachine)", width="large")
def show_help_dialog():
    """Renders the help guide in a native Streamlit modal overlay."""
    from ui.views.help_drawer import get_formatted_documentation_markdown
    doc_markdown = get_formatted_documentation_markdown()
    st.markdown(doc_markdown)

def render_workspace():
    """Main Streamlit Page Renderer for the 3-Panel Unified Workspace."""
    # 1. Config page settings immediately
    st.set_page_config(
        page_title="TemplateMachine Workspace",
        layout="wide",
        page_icon="🚀",
        initial_sidebar_state="expanded"
    )

    # 2. Initialize persistent states
    for key in [
        "txt_auto_folder", "txt_batch_sample", "txt_batch_folder", 
        "txt_pair_file1", "txt_pair_file2", "editor_config_path", 
        "editor_template_path", "editor_name_pattern",
        "gen_config_path", "gen_output_dir", "analysis_output_dir",
        "last_opened_folder", "last_opened_config", "last_opened_template",
        "pm_folder_path", "analysis_mode", "editor_selected_sheet", "current_view",
        "loaded_config_sheet", "loaded_gen_sheet", "gen_template_path", "active_selection_type",
        "selected_folder_path", "selected_template_path"
    ]:
        init_state_key(key, "")

    init_state_key("gen_sheet_select", "all (Всі аркуші)")
    init_state_key("gen_row_select", "all (Всі рядки з даними)")
    init_state_key("pm_only_docs", False)
    init_state_key("last_operation_logs", [])
    init_state_key("last_operation_status", None)
    init_state_key("last_operation_cmd", "")
    init_state_key("pm_selected_doc", None)
    init_state_key("current_sheet_headers", [])
    init_state_key("current_sheet_data", [])
    init_state_key("last_preview_row_idx", 0)
    init_state_key("pm_editing_vars", None)
    init_state_key("pending_template_renames", [])
    init_state_key("pm_loaded_doc_key", "")
    init_state_key("gen_completion_status", None)
    init_state_key("last_gen_params", "")

    initialize_all_states()
    sync_data_editor_states()
    sync_pm_inputs()
    theme = st.session_state["theme"]

    # Reset config sheet loading state when selection type changes to force reload fresh values from the Excel file
    if "va_last_active_selection_type" not in st.session_state:
        st.session_state["va_last_active_selection_type"] = st.session_state.get("active_selection_type")
    if st.session_state.get("active_selection_type") != st.session_state["va_last_active_selection_type"]:
        st.session_state["va_last_active_selection_type"] = st.session_state.get("active_selection_type")
        st.session_state["loaded_config_sheet"] = None

    # 3. Inject Dynamic CSS
    css_str = get_custom_css(theme)
    st.markdown(css_str, unsafe_allow_html=True)

    # 3.1. Native sidebar is disabled and hidden in favor of stable column layout

    # 4. Render Title Header inside top header bar container
    with st.container():
        st.markdown('<div class="header-bar-marker"></div>', unsafe_allow_html=True)
        st.markdown('<div class="main-title" style="font-size: 2.2rem; margin-bottom: 0px; line-height: 1.2;">🚀 Панель керування TemplateMachine</div>', unsafe_allow_html=True)
        st.markdown('<div class="subtitle" style="font-size: 0.95rem; font-weight: 300; margin-top: 4px; margin-bottom: 0px;">Універсальний комбайн для автоматизації документів та аналізу архівів</div>', unsafe_allow_html=True)

    pm_path = st.session_state.get("pm_folder_path", "")

    # Onboarding view if no project directory is loaded
    if not pm_path or not os.path.exists(pm_path):
        # Render the sidebar for onboarding to avoid the empty white panel and expose useful controls
        with st.sidebar:
            col_help, col_theme = st.columns([3, 1], vertical_alignment="center")
            with col_help:
                if st.button("📖 Посібник", key="global_help_btn", use_container_width=True):
                    show_help_dialog()
            with col_theme:
                theme_emoji = "🌙" if theme == "light" else "☀️"
                if st.button(theme_emoji, key="theme_toggle_btn", use_container_width=True):
                    st.session_state["theme"] = "dark" if theme == "light" else "light"
                    st.session_state["theme_changed"] = True
                    save_persistent_state()
                    st.rerun()
            
            st.markdown("---")
            st.markdown("### 🤖 Про TemplateMachine")
            st.markdown(
                """
                **TemplateMachine** — це універсальний інструмент для інтелектуальної автоматизації роботи з документами:
                
                * 📂 **Аналіз та шаблонізація**: швидке виявлення повторюваних фрагментів і автоматичне створення шаблонів на основі Excel та Word файлів.
                * ⚙️ **Управління конфігураціями**: зручне редагування правил заповнення даних, налаштування формул і форматування.
                * 🚀 **Генерація документів**: пакетне створення сотень унікальних документів (договорів, актів, звітів) за секунди.
                * 🔍 **Розумний імпорт**: інтеграція нових даних в існуючі структури конфігурацій без втрати налаштувань.
                """
            )

        st.markdown('<div class="card" style="padding: 2.5rem; text-align: center;">', unsafe_allow_html=True)
        st.info("👋 Вітаємо у TemplateMachine! Для початку роботи оберіть робочу папку.")
        st.markdown("""
        <div style="text-align: left; margin-top: 1.5rem; margin-bottom: 2rem; padding: 1.5rem; background-color: rgba(49, 130, 206, 0.05); border-left: 4px solid #3182ce; border-radius: 4px;">
            <h4 style="margin-top: 0; color: #3182ce;">📂 Що таке робоча папка проекту та для чого вона потрібна?</h4>
            <p style="margin-bottom: 0.8rem; font-size: 0.95rem;">
                <b>Робоча папка (Workspace)</b> — це головний каталог вашого проекту, де розміщуються вихідні документи, шаблони та файли налаштувань.
            </p>
            <ul style="margin-bottom: 0; padding-left: 1.5rem; font-size: 0.9rem; line-height: 1.5;">
                <li><b>Побудова файлової структури:</b> Система сканує цю папку для відображення інтерактивного дерева файлів у боковому меню ліворуч.</li>
                <li><b>Пошук конфігів та шаблонів:</b> Додаток автоматично знаходить усі наявні файли конфігурацій (наприклад, <code>config_Auto.xlsx</code>) та шаблони в межах обраної папки.</li>
                <li><b>Відносне збереження результатів:</b> Усі відносні шляхи до шаблонів та згенерованих документів (параметр <code>rel_dir</code>) вираховуються відносно цієї папки. Це дозволяє переносити проект на будь-який інший комп'ютер без втрати зв'язків.</li>
            </ul>
        </div>
        """, unsafe_allow_html=True)
        col_on1, col_on2 = st.columns([3, 1])
        with col_on1:
            st.text_input(
                "📁 Шлях до робочої папки проекту:",
                placeholder="Вкажіть шлях до папки (наприклад, example)...",
                key="pm_folder_path",
                on_change=save_persistent_state
            )
        with col_on2:
            st.write(" ")
            st.write(" ")
            st.button(
                "📁 Обрати папку",
                key="btn_onboard_folder",
                use_container_width=True,
                on_click=_cb_pick_project_folder,
                args=("Оберіть робочу папку",)
            )
        st.markdown('</div>', unsafe_allow_html=True)
        return

    # Pre-populate active selection to root folder by default if not set
    if not st.session_state.get("active_selection_type"):
        st.session_state["active_selection_type"] = "folder"
        st.session_state["selected_folder_path"] = os.path.abspath(pm_path)
        save_persistent_state()

    sel_type = st.session_state["active_selection_type"]

    # 5. Define Main Workspace Container
    col_main = st.container()

    # 6. PANEL 1: LEFT NAVIGATION TREE (Streamlit Native Sidebar Splitter)
    with st.sidebar:
        col_help, col_theme = st.columns([3, 1], vertical_alignment="center")
        with col_help:
            if st.button("📖 Посібник", key="global_help_btn", use_container_width=True):
                show_help_dialog()
        with col_theme:
            theme_emoji = "🌙" if theme == "light" else "☀️"
            if st.button(theme_emoji, key="theme_toggle_btn", use_container_width=True):
                st.session_state["theme"] = "dark" if theme == "light" else "light"
                st.session_state["theme_changed"] = True
                save_persistent_state()
                st.rerun()

        st.subheader("🌳 Проект")
        
        # 1. Option to switch to analysis and template creation mode (First by logic)
        is_active_analysis = (
            st.session_state.get("active_selection_type") == "folder"
            and st.session_state.get("selected_folder_path") == os.path.abspath(pm_path)
        )
        analysis_btn_type = "primary" if is_active_analysis else "secondary"
        if st.button("✈️ Створення конфігів", key="nav_to_analysis_btn", use_container_width=True, type=analysis_btn_type):
            st.session_state["active_selection_type"] = "folder"
            st.session_state["selected_folder_path"] = os.path.abspath(pm_path)
            st.session_state["pm_selected_doc"] = None
            st.session_state["selected_template_path"] = None
            st.session_state["editor_config_path"] = None
            st.session_state["editor_selected_sheet"] = None
            save_persistent_state()
            st.rerun()
            
        if st.button("🌐 Усі змінні", use_container_width=True, type="primary" if st.session_state.get("active_selection_type") == "all_variables" else "secondary"):
            st.session_state["active_selection_type"] = "all_variables"
            st.session_state["force_expand_tree"] = False
            save_persistent_state()
            st.rerun()
            
        st.button(
            "📁 Папка конфігів",
            key="nav_change_folder_btn",
            use_container_width=True,
            on_click=_cb_pick_project_folder,
            args=("Оберіть робочу папку (папку конфігів)",)
        )
                
        # 3. Current folder display
        st.caption(f"**Робоча папка:**\n`{os.path.abspath(pm_path)}`")

        is_only_docs = st.session_state.get("pm_only_docs", False)
        btn_label = "Конфіги і шаблони" if is_only_docs else "Віртуальні документи"
        if st.button(btn_label, use_container_width=True):
            st.session_state["pm_only_docs"] = not is_only_docs
            st.session_state["force_expand_tree"] = True
            save_persistent_state()
            st.rerun()
            
        with st.spinner("⏳ Сканування..."):
            config_files = scan_recursive_configs(pm_path)
            dir_tree = build_dir_tree(config_files, pm_path)

        if st.session_state.get("pm_only_docs"):
            docs_tree = build_docs_only_tree(config_files, pm_path, config_loader=get_cached_config)
            if not docs_tree:
                st.info("Документів не знайдено.")
            else:
                render_docs_only_tree(docs_tree, 0)
        else:
            for name, val in dir_tree.items():
                render_tree_node(name, val, 0, current_path=pm_path)

        # Scroll retention helper for the sidebar (runs in main window via st.html)
        st.html(
            """
            <script>
            (function() {
                const parentDoc = window.parent.document;
                const parentWin = parentDoc.defaultView || window;
                
                function setupSidebarScroll() {
                    const sidebar = parentDoc.querySelector('[data-testid="stSidebarUserContent"]');
                    if (sidebar) {
                        const isPopulated = sidebar.querySelector('[data-testid="stExpander"]') || sidebar.scrollHeight > 350;
                        if (isPopulated) {
                            const savedScroll = sessionStorage.getItem('sidebar_scroll');
                            if (savedScroll) {
                                const scrollVal = parseInt(savedScroll, 10);
                                if (sidebar.scrollTop !== scrollVal) {
                                    sidebar.scrollTop = scrollVal;
                                }
                            }
                        }
                        if (!sidebar.dataset.hasScrollListener) {
                            sidebar.dataset.hasScrollListener = 'true';
                            sidebar.addEventListener('scroll', () => {
                                const isReady = sidebar.querySelector('[data-testid="stExpander"]') || sidebar.scrollHeight > 350;
                                if (isReady) {
                                    if (sidebar.scrollTop === 0 && sidebar.scrollHeight < 350) {
                                        return;
                                    }
                                    sessionStorage.setItem('sidebar_scroll', sidebar.scrollTop);
                                }
                            });
                        }
                    }
                }
                
                // Run immediately
                setupSidebarScroll();
                
                // Set up MutationObserver to re-apply scroll when Streamlit updates the DOM
                if (!parentWin.sidebarObserver) {
                    const sidebarOuter = parentDoc.querySelector('[data-testid="stSidebar"]');
                    if (sidebarOuter) {
                        parentWin.sidebarObserver = new MutationObserver(() => {
                            setupSidebarScroll();
                        });
                        parentWin.sidebarObserver.observe(sidebarOuter, { childList: true, subtree: true });
                    }
                }
            })();
            </script>
            """
        )

    # 7. PANEL 2 & PANEL 3 ROUTING BY ACTIVE NODE SELECTION

    # --- MODE A: FOLDER SELECTED (ANALYSIS VIEWS) ---
    if sel_type == "folder":
        folder_path = st.session_state.get("selected_folder_path", pm_path)
        
        with col_main:
            st.header("🔍 Режими аналізу та розпізнавання шаблонів")
            st.write("Скрипт проведе інтелектуальне порівняння документів у папці, виділить змінні і згенерує шаблони.")
            
            analysis_modes = [
                "✈️ Повний автопілот (Групування та створення шаблонів)",
                "🔍 Пакетний аналіз (Один файл-зразок + папка з іншими файлами)",
                "⚖️ Попарне порівняння (Точне порівняння двох конкретних файлів)"
            ]
            if "analysis_mode" not in st.session_state:
                st.session_state["analysis_mode"] = analysis_modes[0]
                
            mode = st.radio(
                "Оберіть режим аналізу:",
                analysis_modes,
                key="analysis_mode",
                on_change=save_persistent_state
            )
            
            st.markdown("---")
            
            if "Повний автопілот" in mode:
                st.subheader("✈️ Режим 1: Повний автопілот")
                if not st.session_state.get("txt_auto_folder"):
                    st.session_state["txt_auto_folder"] = folder_path
                col1, col2 = st.columns([3, 1])
                with col1:
                    st.text_input(
                        "Шлях до папки з архівом документів:",
                        key="txt_auto_folder",
                        on_change=save_persistent_state
                    )
                with col2:
                    st.write(" ")
                    st.write(" ")
                    st.button("📁 Обрати", key="btn_auto_folder_pick", on_click=_cb_pick_folder, args=("txt_auto_folder", "Оберіть папку з архівом"))
            elif "Пакетний аналіз" in mode:
                st.subheader("🔍 Режим 2: Пакетний аналіз")
                col1, col2 = st.columns([3, 1])
                with col1:
                    st.text_input(
                        "Шлях до файлу-зразка (.docx або .xlsx):",
                        placeholder="Оберіть файл-зразок...",
                        key="txt_batch_sample",
                        on_change=save_persistent_state
                    )
                with col2:
                    st.write(" ")
                    st.write(" ")
                    st.button("📄 Обрати", key="btn_batch_sample_pick", on_click=_cb_pick_file, args=("txt_batch_sample",))
                            
                col1, col2 = st.columns([3, 1])
                with col1:
                    if not st.session_state.get("txt_batch_folder"):
                        st.session_state["txt_batch_folder"] = folder_path
                    st.text_input(
                        "Шлях до папки порівняння:",
                        key="txt_batch_folder",
                        on_change=save_persistent_state
                    )
                with col2:
                    st.write(" ")
                    st.write(" ")
                    st.button("📁 Обрати", key="btn_batch_folder_pick", on_click=_cb_pick_folder, args=("txt_batch_folder", "Оберіть папку для пакетного аналізу"))
                            
            elif "Попарне порівняння" in mode:
                st.subheader("⚖️ Режим 3: Попарне порівняння")
                col1, col2 = st.columns([3, 1])
                with col1:
                    st.text_input(
                        "Шлях до Першого файлу:",
                        key="txt_pair_file1",
                        on_change=save_persistent_state
                    )
                with col2:
                    st.write(" ")
                    st.write(" ")
                    st.button("📄 Обрати 1", key="btn_pair_file1_pick", on_click=_cb_pick_file, args=("txt_pair_file1",))
                            
                col1, col2 = st.columns([3, 1])
                with col1:
                    st.text_input(
                        "Шлях до Другого файлу:",
                        key="txt_pair_file2",
                        on_change=save_persistent_state
                    )
                with col2:
                    st.write(" ")
                    st.write(" ")
                    st.button("📄 Обрати 2", key="btn_pair_file2_pick", on_click=_cb_pick_file, args=("txt_pair_file2",))

            st.markdown("---")
            st.subheader("⚙️ Управління та запуск аналізу")
            
            col_ao1, col_ao2 = st.columns([3, 1])
            with col_ao1:
                st.text_input(
                    "📁 Папка для результатів:",
                    placeholder="example",
                    key="analysis_output_dir",
                    on_change=save_persistent_state
                )
            with col_ao2:
                st.write(" ")
                st.write(" ")
                st.button("📁 Обрати", key="btn_analysis_output_dir_pick", help="Оберіть папку результатів", on_click=_cb_pick_folder, args=("analysis_output_dir", "Оберіть папку результатів"))
                        
            st.write(" ")
            
            if "Повний автопілот" in mode:
                st.checkbox(
                    "Ігнорувати одиничні файли",
                    value=False,
                    key="chk_ignore_single"
                )
                st.caption("ℹ️ *Якщо активовано, одиничні файли (без дублікатів) не оброблятимуться. Якщо вимкнено, файли будуть винесені в окремий конфіг config__NODublicate_.xlsx та папку _NODublicate_ у папці результатів і додаватимуться як окремі аркуші в конфігу.*")
                st.write(" ")
                
                if st.button("🚀 Запустити аналіз", key="btn_run_autopilot", type="primary", use_container_width=True):
                    f_path = st.session_state.get("txt_auto_folder")
                    if not f_path or not os.path.exists(f_path):
                        st.error("Вкажіть дійсну папку!")
                    else:
                        args = [f_path]
                        if st.session_state.get("chk_ignore_single"):
                            args.append("--ignore-single")
                        ret_code, _ = run_subprocess_and_stream(args)
                        if ret_code == 0 and st.session_state.get("analysis_output_dir"):
                            move_autopilot_outputs(st.session_state["analysis_output_dir"])
                        st.rerun()
            elif "Пакетний аналіз" in mode:
                if st.button("🚀 Запустити аналіз", key="btn_run_batch", type="primary", use_container_width=True):
                    sample = st.session_state.get("txt_batch_sample")
                    f_path = st.session_state.get("txt_batch_folder")
                    if not sample or not os.path.exists(sample) or not f_path or not os.path.exists(f_path):
                        st.error("Перевірте шляхи до файлу-зразка та папки!")
                    else:
                        ret_code, _ = run_subprocess_and_stream([sample, f_path])
                        if ret_code == 0 and st.session_state.get("analysis_output_dir"):
                            move_batch_outputs(sample, st.session_state["analysis_output_dir"])
                        st.rerun()
            elif "Попарне порівняння" in mode:
                if st.button("🚀 Запустити аналіз", key="btn_run_pairwise", type="primary", use_container_width=True):
                    file1 = st.session_state.get("txt_pair_file1")
                    file2 = st.session_state.get("txt_pair_file2")
                    if not file1 or not os.path.exists(file1) or not file2 or not os.path.exists(file2):
                        st.error("Вкажіть обидва файли для порівняння!")
                    else:
                        ret_code, _ = run_subprocess_and_stream([file1, file2])
                        if ret_code == 0 and st.session_state.get("analysis_output_dir"):
                            move_pairwise_outputs(file1, st.session_state["analysis_output_dir"])
                        st.rerun()
                        
            st.write(" ")
            show_last_operation_logs()

    # --- MODE B: SPREADSHEET SHEET SELECTED (INTERACTIVE GRID EDITOR & MASS GEN) ---
    elif sel_type == "sheet":
        cfg_path = st.session_state.get("editor_config_path")
        selected_sheet = st.session_state.get("editor_selected_sheet")
        
        with col_main:
            st.header("📝 Інтерактивний редактор таблиці Excel")
            st.caption(f"**Файл:** `{os.path.basename(cfg_path)}` | **Аркуш:** `{selected_sheet}`")
            
            # Render the configuration table grid editor
            render_config_editor(cfg_path)
            
            # Integrate Mass Generation here directly inside the Spreadsheet panel
            st.markdown("---")
            st.subheader("⚡ Масова генерація документів з цієї таблиці")
            
            current_rows = st.session_state.get("current_sheet_data")
            current_headers = st.session_state.get("current_sheet_headers")
            
            # Fallback to load from disk if session state is somehow missing
            if current_rows is None or current_headers is None:
                try:
                    g_sheets_data = load_excel_config(cfg_path)
                    if g_sheets_data and selected_sheet in g_sheets_data:
                        current_rows = g_sheets_data[selected_sheet]["rows"]
                        current_headers = g_sheets_data[selected_sheet]["headers"]
                except Exception:
                    pass
                    
            if current_rows is not None and current_headers is not None:
                rows_count = len(current_rows)
                
                col_g1, col_g2 = st.columns([3, 1], vertical_alignment="bottom")
                with col_g1:
                    st.text_input(
                        "📁 Зберегти готові документи в:",
                        placeholder="За замовчуванням (поруч з конфігом)",
                        key="gen_output_dir",
                        on_change=save_persistent_state
                    )
                with col_g2:
                    st.button(
                        "📁 Обрати",
                        key="btn_gen_output_dir_pick",
                        help="Оберіть папку збереження",
                        use_container_width=True,
                        on_click=_cb_pick_folder,
                        args=("gen_output_dir", "Оберіть папку збереження готових документів")
                    )
                            
                col_g3, col_g4 = st.columns([3, 1])
                with col_g3:
                    scope_options = [
                        "Тільки поточний аркуш",
                        "Усі аркуші з конфігу"
                    ]
                    
                    # Sanitize scope select for backward compatibility
                    if "gen_scope_select" in st.session_state:
                        val = st.session_state["gen_scope_select"]
                        if val not in scope_options:
                            if "Усі аркуші" in val:
                                st.session_state["gen_scope_select"] = "Усі аркуші з конфігу"
                            else:
                                st.session_state["gen_scope_select"] = "Тільки поточний аркуш"
                                
                    selected_scope = st.selectbox(
                        f"Масштаб генерації (поточний аркуш: {selected_sheet}):",
                        scope_options,
                        key="gen_scope_select",
                        on_change=save_persistent_state
                    )
                    
                    # Row selection dropdown for processing
                    row_options = ["all (Всі рядки з даними)"]
                    for i in range(rows_count):
                        preview_fields = []
                        row_dict = current_rows[i]
                        headers_to_show = current_headers[:3]
                        for h in headers_to_show:
                            val = row_dict.get(h, "")
                            if val:
                                preview_fields.append(f"{h}: {val}")
                        preview_str = ", ".join(preview_fields) if preview_fields else "Порожній рядок"
                        row_options.append(f"{5 + i} — ({preview_str})")
                        
                    # Safely sanitize gen_row_select before selectbox to prevent Streamlit options exception
                    if "gen_row_select" in st.session_state:
                        if st.session_state["gen_row_select"] not in row_options:
                            saved_val = st.session_state["gen_row_select"]
                            matched = False
                            if " — " in saved_val:
                                prefix = saved_val.split(" — ")[0] + " — "
                                for opt in row_options:
                                    if opt.startswith(prefix):
                                        st.session_state["gen_row_select"] = opt
                                        matched = True
                                        break
                            if not matched:
                                st.session_state["gen_row_select"] = row_options[0]
                                
                    if selected_scope == "Усі аркуші з конфігу":
                        selected_g_row = "all"
                    else:
                        selected_g_row_str = st.selectbox(
                            "Оберіть рядок для обробки (за номером рядка в Excel):",
                            row_options,
                            key="gen_row_select",
                            on_change=save_persistent_state
                        )
                        selected_g_row = "all"
                        if "all" not in selected_g_row_str:
                            selected_g_row = selected_g_row_str.split(" — ")[0].strip()
                    
                if st.button("⚡ Запустити генерацію документів", type="primary", key="btn_run_generation", use_container_width=True):
                    g_sheet_arg = "all" if selected_scope == "Усі аркуші з конфігу" else selected_sheet
                    args = [cfg_path, g_sheet_arg, selected_g_row]
                    if st.session_state["gen_output_dir"]:
                        args.append(st.session_state["gen_output_dir"])
                        
                    run_subprocess_and_stream(args)
                    st.session_state["gen_completion_status"] = "success"
                    save_persistent_state()
                    st.rerun()
                    
                if st.session_state.get("gen_completion_status") == "success":
                    st.success("🎉 Документи успішно згенеровано!")
                    st.session_state["gen_completion_status"] = None
                    save_persistent_state()
                    
                show_last_operation_logs()

    # --- MODE C: CLIENT DOCUMENT SELECTED (FORM + LIVE PREVIEW) ---
    elif sel_type == "document":
        selected_doc = st.session_state.get("pm_selected_doc")
        if not selected_doc:
            st.info("👈 Оберіть документ у дереві ліворуч.")
            return
            
        config_path = selected_doc["config_path"]
        sheet_name = selected_doc["sheet_name"]
        row_idx = selected_doc["row_idx"]
        sheets_data = get_cached_config(config_path)
        if not sheets_data or sheet_name not in sheets_data:
            st.error("Помилка завантаження даних документа.")
            return
            
        sheet_info = sheets_data[sheet_name]
        template_path = sheet_info.get("template_path", "")
        name_pattern = sheet_info.get("name_pattern", "")
        headers = sheet_info["headers"]
        
        if st.session_state.get("pm_editing_vars") is None:
            st.session_state["pm_editing_vars"] = dict(sheet_info["rows"][row_idx])
            
        edited_vars = st.session_state["pm_editing_vars"]
        
        with col_main:
            st.header("✏️ Форма редагування та перегляду документа")
            st.caption(f"**Конфіг:** `{os.path.basename(config_path)}` | **Шаблон:** `{os.path.basename(template_path) if template_path else 'Не вказано'}` | **Рядок:** `{row_idx + 5}`")
            
            current_resolved_name = resolve_virtual_doc_name(name_pattern, edited_vars, template_path)
            st.markdown(f"**📄 Вихідне ім'я:** `{current_resolved_name}`")
            
            cfg_dir = os.path.dirname(os.path.abspath(config_path))
            actual_template_path = resolve_path(cfg_dir, template_path)
            
            if not os.path.exists(actual_template_path):
                st.warning(f"Шаблон не знайдено: {template_path}")
            else:
                ext = os.path.splitext(actual_template_path)[1].lower()
                with st.spinner("⏳ Генерація прев'ю..."):
                    if ext == ".docx":
                        preview_html = generate_docx_preview(actual_template_path, edited_vars, config_path=config_path)
                    elif ext == ".xlsx":
                        preview_html = generate_xlsx_preview(actual_template_path, edited_vars, config_path=config_path)
                    else:
                        preview_html = f"<div style='color: #e53e3e;'>Непідтримуваний тип: {ext}</div>"
                        
                st.markdown(
                    f"""
                    <div class="document-preview-container" style="border: 2px solid #3182ce; border-radius: 8px; padding: 20px; background-color: #ffffff; max-height: 500px; overflow-y: auto; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1); border-left: 6px solid #3182ce; margin-bottom: 1.5rem;">
                        {preview_html}
                    </div>
                    """,
                    unsafe_allow_html=True
                )
                
            col_b1, col_b2 = st.columns(2)
            with col_b1:
                if st.button("💾 Змінити конфіг", type="primary", use_container_width=True):
                    full_data = load_excel_config(config_path)
                    if full_data and sheet_name in full_data:
                        full_data[sheet_name]["rows"][row_idx] = edited_vars
                        success = save_excel_config(
                            config_path,
                            sheet_name,
                            full_data[sheet_name]["template_path"],
                            full_data[sheet_name]["name_pattern"],
                            full_data[sheet_name]["headers"],
                            full_data[sheet_name]["rows"]
                        )
                        if success:
                            if config_path in st.session_state["pm_cached_configs"]:
                                del st.session_state["pm_cached_configs"][config_path]
                            st.toast("💾 Дані збережено в конфіг!", icon="💾")
                            time.sleep(0.5)
                            st.rerun()
            with col_b2:
                if st.button("📄 Згенерувати файл", use_container_width=True):
                    save_generated_document_dialog(actual_template_path, edited_vars, config_path, name_pattern=name_pattern)
                    
            pending_download = st.session_state.get("pending_download")
            if pending_download and isinstance(pending_download, dict):
                st.write(" ")
                st.download_button(
                    label="⬇️ Скачати документ через браузер",
                    data=pending_download["bytes"],
                    file_name=pending_download["name"],
                    mime=pending_download["mime"],
                    key="pm_download_fallback_btn",
                    use_container_width=True
                )

            st.write(" ")
            with st.expander("✏️ Змінні документа", expanded=False):
                var_keys = [h for h in headers if h]
                var_vals = [edited_vars.get(h, "") for h in var_keys]
                
                doc_vars_df = pd.DataFrame({
                    "Змінна": var_keys,
                    "Значення": var_vals
                })
                
                clean_cfg_path = "".join([c if c.isalnum() else "_" for c in config_path])
                edited_doc_vars_df = st.data_editor(
                    doc_vars_df,
                    column_config={
                        "Змінна": st.column_config.TextColumn(disabled=True),
                        "Значення": st.column_config.TextColumn(disabled=False)
                    },
                    hide_index=True,
                    width="stretch",
                    key=f"doc_vars_data_editor_{clean_cfg_path}_{sheet_name}_{row_idx}"
                )
                
                has_changes = False
                for _, row in edited_doc_vars_df.iterrows():
                    k = row["Змінна"]
                    v = str(row["Значення"])
                    if edited_vars.get(k, "") != v:
                        edited_vars[k] = v
                        has_changes = True
                        
                if has_changes:
                    save_persistent_state()
                    st.rerun()

    # --- MODE D: TEMPLATE SELECTED (PLACEHOLDERS EXTRACTION & GLOBAL RENAME) ---
    elif sel_type == "template":
        selected_template_path = st.session_state.get("selected_template_path")
        if not selected_template_path:
            st.info("👈 Оберіть шаблон у дереві ліворуч.")
            return
            
        actual_t_path = os.path.abspath(selected_template_path)
        
        with col_main:
            st.header("Управління шаблоном")
            st.caption(f"**Шлях до файлу:** `{selected_template_path}`")
            
            if st.button("🖥️ Відкрити шаблон", key="btn_open_template_top", use_container_width=False):
                try:
                    os.startfile(actual_t_path)
                    st.toast("🖥️ Шаблон відчинено у зовнішньому додатку!", icon="🖥️")
                except Exception as e:
                    st.error(f"Не вдалося відкрити файл: {e}")
                    
            if not os.path.exists(actual_t_path):
                st.error("Файл шаблону не знайдено!")
            else:
                st.markdown("---")
                st.subheader("👁️ Швидкий перегляд шаблону")
                
                ext = os.path.splitext(actual_t_path)[1].lower()
                with st.spinner("⏳ Генерація прев'ю..."):
                    if ext == ".docx":
                        preview_html = generate_docx_preview(actual_t_path, {}, config_path="")
                    elif ext == ".xlsx":
                        preview_html = generate_xlsx_preview(actual_t_path, {}, config_path="")
                    else:
                        preview_html = f"<div style='color: #e53e3e;'>Непідтримуваний тип: {ext}</div>"
                        
                st.markdown(
                    f"""
                    <div class="document-preview-container" style="border: 2px solid #3182ce; border-radius: 8px; padding: 20px; background-color: #ffffff; max-height: 500px; overflow-y: auto; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1); border-left: 6px solid #3182ce; margin-bottom: 1.5rem;">
                        {preview_html}
                    </div>
                    """,
                    unsafe_allow_html=True
                )
                with st.spinner("⏳ Вилучення змінних з шаблону..."):
                    placeholders = extract_placeholders_with_context(actual_t_path)
                    
                if not placeholders:
                    st.info("У цьому шаблоні не знайдено змінних у форматі {{ змінна }}.")
                else:
                    with st.expander("✏️ Редагувати назви змінних (шаблон та конфіг)", expanded=False):
                        st.caption("Подвійний клік на осередки правої колонки для перейменування змінної у шаблоні та пов'язаному конфігу.")
                        tpl_vars = list(placeholders.keys())
                        tpl_df = pd.DataFrame({
                            "Поточне ім'я змінної": tpl_vars,
                            "Нова назва змінної": tpl_vars
                        })
                        
                        edited_tpl_df = st.data_editor(
                            tpl_df,
                            column_config={
                                "Поточне ім'я змінної": st.column_config.TextColumn(disabled=True),
                                "Нова назва змінної": st.column_config.TextColumn(disabled=False)
                            },
                            hide_index=True,
                            width="stretch",
                            key="tpl_rename_data_editor"
                        )
                        
                        if st.button("💾 Застосувати перейменування", type="primary", use_container_width=True):
                            renamed_count = 0
                            renamed_configs_count = 0
                            with st.spinner("⏳ Перейменування в шаблоні та конфігах..."):
                                referencing_configs = []
                                try:
                                    cfg_path = st.session_state.get("last_opened_config")
                                    if cfg_path and os.path.exists(cfg_path):
                                        sheets_data = load_excel_config(cfg_path)
                                        if sheets_data:
                                            for sheet_name, info in sheets_data.items():
                                                tpl_path = info.get("template_path", "")
                                                if tpl_path:
                                                    cfg_dir = os.path.dirname(os.path.abspath(cfg_path))
                                                    sheet_actual_t_path = resolve_path(cfg_dir, tpl_path)
                                                    if os.path.abspath(sheet_actual_t_path) == actual_t_path:
                                                        referencing_configs.append((cfg_path, sheet_name, info))
                                except Exception as e_scan:
                                    st.warning(f"Не вдалося перевірити конфіг на використання шаблону: {e_scan}")

                                modified_cfgs = set()
                                for _, row in edited_tpl_df.iterrows():
                                    old_n = row["Поточне ім'я змінної"]
                                    new_n = str(row["Нова назва змінної"]).strip()
                                    if new_n and new_n != old_n and " " not in new_n:
                                        if rename_placeholder_in_template(actual_t_path, old_n, new_n):
                                            renamed_count += 1
                                            for cfg_path, sheet_name, info in referencing_configs:
                                                if old_n in info["headers"]:
                                                    h_idx = info["headers"].index(old_n)
                                                    info["headers"][h_idx] = new_n
                                                    for r in info["rows"]:
                                                        if old_n in r:
                                                            r[new_n] = r.pop(old_n)
                                                    
                                                    # Also rename in template_path and name_pattern of referencing configs
                                                    import re
                                                    pat = re.compile(r'(\{\{\s*)' + re.escape(old_n) + r'(\s*\}\})')
                                                    if info.get("template_path"):
                                                        info["template_path"] = pat.sub(lambda m: m.group(1) + new_n + m.group(2), info["template_path"])
                                                    if info.get("name_pattern"):
                                                        info["name_pattern"] = pat.sub(lambda m: m.group(1) + new_n + m.group(2), info["name_pattern"])
                                                            
                                                    modified_cfgs.add((cfg_path, sheet_name))

                                if modified_cfgs:
                                    saved_cfg_paths = set()
                                    for cfg_path, sheet_name in modified_cfgs:
                                        for c_p, s_n, info in referencing_configs:
                                            if c_p == cfg_path and s_n == sheet_name:
                                                save_excel_config(
                                                    cfg_path,
                                                    sheet_name,
                                                    info["template_path"],
                                                    info["name_pattern"],
                                                    info["headers"],
                                                    info["rows"]
                                                )
                                                saved_cfg_paths.add(cfg_path)
                                    for c_p in saved_cfg_paths:
                                        if "pm_cached_configs" in st.session_state:
                                            st.session_state["pm_cached_configs"].pop(c_p, None)
                                    renamed_configs_count = len(saved_cfg_paths)

                            if renamed_count > 0:
                                msg = f"✅ Перейменовано {renamed_count} змінних у шаблоні!"
                                if renamed_configs_count > 0:
                                    msg += f" Оновлено {renamed_configs_count} конфіг(ів)!"
                                st.toast(msg, icon="💾")
                                time.sleep(0.5)
                                st.rerun()
                            else:
                                st.warning("Немає нових (валідних) назв для перейменування.")
                                
                    st.write(" ")
                    with st.expander("🔍 Виявлені плейсхолдери та їх контекст", expanded=False):
                        p_data = [{"Змінна": k, "Контекст використання": v} for k, v in placeholders.items()]
                        st.table(pd.DataFrame(p_data))

    # --- MODE D: GLOBAL VARIABLES ANALYZER ---
    elif sel_type == "all_variables":
        with col_main:
            render_variables_analyzer(pm_path)
