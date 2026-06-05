import streamlit as st
import os
from ui.state_manager import (
    save_persistent_state,
    clear_pm_input_keys,
    get_cached_config
)
from core.text_processor import resolve_virtual_doc_name, resolve_path
from core.io_utils import build_docs_only_tree

def scan_recursive_templates(root_folder):
    """Walks the folder structure recursively and returns a list of template files (template_*)."""
    template_files = []
    if not os.path.exists(root_folder) or not os.path.isdir(root_folder):
        return []
    for dirpath, _, filenames in os.walk(root_folder):
        for f in filenames:
            if (f.endswith('.docx') or f.endswith('.xlsx')) and f.startswith('template_') and not f.startswith('~$'):
                full_path = os.path.abspath(os.path.join(dirpath, f))
                template_files.append(full_path)
    return sorted(template_files)

def build_path_tree(paths, root_path):
    tree = {}
    for path in paths:
        rel_path = os.path.relpath(path, root_path)
        parts = rel_path.split(os.sep)
        current = tree
        for part in parts[:-1]:
            if part not in current:
                current[part] = {}
            current = current[part]
        current[parts[-1]] = path
    return tree

def contains_template_path(tree_dict, target_path):
    if not target_path: return False
    try: target_path = os.path.abspath(target_path).lower()
    except Exception: return False
    def check_node(val):
        if isinstance(val, dict): return any(check_node(v) for v in val.values())
        elif isinstance(val, str):
            try: return os.path.abspath(val).lower() == target_path
            except Exception: return False
        return False
    return check_node(tree_dict)

def render_template_tree_node(node_name, node_value, depth=0, current_path=""):
    selected_template = st.session_state.get("selected_template_path")
    if isinstance(node_value, dict):
        should_expand = False
        if selected_template and contains_template_path(node_value, selected_template):
            should_expand = True
        
        folder_abs_path = os.path.abspath(os.path.join(current_path, node_name))
        clean_path_key = "".join([c if c.isalnum() else "_" for c in folder_abs_path])
        exp_key = f"tpl_folder_exp_{clean_path_key}"
        
        with st.expander("📁 " + node_name, expanded=should_expand, key=exp_key):
            for name, val in node_value.items():
                render_template_tree_node(name, val, depth + 1, current_path=folder_abs_path)
    else:
        t_path = node_value
        is_selected = (
            st.session_state.get("active_selection_type") == "template"
            and selected_template
            and os.path.abspath(selected_template).lower() == os.path.abspath(t_path).lower()
        )
        btn_type = "primary" if is_selected else "secondary"
        clean_t_key = "".join([c if c.isalnum() else "_" for c in t_path])
        if st.button(
            f"📄 {node_name}",
            key=f"btn_tpl_{clean_t_key}",
            use_container_width=True,
            type=btn_type
        ):
            st.session_state["active_selection_type"] = "template"
            st.session_state["selected_template_path"] = t_path
            st.session_state["pm_selected_doc"] = None
            st.session_state["editor_config_path"] = None
            st.session_state["editor_selected_sheet"] = None
            st.session_state["selected_folder_path"] = None
            st.session_state["last_opened_template"] = t_path
            save_persistent_state()
            st.rerun()

def render_templates_list(root_path):
    """Renders a tree of scanned templates in the navigation panel."""
    templates = scan_recursive_templates(root_path)
    if not templates:
        return
    
    selected_template = st.session_state.get("selected_template_path")
    has_active_template = st.session_state.get("active_selection_type") == "template" and selected_template is not None
    
    with st.expander("📝 Шаблони документів", expanded=bool(has_active_template), key="templates_list_exp"):
        tree = build_path_tree(templates, root_path)
        for name, val in tree.items():
            render_template_tree_node(name, val, 0, current_path=root_path)

def contains_config_path(tree_dict, target_path):
    if not target_path:
        return False
    try:
        target_path = os.path.abspath(target_path).lower()
    except Exception:
        return False
    
    def check_node(val):
        if isinstance(val, dict):
            return any(check_node(v) for v in val.values())
        elif isinstance(val, str):
            try:
                return os.path.abspath(val).lower() == target_path
            except Exception:
                return False
        return False
        
    return check_node(tree_dict)

def render_docs_only_tree(tree, depth=0, path_prefix=""):
    selected_doc = st.session_state.get("pm_selected_doc")
    target_config = selected_doc.get("config_path") if selected_doc else None
    
    # 1. Render subfolders
    for key, value in tree.items():
        if key == "__docs__":
            continue
        def contains_selected(folder_dict):
            if "__docs__" in folder_dict:
                for doc in folder_dict["__docs__"]:
                    if target_config and os.path.abspath(doc["config_path"]).lower() == os.path.abspath(target_config).lower():
                        if selected_doc.get("sheet_name") == doc["sheet_name"] and selected_doc.get("row_idx") == doc["row_idx"]:
                            return True
            for k, v in folder_dict.items():
                if k != "__docs__" and isinstance(v, dict):
                    if contains_selected(v):
                        return True
            return False
            
        should_expand = bool(selected_doc and contains_selected(value)) or st.session_state.get("force_expand_tree", False)
        exp_key = f"pm_docs_only_exp_{path_prefix}_{key}"
        with st.expander("📁 " + key, expanded=should_expand, key=exp_key):
            render_docs_only_tree(value, depth + 1, path_prefix=f"{path_prefix}_{key}")
            
    # 2. Render files (documents) in the current folder
    if "__docs__" in tree:
        for doc in tree["__docs__"]:
            is_selected = (
                st.session_state.get("active_selection_type") == "document" and
                selected_doc and
                selected_doc.get("config_path") == doc["config_path"] and
                selected_doc.get("sheet_name") == doc["sheet_name"] and
                selected_doc.get("row_idx") == doc["row_idx"]
            )
            button_type = "primary" if is_selected else "secondary"
            if st.button(
                f"📄 {os.path.basename(doc['doc_name'])}",
                key=f"pm_doc_btn_{doc['config_path']}_{doc['sheet_name']}_{doc['row_idx']}",
                use_container_width=True,
                type=button_type
            ):
                st.session_state["active_selection_type"] = "document"
                st.session_state["pm_selected_doc"] = doc
                st.session_state["last_opened_config"] = doc["config_path"]
                st.session_state["last_opened_folder"] = os.path.dirname(os.path.abspath(doc["config_path"]))
                st.session_state["last_opened_template"] = doc["template_path"]
                st.session_state["pending_download"] = None
                st.session_state["force_expand_tree"] = False
                clear_pm_input_keys()
                save_persistent_state()
                st.session_state["pm_editing_vars"] = None
                save_persistent_state()
                st.rerun()

def render_tree_node(node_name, node_value, depth=0, current_path=""):
    selected_doc = st.session_state.get("pm_selected_doc")
        
    target_path = None
    if selected_doc:
        target_path = selected_doc.get("config_path")
    elif st.session_state.get("active_selection_type") == "sheet":
        target_path = st.session_state.get("editor_config_path")
    
    if isinstance(node_value, dict):
        folder_abs_path = os.path.abspath(os.path.join(current_path, node_name))
        
        # Folder node: expand only if it contains the selected config file path or forced
        should_expand = bool(target_path and contains_config_path(node_value, target_path)) or st.session_state.get("force_expand_tree", False)
        clean_path_key = "".join([c if c.isalnum() else "_" for c in folder_abs_path])
        exp_key = f"pm_folder_exp_{clean_path_key}"
        
        with st.expander("📁 " + node_name, expanded=should_expand, key=exp_key):

                
            for name, val in node_value.items():
                render_tree_node(name, val, depth + 1, current_path=folder_abs_path)
    else:
        # Config file node: expand if it is the current config of the selected document
        config_path = node_value
        config_name = node_name
        is_current_config = False
        if selected_doc and selected_doc.get("config_path"):
            try:
                is_current_config = os.path.abspath(selected_doc["config_path"]).lower() == os.path.abspath(config_path).lower()
            except Exception:
                pass
        
        is_active_config = (
            st.session_state.get("active_selection_type") == "sheet"
            and st.session_state.get("editor_config_path") == config_path
        )
        should_expand_config = is_current_config or is_active_config or st.session_state.get("force_expand_tree", False)
        
        clean_cfg = "".join([c if c.isalnum() else "_" for c in config_path])
        exp_cfg_key = f"pm_cfg_exp_{clean_cfg}"
        with st.expander("📊 " + config_name, expanded=should_expand_config, key=exp_cfg_key):
            # Button to edit the entire config
            config_btn_type = "primary" if is_active_config else "secondary"
            if st.button(
                "⚙️ Конфіг",
                key=f"sel_cfg_btn_{clean_cfg}",
                use_container_width=True,
                type=config_btn_type
            ):
                st.session_state["active_selection_type"] = "sheet"
                st.session_state["editor_config_path"] = config_path
                st.session_state["editor_selected_sheet"] = None
                st.session_state["pm_selected_doc"] = None
                st.session_state["selected_template_path"] = None
                st.session_state["selected_folder_path"] = None
                st.session_state["last_opened_config"] = config_path
                st.session_state["force_expand_tree"] = False
                st.session_state["last_opened_folder"] = os.path.dirname(os.path.abspath(config_path))
                save_persistent_state()
                st.rerun()

            sheets_data = get_cached_config(config_path)
            if not sheets_data:
                st.caption("Не вдалося завантажити або порожній конфіг")
                return
            
            for sheet_name, info in sheets_data.items():
                rows = info["rows"]
                template_path = info["template_path"]
                name_pattern = info["name_pattern"]
                
                # Sheet expander: expand if it is the current sheet of the selected document
                is_current_sheet = is_current_config and selected_doc.get("sheet_name") == sheet_name
                is_active_sheet = (
                    st.session_state.get("active_selection_type") == "sheet"
                    and st.session_state.get("editor_config_path") == config_path
                    and st.session_state.get("editor_selected_sheet") == sheet_name
                )
                should_expand_sheet = is_current_sheet or is_active_sheet
                
                exp_sheet_key = f"pm_sheet_exp_{clean_cfg}_{sheet_name}"
                with st.expander(f"📋 {sheet_name} ({len(rows)} док.)", expanded=bool(should_expand_sheet), key=exp_sheet_key):
                    # Button to edit/select template for this sheet
                    cfg_dir = os.path.dirname(os.path.abspath(config_path))
                    actual_template_path = resolve_path(cfg_dir, template_path) if template_path else None
                    
                    is_active_template = (
                        st.session_state.get("active_selection_type") == "template"
                        and st.session_state.get("selected_template_path")
                        and actual_template_path
                        and os.path.abspath(st.session_state.get("selected_template_path")).lower() == os.path.abspath(actual_template_path).lower()
                    )
                    
                    if actual_template_path and os.path.exists(actual_template_path):
                        template_btn_type = "primary" if is_active_template else "secondary"
                        if st.button(
                            "📝 Шаблон",
                            key=f"sel_tpl_btn_{clean_cfg}_{sheet_name}",
                            use_container_width=True,
                            type=template_btn_type
                        ):
                            st.session_state["active_selection_type"] = "template"
                            st.session_state["selected_template_path"] = actual_template_path
                            st.session_state["editor_config_path"] = None
                            st.session_state["editor_selected_sheet"] = None
                            st.session_state["pm_selected_doc"] = None
                            st.session_state["selected_folder_path"] = None
                            st.session_state["last_opened_template"] = actual_template_path
                            st.session_state["last_opened_config"] = config_path
                            st.session_state["force_expand_tree"] = False
                            save_persistent_state()
                            st.rerun()
                    else:
                        st.markdown("<div style='color: #ff4b4b; font-size: 0.85rem; padding: 6px 10px; margin-bottom: 8px; border-left: 3px solid #ff4b4b; background-color: rgba(255, 75, 75, 0.1); border-radius: 0 4px 4px 0;'>⚠️ Шаблон не знайдено</div>", unsafe_allow_html=True)
                        
                    if not rows:
                        st.caption("Немає даних для генерації")
                        continue
                    
                    for idx, row in enumerate(rows):
                        is_selected = (
                            st.session_state.get("active_selection_type") == "document" and
                            isinstance(selected_doc, dict) and
                            selected_doc.get("config_path") == config_path and
                            selected_doc.get("sheet_name") == sheet_name and
                            selected_doc.get("row_idx") == idx
                        )
                        
                        current_row_vars = row
                        if is_selected and st.session_state.get("pm_editing_vars") is not None:
                            current_row_vars = st.session_state["pm_editing_vars"]
                            
                        doc_name = resolve_virtual_doc_name(name_pattern, current_row_vars, template_path)
                        if not doc_name.strip():
                            doc_name = f"document_{idx + 5}"
                            
                        button_type = "primary" if is_selected else "secondary"
                        
                        if st.button(
                            f"📄 {doc_name}",
                            key=f"pm_btn_{config_path}_{sheet_name}_{idx}",
                            use_container_width=True,
                            type=button_type
                        ):
                            st.session_state["active_selection_type"] = "document"
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
                            st.session_state["pending_download"] = None
                            st.session_state["force_expand_tree"] = False
                            clear_pm_input_keys()
                            save_persistent_state()
                            st.session_state["pm_editing_vars"] = None
                            save_persistent_state()
                            st.rerun()

