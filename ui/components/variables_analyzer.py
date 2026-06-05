import streamlit as st
import pandas as pd
import os
import time

from core.io_utils import save_excel_config, scan_recursive_configs
from ui.state_manager import get_cached_config
from core.text_processor import resolve_virtual_doc_name, resolve_path
from ui.components.config_table import rename_placeholder_in_template

def get_config_details_by_key(config_details, key):
    """Safely looks up config details by key, handling relative/absolute path mismatches."""
    if not key:
        return None
    if key in config_details:
        return config_details[key]
        
    normalized_key = key.replace("\\", "/").strip()
    for col_name, details in config_details.items():
        normalized_col = col_name.replace("\\", "/").strip()
        if normalized_col == normalized_key:
            return details
        if normalized_col.endswith(normalized_key) or normalized_key.endswith(normalized_col):
            return details
            
    return None

def build_variables_matrix(pm_path):
    """Будує зведену матрицю використання змінних по всіх конфігах."""
    config_files = scan_recursive_configs(pm_path)
    
    matrix_data = {} # var_name -> {col_name: usage_count}
    columns_list = []
    config_details = {} # Detailed info for Block B
    
    for c_path in config_files:
        sheets_data = get_cached_config(c_path)
        if not sheets_data: continue
            
        rel_c_path = os.path.relpath(c_path, pm_path).replace("\\", "/")
        
        for sheet_name, info in sheets_data.items():
            col_name = f"{rel_c_path} ➜ {sheet_name}"
            columns_list.append(col_name)
            
            headers = info.get("headers", [])
            rows = info.get("rows", [])
            
            config_details[col_name] = {
                "config_path": c_path,
                "sheet_name": sheet_name,
                "template_path": info.get("template_path", ""),
                "name_pattern": info.get("name_pattern", ""),
                "headers": headers,
                "rows": rows
            }
            
            for h in headers:
                h_str = str(h).strip()
                if not h_str: continue
                    
                usage_count = sum(1 for r in rows if r.get(h_str) and str(r.get(h_str)).strip())
                
                if h_str not in matrix_data:
                    matrix_data[h_str] = {}
                matrix_data[h_str][col_name] = usage_count
                
    # Build dataframe
    df_rows = []
    for var_name, usages in sorted(matrix_data.items()):
        row = {"Змінна": var_name}
        for col in columns_list:
            if col in usages:
                row[col] = usages[col]
            else:
                row[col] = pd.NA
        df_rows.append(row)
    df = pd.DataFrame(df_rows)
    return df, columns_list, config_details

def render_variables_analyzer(pm_path):
    st.header("🌐 Усі змінні (Модуль Аналізу)")
    st.write("Тут ви можете переглянути всі унікальні змінні проекту, їх значення в конфігах та здійснити крос-пошук і редагування.")

    with st.spinner("⏳ Аналіз конфігів..."):
        df, columns_list, config_details = build_variables_matrix(pm_path)
    
    if df.empty:
        st.info("У поточному проекті не знайдено жодних змінних.")
        return

    # Session-state tracking to keep selection stable during renaming or rerun
    if "va_last_selected_var" not in st.session_state:
        st.session_state["va_last_selected_var"] = None
    if "va_active_search_val" not in st.session_state:
        st.session_state["va_active_search_val"] = None

    selected_var = st.session_state["va_last_selected_var"]

    # Create side-by-side columns: left (Variable list) and right (Values)
    col_left, col_right = st.columns([1, 3])

    with col_left:
        st.markdown("##### 🔑 Змінні проекту")
        left_df = df[["Змінна"]]
        
        # Visually highlight the selected variable row in the left table
        def highlight_selected_var(row):
            if row["Змінна"] == selected_var:
                return ["background-color: rgba(49, 130, 206, 0.25); font-weight: bold; border: 1px solid #3182ce;"] * len(row)
            return [""] * len(row)
            
        try:
            styled_left_df = left_df.style.apply(highlight_selected_var, axis=1)
        except Exception:
            styled_left_df = left_df
            
        event_left = st.dataframe(
            styled_left_df,
            hide_index=True,
            width="stretch",
            height=435,
            on_select="rerun",
            selection_mode="single-cell",
            key="variables_left_selection"
        )
        
    # Process left selection
    selected_cells = getattr(event_left.selection, "cells", None) if hasattr(event_left, "selection") else None
    if selected_cells and len(selected_cells) > 0:
        row_idx, _ = selected_cells[0]
        if 0 <= row_idx < len(df):
            new_var = df.iloc[row_idx]["Змінна"]
            if new_var != selected_var:
                st.session_state["va_last_selected_var"] = new_var
                st.session_state["va_active_search_val"] = None
                st.rerun()

    search_val = None

    if selected_var:
        with col_right:
            st.markdown(f"##### 📊 Значення змінної у конфігах: `{selected_var}`")
            
            # Only include columns (sheets) that actually contain the selected variable
            relevant_columns = [col for col in columns_list if selected_var in config_details[col]["headers"]]
            
            max_rows = max((len(config_details[col]["rows"]) for col in relevant_columns), default=0)
            vals_data = []
            for r_idx in range(max_rows):
                row_dict = {}
                for col in relevant_columns:
                    rows = config_details[col]["rows"]
                    if r_idx < len(rows):
                        row_dict[col] = rows[r_idx].get(selected_var, "")
                    else:
                        row_dict[col] = None
                vals_data.append(row_dict)
                
            vals_df = pd.DataFrame(vals_data, columns=relevant_columns)
            
            # Highlight selected search value cell
            search_val = st.session_state.get("va_active_search_val")
            
            def highlight_search_val(val):
                if pd.notna(val) and str(val).strip() == search_val:
                    return "background-color: rgba(49, 130, 206, 0.25); font-weight: bold; border: 1px solid #3182ce;"
                return ""
                
            try:
                styled_vals_df = vals_df.style.map(highlight_search_val)
            except AttributeError:
                try:
                    styled_vals_df = vals_df.style.applymap(highlight_search_val)
                except Exception:
                    styled_vals_df = vals_df
            except Exception:
                styled_vals_df = vals_df

            event_vals = st.dataframe(
                styled_vals_df,
                hide_index=True,
                width="stretch",
                height=385,
                selection_mode="single-cell",
                on_select="rerun",
                key=f"va_values_selection_{selected_var}"
            )
            
            selected_cells_vals = getattr(event_vals.selection, "cells", None) if hasattr(event_vals, "selection") else None
            col_name_selected = None
            if selected_cells_vals and len(selected_cells_vals) > 0:
                row_idx, col_idx = selected_cells_vals[0]
                if 0 <= row_idx < len(vals_df):
                    if isinstance(col_idx, int):
                        col_name_selected = vals_df.columns[col_idx] if 0 <= col_idx < len(vals_df.columns) else None
                    else:
                        col_name_selected = col_idx if col_idx in vals_df.columns else None
                        
                    if col_name_selected:
                        val = str(vals_df.iloc[row_idx][col_name_selected]).strip()
                        if val and val != "None" and val != "<NA>" and pd.notna(val):
                            if val != search_val:
                                st.session_state["va_active_search_val"] = val
                                st.rerun()

            search_val = st.session_state["va_active_search_val"]

            # Render Open Config / Open Template buttons based on column
            btn_col1, btn_col2 = st.columns(2)
            
            if col_name_selected and col_name_selected in config_details:
                details = config_details[col_name_selected]
                config_path = details["config_path"]
                template_path = details["template_path"]
                
                cfg_dir = os.path.dirname(os.path.abspath(config_path))
                actual_t_path = resolve_path(cfg_dir, template_path) if template_path else ""
                
                with btn_col1:
                    if st.button("📂 Відкрити конфіг", key=f"btn_va_open_config_{selected_var}", use_container_width=True):
                        try:
                            if os.path.exists(config_path):
                                os.startfile(config_path)
                                st.toast("🖥️ Конфіг відчинено у зовнішньому додатку!", icon="🖥️")
                            else:
                                st.error("Файл конфігу не знайдено!")
                        except Exception as e:
                            st.error(f"Не вдалося відкрити конфіг: {e}")
                            
                with btn_col2:
                    if st.button("📄 Відкрити шаблон", key=f"btn_va_open_template_{selected_var}", use_container_width=True, disabled=not actual_t_path):
                        try:
                            if actual_t_path and os.path.exists(actual_t_path):
                                os.startfile(actual_t_path)
                                st.toast("🖥️ Шаблон відчинено у зовнішньому додатку!", icon="🖥️")
                            else:
                                st.error("Файл шаблону не знайдено!")
                        except Exception as e:
                            st.error(f"Не вдалося відкрити шаблон: {e}")
            else:
                with btn_col1:
                    st.button("📂 Відкрити конфіг", key=f"btn_va_open_config_disabled_{selected_var}", use_container_width=True, disabled=True, help="Оберіть комірку в таблиці значень для активації дії")
                with btn_col2:
                    st.button("📄 Відкрити шаблон", key=f"btn_va_open_template_disabled_{selected_var}", use_container_width=True, disabled=True, help="Оберіть комірку в таблиці значень для активації дії")

            if not search_val:
                st.info("👆 Клікніть на будь-яке значення змінної у таблиці вище для виклику крос-пошуку.")
    else:
        with col_right:
            st.info("👈 Оберіть змінну у списку ліворуч для перегляду її значень.")

    # Render Cross-Search outside of columns layout (full width)
    if selected_var and search_val:
        st.markdown("---")
        


        st.markdown(f"##### 🔍 Крос-пошук для значення: `{search_val}`")
        
        search_results = []
        for other_cfg_name, other_details in config_details.items():
            for other_r_idx, other_r in enumerate(other_details["rows"]):
                doc_name_pattern = other_details["name_pattern"]
                
                doc_name_resolved = resolve_virtual_doc_name(
                    other_details["name_pattern"],
                    other_r,
                    other_details["template_path"]
                )
                if not doc_name_resolved.strip():
                    doc_name_resolved = f"Документ {other_r_idx + 1}"
                    
                for k, v in other_r.items():
                    if str(v).strip() == search_val:
                        row_data = {
                            "Конфіг (аркуш)": other_cfg_name,
                            "Ім'я документа": doc_name_resolved,
                            "Шаблон імені документа": doc_name_pattern,
                            "Змінна зі збігом": k,
                            "_old_var": k,
                            "_r_idx": other_r_idx
                        }
                        for other_k, other_v in other_r.items():
                            if not other_k.startswith("_") and not other_k.endswith("_type_code") and other_k not in row_data:
                                row_data[other_k] = str(other_v) if pd.notna(other_v) else ""
                        search_results.append(row_data)

        if not search_results:
            st.info("Більше збігів не знайдено.")
        else:
            res_df = pd.DataFrame(search_results)
            reset_counter = st.session_state.get("va_editor_reset_counter", 0)
            editor_key = f"va_search_editor4_{selected_var}_{search_val}_{reset_counter}"
            editor_key_clean = "".join([c if c.isalnum() else "_" for c in editor_key])

            # 1. Process any incoming edits from the data_editor widget and write them immediately to the in-memory cache
            editor_state = st.session_state.get(editor_key_clean)
            if editor_state and "edited_rows" in editor_state and editor_state["edited_rows"]:
                edited_rows = editor_state["edited_rows"]
                
                # Initialize persistent modification tracking keys
                if "va_modified_configs" not in st.session_state:
                    st.session_state["va_modified_configs"] = set()
                if "va_template_renames" not in st.session_state:
                    st.session_state["va_template_renames"] = []
                    
                # Pre-pass: collect variable name renames by config sheet
                sheet_renames = {}
                for row_idx_str, edits in edited_rows.items():
                    row_idx = int(row_idx_str)
                    if row_idx < len(res_df):
                        original_row = res_df.iloc[row_idx]
                        cfg = original_row["Конфіг (аркуш)"]
                        old_var = original_row["_old_var"]
                        val_var = edits.get("Змінна зі збігом")
                        new_var = val_var.strip() if val_var is not None else old_var
                        
                        if new_var and new_var != old_var:
                            if cfg not in sheet_renames:
                                sheet_renames[cfg] = {}
                            sheet_renames[cfg][old_var] = new_var
                            if old_var == st.session_state.get("va_last_selected_var"):
                                st.session_state["va_last_selected_var"] = new_var

                # Validate renames first
                error_msg = None
                duplicate_conflict = None
                for cfg, renames in sheet_renames.items():
                    details = get_config_details_by_key(config_details, cfg)
                    if not details:
                        continue
                    for old_v, new_v in renames.items():
                        if not new_v:
                            error_msg = "Помилка: ім'я змінної не може бути порожнім!"
                            break
                        if new_v in details["headers"] and new_v != old_v:
                            duplicate_conflict = {
                                "cfg": cfg,
                                "old_var": old_v,
                                "new_var": new_v
                            }
                            break
                    if error_msg or duplicate_conflict:
                        break
                        
                if error_msg:
                    st.error(error_msg)
                    st.session_state["va_editor_reset_counter"] = st.session_state.get("va_editor_reset_counter", 0) + 1
                    if editor_key_clean in st.session_state:
                        del st.session_state[editor_key_clean]
                    time.sleep(2)
                    st.rerun()
                    
                if duplicate_conflict:
                    st.session_state["va_pending_duplicate_conflict"] = duplicate_conflict
                    editor_state["edited_rows"] = {}
                    st.rerun()

                # Apply variable name renames to in-memory cache
                for cfg, renames in sheet_renames.items():
                    details = get_config_details_by_key(config_details, cfg)
                    if not details:
                        continue
                    for old_v, new_v in renames.items():
                        if old_v in details["headers"]:
                            h_idx = details["headers"].index(old_v)
                            details["headers"][h_idx] = new_v
                        for r in details["rows"]:
                            if old_v in r:
                                r[new_v] = r.pop(old_v)
                                
                        # Rename in A1/A2 metadata
                        import re
                        pat = re.compile(r'(\{\{\s*)' + re.escape(old_v) + r'(\s*\}\})')
                        if details.get("template_path"):
                            details["template_path"] = pat.sub(lambda m: m.group(1) + new_v + m.group(2), details["template_path"])
                        if details.get("name_pattern"):
                            details["name_pattern"] = pat.sub(lambda m: m.group(1) + new_v + m.group(2), details["name_pattern"])
                            
                        # Update in the actual cache to make it persist across reruns
                        cfg_path = details["config_path"]
                        sheet_name = details["sheet_name"]
                        if "pm_cached_configs" in st.session_state and cfg_path in st.session_state["pm_cached_configs"]:
                            cached_sheet = st.session_state["pm_cached_configs"][cfg_path]["data"].get(sheet_name)
                            if cached_sheet:
                                if cached_sheet.get("template_path"):
                                    cached_sheet["template_path"] = pat.sub(lambda m: m.group(1) + new_v + m.group(2), cached_sheet["template_path"])
                                if cached_sheet.get("name_pattern"):
                                    cached_sheet["name_pattern"] = pat.sub(lambda m: m.group(1) + new_v + m.group(2), cached_sheet["name_pattern"])
                                
                        st.session_state["va_modified_configs"].add(cfg)
                        st.session_state["va_template_renames"].append((old_v, new_v, details["config_path"], details["sheet_name"]))

                # Main pass: update patterns and values in the in-memory cache
                for row_idx_str, edits in edited_rows.items():
                    row_idx = int(row_idx_str)
                    if row_idx < len(res_df):
                        original_row = res_df.iloc[row_idx]
                        cfg = original_row["Конфіг (аркуш)"]
                        r_idx = original_row["_r_idx"]
                        details = get_config_details_by_key(config_details, cfg)
                        if not details:
                            continue
                        old_var = original_row["_old_var"]
                        current_var_renames = sheet_renames.get(cfg, {})

                        # Update naming pattern
                        if "Шаблон імені документа" in edits:
                            new_pattern = edits["Шаблон імені документа"].strip()
                            if new_pattern != details["name_pattern"]:
                                details["name_pattern"] = new_pattern
                                
                                # Update in the actual cache to make it persist across reruns
                                cfg_path = details["config_path"]
                                sheet_name = details["sheet_name"]
                                if "pm_cached_configs" in st.session_state and cfg_path in st.session_state["pm_cached_configs"]:
                                    cached_sheet = st.session_state["pm_cached_configs"][cfg_path]["data"].get(sheet_name)
                                    if cached_sheet:
                                        cached_sheet["name_pattern"] = new_pattern
                                        
                                st.session_state["va_modified_configs"].add(cfg)

                        # Update cell values
                        for col_name, new_val in edits.items():
                            if col_name not in ["Конфіг (аркуш)", "Ім'я документа", "Шаблон імені документа", "Змінна зі збігом", "_old_var", "_r_idx"]:
                                actual_col_name = col_name
                                if col_name in current_var_renames:
                                    actual_col_name = current_var_renames[col_name]
                                elif col_name == old_var and old_var in current_var_renames:
                                    actual_col_name = current_var_renames[old_var]

                                if actual_col_name in details["rows"][r_idx]:
                                    details["rows"][r_idx][actual_col_name] = str(new_val).strip()
                                    st.session_state["va_modified_configs"].add(cfg)

                # Clear edited rows state and rerun so the UI is updated from the cache
                editor_state["edited_rows"] = {}
                st.rerun()

            col_config = {
                "Конфіг (аркуш)": st.column_config.TextColumn(disabled=True),
                "Ім'я документа": st.column_config.TextColumn(disabled=True),
                "Шаблон імені документа": st.column_config.TextColumn(disabled=False),
                "Змінна зі збігом": st.column_config.TextColumn(disabled=False),
                "_old_var": None,
                "_r_idx": None
            }
            for col in res_df.columns:
                if col not in col_config:
                    col_config[col] = st.column_config.TextColumn(disabled=False)

            # Visually highlight matching variable name cells in the cross-search results
            def highlight_match(row):
                styles = [""] * len(row)
                try:
                    col_idx = list(row.index).index("Змінна зі збігом")
                    styles[col_idx] = "background-color: rgba(49, 130, 206, 0.25); font-weight: bold; border: 1px solid #3182ce;"
                except Exception:
                    pass
                return styles
                
            try:
                styled_res_df = res_df.style.apply(highlight_match, axis=1)
            except Exception:
                styled_res_df = res_df

            edited_search_df = st.data_editor(
                styled_res_df,
                column_config=col_config,
                hide_index=True,
                width="stretch",
                height=400,
                key=editor_key_clean
            )

            st.caption("💡 При редагуванні **імен змінних** у стовпчику «Змінна зі збігом» зміни автоматично вносяться як до відповідних аркушів конфігів, так і безпосередньо у файли Word/Excel шаблонів (заміна Jinja-тегів). Також ви можете редагувати **шаблони імені документа** та **значення змінних** в інших стовпчиках.")

    # 2. Conflict Resolution UI (shown globally at the bottom of the page if there is an unresolved conflict)
    conflict = st.session_state.get("va_pending_duplicate_conflict")
    if conflict:
        st.markdown("---")
        st.markdown(f"##### ⚠️ Конфлікт перейменування: `{conflict['old_var']}` ➜ `{conflict['new_var']}`")
        st.error(
            f"**Помилка: змінна `{conflict['new_var']}` вже існує в конфігу `{conflict['cfg']}`!**\n\n"
            f"Ви намагаєтеся перейменувати змінну `{conflict['old_var']}` у `{conflict['new_var']}`. "
            f"Оскільки змінна з таким ім'ям вже є в цьому аркуші конфігу, виберіть одну з наступних дій:"
        )
        st.markdown(
            f"""
            ### 📋 Доступні варіанти вирішення:
            
            1. **❌ Скасувати зміни**
               * Жодних змін до конфігу чи шаблону не вноситься.
               * Повернення до початкового стану.
               
            2. **⚠️ Злити змінні та видалити стару**
               * **Увага (Втрата даних!):** Колонка `{conflict['new_var']}` буде повністю видалена з аркуша `{conflict['cfg']}`. **Усі поточні значення цієї змінної для всіх рядків конфігу буде безповоротно втрачено!**
               * Колонку змінної `{conflict['old_var']}` буде перейменовано на `{conflict['new_var']}` у конфігу.
               * У підключеному шаблоні існуюча змінна `{conflict['new_var']}` **залишається** без змін, а згадки колишньої змінної `{conflict['old_var']}` будуть **перейменовані** на `{conflict['new_var']}`.
               * Таким чином, шаблон посилатиметься на нову об'єднану змінну `{conflict['new_var']}` в обох місцях.
            """
        )
        col_c1, col_c2 = st.columns(2)
        with col_c1:
            if st.button("❌ Скасувати зміни", key="btn_cancel_conflict", use_container_width=True):
                st.session_state["va_pending_duplicate_conflict"] = None
                st.session_state["va_editor_reset_counter"] = st.session_state.get("va_editor_reset_counter", 0) + 1
                st.rerun()
        with col_c2:
            if st.button("⚠️ Злити змінні та видалити стару", key="btn_confirm_conflict", type="primary", use_container_width=True):
                cfg = conflict["cfg"]
                old_var = conflict["old_var"]
                new_var = conflict["new_var"]
                
                details = get_config_details_by_key(config_details, cfg)
                if details:
                    # 1. Delete old existing new_var from config headers & rows (deleting its data)
                    if new_var in details["headers"]:
                        details["headers"].remove(new_var)
                    for r in details["rows"]:
                        r.pop(new_var, None)
                        
                    # 2. Rename old_var to new_var in config headers & rows
                    if old_var in details["headers"]:
                        h_idx = details["headers"].index(old_var)
                        details["headers"][h_idx] = new_var
                    for r in details["rows"]:
                        if old_var in r:
                            r[new_var] = r.pop(old_var)
                            
                    # Rename in A1/A2 metadata
                    import re
                    pat = re.compile(r'(\{\{\s*)' + re.escape(old_var) + r'(\s*\}\})')
                    if details.get("template_path"):
                        details["template_path"] = pat.sub(lambda m: m.group(1) + new_var + m.group(2), details["template_path"])
                    if details.get("name_pattern"):
                        details["name_pattern"] = pat.sub(lambda m: m.group(1) + new_var + m.group(2), details["name_pattern"])
                        
                    # Update in the actual cache to make it persist across reruns
                    cfg_path = details["config_path"]
                    sheet_name = details["sheet_name"]
                    if "pm_cached_configs" in st.session_state and cfg_path in st.session_state["pm_cached_configs"]:
                        cached_sheet = st.session_state["pm_cached_configs"][cfg_path]["data"].get(sheet_name)
                        if cached_sheet:
                            if cached_sheet.get("template_path"):
                                cached_sheet["template_path"] = pat.sub(lambda m: m.group(1) + new_var + m.group(2), cached_sheet["template_path"])
                            if cached_sheet.get("name_pattern"):
                                cached_sheet["name_pattern"] = pat.sub(lambda m: m.group(1) + new_var + m.group(2), cached_sheet["name_pattern"])
                            
                    # 3. Track modified configs & template renames
                    if "va_modified_configs" not in st.session_state:
                        st.session_state["va_modified_configs"] = set()
                    if "va_template_renames" not in st.session_state:
                        st.session_state["va_template_renames"] = []
                        
                    st.session_state["va_modified_configs"].add(cfg)
                    st.session_state["va_template_renames"].append((old_var, new_var, details["config_path"], details["sheet_name"]))
                    
                    if old_var == st.session_state.get("va_last_selected_var"):
                        st.session_state["va_last_selected_var"] = new_var
                        
                st.session_state["va_pending_duplicate_conflict"] = None
                st.success("🎉 Злиття успішно застосовано в пам'яті. Натисніть кнопку збереження нижче для запису на диск!")
                time.sleep(1)
                st.rerun()
        return

    # 3. Global Save Block (shown always at the bottom if there are unsaved changes, or if the third table is visible)
    has_unsaved = (
        "va_modified_configs" in st.session_state 
        and st.session_state["va_modified_configs"]
    )

    if (selected_var and search_val and locals().get("search_results")) or has_unsaved:
        st.markdown("---")
        if has_unsaved:
            st.warning("⚠️ У вас є незбережені зміни в пам'яті!")

        save_clicked = st.button(
            "💾 Зберегти зміни з таблиці (Крос-пошук)",
            type="primary",
            key="btn_save_cross_global",
            use_container_width=True,
            disabled=not has_unsaved
        )

        if save_clicked:
            with st.spinner("⏳ Збереження змін у конфігах та шаблонах..."):
                modified_cfgs = st.session_state.get("va_modified_configs", set())
                template_renames = st.session_state.get("va_template_renames", [])
                
                # Save all modified configs to disk
                for cfg in modified_cfgs:
                    details = get_config_details_by_key(config_details, cfg)
                    if details:
                        save_excel_config(
                            details["config_path"],
                            details["sheet_name"],
                            details["template_path"],
                            details["name_pattern"],
                            details["headers"],
                            details["rows"]
                        )
                        if "pm_cached_configs" in st.session_state:
                            st.session_state["pm_cached_configs"].pop(details["config_path"], None)

                # Process template renames
                renamed_templates_count = 0
                for item in template_renames:
                    if len(item) == 4:
                        old_v, new_v, cfg_path, s_name = item
                        details = next((d for c, d in config_details.items() if d["config_path"] == cfg_path and d["sheet_name"] == s_name), None)
                    else:
                        old_v, new_v, cfg_path = item
                        details = next((d for c, d in config_details.items() if d["config_path"] == cfg_path), None)
                    if details:
                        tpl_path = details.get("template_path", "")
                        if tpl_path:
                            cfg_dir = os.path.dirname(os.path.abspath(cfg_path))
                            actual_t_path = resolve_path(cfg_dir, tpl_path)
                            if os.path.exists(actual_t_path):
                                if rename_placeholder_in_template(actual_t_path, old_v, new_v):
                                    renamed_templates_count += 1
                                    
                # Clear tracking states
                st.session_state["va_modified_configs"] = set()
                st.session_state["va_template_renames"] = []
                
                st.success(f"🎉 Зміни успішно збережено в {len(modified_cfgs)} конфіг(ах)!")
                if renamed_templates_count > 0:
                    st.toast(f"✅ Також оновлено шаблонів: {renamed_templates_count}", icon="📄")
                time.sleep(0.5)
                st.rerun()
