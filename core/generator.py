import os
from core.io_utils import load_excel_config
from core.text_processor import get_now_vars, resolve_path, render_string_template, get_unique_path, clean_relative_path
from core.file_handlers.docx_handler import process_word
from core.file_handlers.xlsx_handler import process_excel

def run_generation(excel_file, sheet_selector="all", row_selector="all", custom_out_dir=None):
    excel_dir = os.path.dirname(os.path.abspath(excel_file))
    try:
        sheets_data = load_excel_config(excel_file)
    except Exception as e:
        print(f"Помилка при відкритті Excel: {e}")
        return
        
    now_vars = get_now_vars()
    for k in list(now_vars.keys()): 
        now_vars[f"{k}_type"] = "s"
        
    TYPE_MAP = {'s': 'string', 'n': 'number', 'd': 'date', 'b': 'boolean', 'f': 'formula', 'e': 'error'}
    
    sheets_to_process = []
    available_sheets = list(sheets_data.keys())
    
    if sheet_selector.lower() == "all":
        sheets_to_process = available_sheets
    else:
        try:
            idx = int(sheet_selector) - 1
            if 0 <= idx < len(available_sheets):
                sheets_to_process = [available_sheets[idx]]
        except ValueError:
            pass
        if not sheets_to_process:
            if sheet_selector in sheets_data:
                sheets_to_process = [sheet_selector]
        if not sheets_to_process:
            sel_clean = sheet_selector.strip().lower()
            for name in available_sheets:
                if name.strip().lower() == sel_clean:
                    sheets_to_process = [name]
                    break
                    
    if not sheets_to_process:
        print(f"Аркуш '{sheet_selector}' не знайдено.")
        return

    for sheet_name in sheets_to_process:
        print(f"Обробка аркуша: {sheet_name}")
        sheet_info = sheets_data[sheet_name]
        
        template_rel_path = sheet_info["template_path"]
        if not template_rel_path:
            print(f"  Пропуск: Не вказано шаблон у комірці A1.")
            continue
            
        template_path = resolve_path(excel_dir, template_rel_path)
        if not os.path.exists(template_path):
            print(f"  Помилка: Шаблон не знайдено: {template_path}")
            continue
            
        output_pattern = sheet_info["name_pattern"]
        if not output_pattern:
            print(f"  Пропуск: Не вказано шлях результату в A2.")
            continue
            
        headers = sheet_info["headers"]
        print(f"  Знайдені змінні: {', '.join([h for h in headers if h])}")
        
        rows = sheet_info["rows"]
        data_rows = []
        
        if row_selector.lower() == "all":
            for idx, r_dict in enumerate(rows):
                data_rows.append((5 + idx, r_dict))
        else:
            try:
                r_idx = int(row_selector)
                list_idx = r_idx - 5
                if 0 <= list_idx < len(rows):
                    data_rows.append((r_idx, rows[list_idx]))
                else:
                    print(f"  Рядок {r_idx} поза межами діапазону даних.")
            except ValueError:
                print(f"  Невірний формат номера рядка: {row_selector}")
                
        for r_num, row_dict in data_rows:
            variables = {**now_vars}
            for h in headers:
                if not h:
                    continue
                val = row_dict.get(h, "")
                v_type = row_dict.get(f"{h}_type_code", "s")
                variables[h] = val
                variables[f"{h}_type"] = TYPE_MAP.get(v_type, v_type)
                variables[f"{h}_type_code"] = v_type
                
            rendered_out_path = render_string_template(str(output_pattern), variables)
            rendered_out_path = clean_relative_path(rendered_out_path)
            ext = os.path.splitext(template_path)[1].lower()
            if not rendered_out_path.lower().endswith(ext):
                rendered_out_path += ext
            rendered_out_path = rendered_out_path.replace('/', os.sep).replace('\\', os.sep)
            if custom_out_dir:
                final_output_path = resolve_path(custom_out_dir, rendered_out_path)
            else:
                final_output_path = resolve_path(excel_dir, rendered_out_path)
            final_output_path = get_unique_path(final_output_path)
            
            print(f"  Рядок {r_num}:")
            if ext == '.docx':
                process_word(template_path, final_output_path, variables)
            elif ext == '.xlsx':
                process_excel(template_path, final_output_path, variables)
            else:
                print(f"  [Помилка] Непідтримуваний формат шаблону: {ext}")
