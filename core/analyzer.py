import os
import re
import difflib
import openpyxl
from datetime import datetime
from openpyxl.styles import Alignment
from docx import Document

from core.text_processor import get_norm_key, tokenize, get_unique_path, render_string_template
from core.file_handlers.docx_handler import get_formatting_chunks, get_run_props, apply_run_props
from core.excel_styles import check_cell_style_diffs, serialize_color, serialize_fill, serialize_alignment, serialize_border

def find_all_vars_in_slot(strings, diff_map, v_idx, char_to_props=None):
    s0 = str(strings[0] or "")
    tok0 = tokenize(s0)
    
    tok_char_idx = []
    curr_idx = 0
    for tok in tok0:
        tok_char_idx.append(curr_idx)
        curr_idx += len(tok)
    tok_char_idx.append(curr_idx)

    is_var = [False] * len(tok0)
    for si in strings[1:]:
        if s0 == si: continue
        toki = tokenize(str(si or ""))
        sm = difflib.SequenceMatcher(None, tok0, toki)
        raw_ops = sm.get_opcodes()
        smooth_ops = []
        for i in range(len(raw_ops)):
            tag, i1, i2, j1, j2 = raw_ops[i]
            if tag == 'equal':
                text = "".join(toki[j1:j2])
                if 0 < i < len(raw_ops)-1 and len(text) < 3:
                    should_merge = True
                    if char_to_props is not None:
                        prev_op = raw_ops[i-1]
                        next_op = raw_ops[i+1]
                        start_char = tok_char_idx[prev_op[1]]
                        end_char = tok_char_idx[next_op[2]] if next_op[2] < len(tok_char_idx) else len(char_to_props)
                        props_slice = char_to_props[start_char:end_char]
                        if props_slice:
                            should_merge = (len(set(props_slice)) <= 1)
                    if should_merge:
                        tag = 'replace'
            smooth_ops.append((tag, i1, i2, j1, j2))
        for tag, i1, i2, j1, j2 in smooth_ops:
            if tag != 'equal':
                if tag == 'insert':
                    if i1 < len(tok0): is_var[i1] = True
                    elif i1 > 0: is_var[i1 - 1] = True
                for k in range(i1, i2): is_var[k] = True
    var_ranges = []
    i = 0
    while i < len(tok0):
        if is_var[i]:
            start = i
            i += 1
            if char_to_props is not None:
                start_char = tok_char_idx[start]
                while i < len(tok0) and is_var[i]:
                    curr_char = tok_char_idx[i]
                    next_char = tok_char_idx[i+1] if i+1 < len(tok_char_idx) else len(char_to_props)
                    slice_props = char_to_props[curr_char:next_char]
                    if slice_props and len(set([char_to_props[start_char]] + slice_props)) > 1:
                        break
                    i += 1
            else:
                while i < len(tok0) and is_var[i]: i += 1
            var_ranges.append((start, i))
        else: i += 1
    all_values = [{} for _ in range(len(strings))]
    template_parts = []
    last_idx = 0
    for idx, (v_start, v_end) in enumerate(var_ranges):
        template_parts.append("".join(tok0[last_idx:v_start]))
        vals = []
        for f_idx, si in enumerate(strings):
            if f_idx == 0: val = "".join(tok0[v_start:v_end])
            else:
                toki = tokenize(str(si or ""))
                sm = difflib.SequenceMatcher(None, tok0, toki)
                val_bits = []
                for tag, i1, i2, j1, j2 in sm.get_opcodes():
                    if tag == 'insert':
                        if v_start <= i1 <= v_end:
                            val_bits.append("".join(toki[j1:j2]))
                    else:
                        o_start = max(i1, v_start)
                        o_end = min(i2, v_end)
                        if o_start < o_end:
                            if tag == 'equal': val_bits.append("".join(tok0[o_start:o_end]))
                            else: val_bits.append("".join(toki[j1:j2]))
                val = "".join(val_bits)
            vals.append(val)
        template_parts.append(f"{{{{VAR_PLACEHOLDER_{idx}}}}}")
        for f_idx in range(len(strings)): all_values[f_idx][idx] = vals[f_idx]
        last_idx = v_end
    template_parts.append("".join(tok0[last_idx:]))
    return "".join(template_parts), var_ranges, all_values

def get_var_name(vals, diff_map, v_idx_list):
    norm_vals = tuple(get_norm_key(v) for v in vals)
    if norm_vals in diff_map: return diff_map[norm_vals]
    v_name = f"field_{v_idx_list[0]}"
    diff_map[norm_vals] = v_name
    v_idx_list[0] += 1
    return v_name

def process_paragraph_list(pars_matrix, diff_map, v_idx_list, all_data):
    num_files = len(pars_matrix)
    num_pars = min(len(p) for p in pars_matrix)
    for p_idx in range(num_pars):
        p_list = [pars_matrix[f_idx][p_idx] for f_idx in range(num_files)]
        
        chunk_lists = [get_formatting_chunks(p) for p in p_list]
        can_compare_chunks = all(len(cl) == len(chunk_lists[0]) for cl in chunk_lists) and chunk_lists[0]

        if can_compare_chunks:
            p0 = p_list[0]
            p0.clear()
            for c_idx in range(len(chunk_lists[0])):
                chunk_texts = [cl[c_idx][0] for cl in chunk_lists]
                chunk_props = chunk_lists[0][c_idx][1]

                if len(set(chunk_texts)) == 1:
                    run = p0.add_run(chunk_texts[0])
                    apply_run_props(run, chunk_props)
                    continue

                t_str, var_rs, vals_f = find_all_vars_in_slot(chunk_texts, diff_map, v_idx_list)
                if var_rs:
                    for idx in range(len(var_rs)):
                        v_name = get_var_name([v[idx] for v in vals_f], diff_map, v_idx_list)
                        t_str = t_str.replace(f"{{{{VAR_PLACEHOLDER_{idx}}}}}", f"{{{{{v_name}}}}}")
                        for f_idx in range(num_files):
                            all_data[f_idx][v_name] = vals_f[f_idx][idx]
                
                run = p0.add_run(t_str)
                apply_run_props(run, chunk_props)
            continue

        all_texts = [p.text for p in p_list]
        if len(set(all_texts)) == 1: continue

        p0 = p_list[0]
        char_to_props = []
        for r in p0.runs:
            props = get_run_props(r)
            for _ in range(len(r.text)):
                char_to_props.append(props)

        t_str, var_rs, vals_f = find_all_vars_in_slot(all_texts, diff_map, v_idx_list, char_to_props=char_to_props)
        if not var_rs: continue

        tok0 = tokenize(all_texts[0])
        tok_char_idx = []
        curr_idx = 0
        for tok in tok0:
            tok_char_idx.append(curr_idx)
            curr_idx += len(tok)
        tok_char_idx.append(curr_idx)

        new_char_props = []
        last_idx = 0

        for idx, (v_start, v_end) in enumerate(var_rs):
            v_name = get_var_name([v[idx] for v in vals_f], diff_map, v_idx_list)
            for f_idx in range(num_files):
                all_data[f_idx][v_name] = vals_f[f_idx][idx]

            start_char = tok_char_idx[last_idx]
            end_char = tok_char_idx[v_start]
            for i in range(start_char, end_char):
                if i < len(char_to_props):
                    new_char_props.append((all_texts[0][i], char_to_props[i]))

            placeholder = f"{{{{{v_name}}}}}"
            var_start_char = tok_char_idx[v_start]
            
            if var_start_char < len(char_to_props):
                p_props = char_to_props[var_start_char]
            else:
                p_props = char_to_props[-1] if char_to_props else None

            for ch in placeholder:
                new_char_props.append((ch, p_props))

            last_idx = v_end

        start_char = tok_char_idx[last_idx]
        for i in range(start_char, len(all_texts[0])):
            if i < len(char_to_props):
                new_char_props.append((all_texts[0][i], char_to_props[i]))

        new_fragments = []
        if new_char_props:
            cur_text = new_char_props[0][0]
            cur_props = new_char_props[0][1]
            for ch, props in new_char_props[1:]:
                if props == cur_props:
                    cur_text += ch
                else:
                    new_fragments.append((cur_text, cur_props))
                    cur_text = ch
                    cur_props = props
            new_fragments.append((cur_text, cur_props))

        p0.clear()
        for text, props in new_fragments:
            run = p0.add_run(text)
            apply_run_props(run, props)

def compare_word_group(files, template_out):
    docs = [Document(f) for f in files]
    diff_map, v_idx, all_data = {}, [1], [{} for _ in range(len(files))]
    process_paragraph_list([d.paragraphs for d in docs], diff_map, v_idx, all_data)
    processed_cells = set()
    for t_i in range(min(len(d.tables) for d in docs)):
        t_list = [d.tables[t_i] for d in docs]
        for r in range(min(len(t.rows) for t in t_list)):
            c_len = min(len(t.rows[r].cells) for t in t_list)
            for c in range(c_len):
                c_list = [t.rows[r].cells[c] for t in t_list]
                cell_key = tuple(cell._tc if cell is not None else None for cell in c_list)
                if any(x is None for x in cell_key) or cell_key in processed_cells:
                    continue
                processed_cells.add(cell_key)
                if len(set(cell.text for cell in c_list if cell is not None)) > 1:
                    process_paragraph_list([cell.paragraphs for cell in c_list if cell is not None], diff_map, v_idx, all_data)
    t_p = get_unique_path(template_out)
    docs[0].save(t_p)
    return t_p, all_data, v_idx[0]-1

def compare_excel_group(files, template_out):
    wbs_d = [openpyxl.load_workbook(f, data_only=True) for f in files]
    wb_t = openpyxl.load_workbook(files[0])
    for wb in wbs_d:
        for ws in wb.worksheets:
            normalize_sheet_numeric_cells(ws)
    for ws in wb_t.worksheets:
        normalize_sheet_numeric_cells(ws)
    diff_map, v_idx, all_data = {}, [1], [{} for _ in range(len(files))]
    num_sheets = min(len(wb.sheetnames) for wb in wbs_d)
    for i in range(num_sheets):
        titles = [wb.sheetnames[i] for wb in wbs_d]
        if len(set(titles)) > 1:
            t_str, var_rs, vals_f = find_all_vars_in_slot(titles, diff_map, v_idx)
            if var_rs:
                for idx in range(len(var_rs)):
                    v_name = get_var_name([v[idx] for v in vals_f], diff_map, v_idx)
                    t_str = t_str.replace(f"{{{{VAR_PLACEHOLDER_{idx}}}}}", f"{{{{{v_name}}}}}")
                    for f_idx in range(len(files)): all_data[f_idx][v_name] = vals_f[f_idx][idx]
                wb_t.worksheets[i].title = t_str
    for i in range(num_sheets):
        sheets_d = [wb.worksheets[i] for wb in wbs_d]
        st = wb_t.worksheets[i]
        max_r = max((s.max_row or 1) for s in sheets_d)
        max_c = max((s.max_column or 1) for s in sheets_d)
        for r in range(1, max_r + 1):
            for c in range(1, max_c + 1):
                cell_t = st.cell(row=r, column=c)
                if type(cell_t).__name__ == 'MergedCell': continue
                if cell_t.data_type == 'f' or (isinstance(cell_t.value, str) and str(cell_t.value).startswith('=')): continue
                vals = [s.cell(row=r, column=c).value for s in sheets_d]
                if len(set(vals)) > 1:
                    if not all(isinstance(v, str) for v in vals if v is not None):
                        v_name = get_var_name(vals, diff_map, v_idx)
                        cell_t.value = f"{{{{{v_name}}}}}"
                        for f_idx in range(len(files)): all_data[f_idx][v_name] = vals[f_idx]
                    else:
                        t_str, var_rs, vals_f = find_all_vars_in_slot([str(v or "") for v in vals], diff_map, v_idx)
                        if var_rs:
                            for idx in range(len(var_rs)):
                                v_name = get_var_name([v[idx] for v in vals_f], diff_map, v_idx)
                                t_str = t_str.replace(f"{{{{VAR_PLACEHOLDER_{idx}}}}}", f"{{{{{v_name}}}}}")
                                for f_idx in range(len(files)): all_data[f_idx][v_name] = vals_f[f_idx][idx]
                            cell_t.value = t_str
    t_p = get_unique_path(template_out)
    wb_t.save(t_p)
    return t_p, all_data, v_idx[0]-1

def generate_name_template(filenames, all_data):
    base_names = [os.path.splitext(os.path.basename(f))[0] for f in filenames]
    best_pattern = None
    max_placeholders = 0
    for i in range(len(base_names)):
        name = base_names[i]
        row_data = all_data[i]
        vars_sorted = sorted([(k, str(v).strip()) for k, v in row_data.items() if v and len(str(v).strip()) >= 2], key=lambda x: len(x[1]), reverse=True)
        pattern = name
        placeholders_count = 0
        for k, v in vars_sorted:
            if v in pattern:
                pattern = pattern.replace(v, f"{{{{{k}}}}}")
                placeholders_count += 1
        if placeholders_count > max_placeholders:
            max_placeholders = placeholders_count
            best_pattern = pattern
    if best_pattern: return best_pattern
    return f"{base_names[0]}_{{{{YYYY}}}}{{{{MM}}}}{{{{DD}}}}_{{{{hh}}}}{{{{mm}}}}"

def populate_config_sheet(ws, path, template_path, filenames, all_data, num_vars, relative_to_folder=None):
    try: rel_t = os.path.relpath(template_path, os.path.dirname(os.path.abspath(path)))
    except: rel_t = template_path
    
    name_pattern = generate_name_template(filenames, all_data)
    
    has_rel_dir = False
    has_file_name = False
    if relative_to_folder and filenames:
        has_rel_dir = True
        has_file_name = True
        folder_abs = os.path.abspath(relative_to_folder)
        for i, f in enumerate(filenames):
            try:
                f_abs = os.path.abspath(f)
                rel_file = os.path.relpath(f_abs, folder_abs)
                rel_dir = os.path.dirname(rel_file)
                rel_dir_clean = rel_dir.replace('\\', '/')
                if rel_dir_clean == "." or rel_dir_clean == "..":
                    rel_dir_clean = ""
                if i < len(all_data):
                    all_data[i]["rel_dir"] = rel_dir_clean
            except Exception:
                if i < len(all_data):
                    all_data[i]["rel_dir"] = ""
            
            try:
                base = os.path.splitext(os.path.basename(f))[0]
                if i < len(all_data):
                    all_data[i]["file_name"] = base
            except Exception:
                if i < len(all_data):
                    all_data[i]["file_name"] = ""
                    
        # Check if the generated pattern matches all filenames
        pattern_matches_all = True
        for i, f in enumerate(filenames):
            try:
                orig_name = os.path.splitext(os.path.basename(f))[0]
                rendered = render_string_template(name_pattern, all_data[i])
                if rendered != orig_name:
                    pattern_matches_all = False
                    break
            except Exception:
                pattern_matches_all = False
                break
                
        # If the pattern is a fallback pattern (contains YYYY) or doesn't match all original filenames, replace with file_name
        if "{{YYYY}}" in name_pattern or not pattern_matches_all:
            name_pattern = "{{file_name}}"
            
        name_pattern = f"{{{{rel_dir}}}}/{name_pattern}"
    else:
        if filenames:
            try:
                first_file_abs = os.path.abspath(filenames[0])
                config_dir = os.path.dirname(os.path.abspath(path))
                rel_file = os.path.relpath(first_file_abs, config_dir)
                rel_dir = os.path.dirname(rel_file)
                if rel_dir and rel_dir != "." and rel_dir != "..":
                    rel_dir_clean = rel_dir.replace('\\', '/')
                    if not rel_dir_clean.startswith('..'):
                        name_pattern = f"{rel_dir_clean}/{name_pattern}"
            except: pass

    ws['A1'], ws['A2'] = rel_t, name_pattern
    
    headers = []
    style_props = ['bold', 'italic', 'font_size', 'font_name', 'font_color', 'fill', 'alignment', 'border', 'number_format']
    for i in range(1, num_vars + 1):
        f_name = f"field_{i}"
        headers.append(f_name)
        for prop in style_props:
            prop_key = f"{f_name}_{prop}"
            if any(prop_key in row_dict for row_dict in all_data):
                headers.append(prop_key)
    if has_rel_dir:
        headers.append("rel_dir")
    if has_file_name:
        headers.append("file_name")
        
    for i, h in enumerate(headers): ws.cell(row=4, column=i+1).value = h
    
    unique_rows = []
    seen = set()
    for row_dict in all_data:
        row_tuple = tuple(get_norm_key(row_dict.get(h, "")) for h in headers)
        if row_tuple not in seen:
            unique_rows.append(row_dict)
            seen.add(row_tuple)
            
    for r_idx, row_dict in enumerate(unique_rows):
        for c_idx, h in enumerate(headers):
            val = row_dict.get(h, "")
            cell = ws.cell(row=5 + r_idx, column=c_idx + 1)
            if isinstance(val, str):
                val_strip = val.strip()
                if val_strip.isdigit():
                    cell.value = int(val_strip)
                elif re.match(r'^\d+\.\d+$', val_strip):
                    cell.value = float(val_strip)
                else:
                    cell.value = val
            else:
                cell.value = val
            if isinstance(cell.value, str) and '\n' in cell.value: cell.alignment = Alignment(wrapText=True)

def save_single_config(path, template_path, all_data, num_vars, filenames, relative_to_folder=None):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Settings"
    populate_config_sheet(ws, path, template_path, filenames, all_data, num_vars, relative_to_folder=relative_to_folder)
    f_p = get_unique_path(path)
    wb.save(f_p)
    return f_p

def save_master_config(path, results, relative_to_folder=None):
    wb = openpyxl.Workbook()
    if wb.active: wb.remove(wb.active)
    for idx, (filenames, template_path, all_data, num_vars) in enumerate(results):
        base_name = os.path.splitext(os.path.basename(filenames[0]))[0]
        sheet_title = re.sub(r'[\\/*?:\[\]]', "", base_name)[:30]
        if not sheet_title: sheet_title = f"Group_{idx+1}"
        ws = wb.create_sheet(title=sheet_title)
        populate_config_sheet(ws, path, template_path, filenames, all_data, num_vars, relative_to_folder=relative_to_folder)
    f_p = get_unique_path(path)
    wb.save(f_p)
    return f_p

def create_bat_file(bat_path, config_path):
    s_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
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

def is_file_a_template(path, ext):
    try:
        if ext == '.docx':
            doc = Document(path)
            for p in doc.paragraphs:
                if re.search(r'\{\{.*?\}\}', p.text): return True
            for t in doc.tables:
                for r in t.rows:
                    for c in r.cells:
                        if re.search(r'\{\{.*?\}\}', c.text): return True
        elif ext == '.xlsx':
            wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
            for sheet in wb.worksheets:
                for row in sheet.iter_rows(values_only=True):
                    for cell in row:
                        if isinstance(cell, str) and re.search(r'\{\{.*?\}\}', cell): return True
    except: pass
    return False

def normalize_sheet_numeric_cells(ws):
    for row in ws.iter_rows():
        for cell in row:
            if not hasattr(cell, 'value') or cell.value is None:
                continue
            if cell.data_type == 'f' or (isinstance(cell.value, str) and str(cell.value).startswith('=')):
                continue
            val = cell.value
            if isinstance(val, str):
                val_stripped = val.strip()
                if not val_stripped:
                    continue
                # Try integer
                if re.match(r'^-?\d+$', val_stripped):
                    try:
                        cell.value = int(val_stripped)
                        if cell.number_format in [None, 'General', '@']:
                            cell.number_format = '0'
                        continue
                    except:
                        pass
                # Try float with dot decimal separator
                if re.match(r'^-?\d+\.\d+$', val_stripped):
                    try:
                        cell.value = float(val_stripped)
                        if cell.number_format in [None, 'General', '@']:
                            cell.number_format = '0.00'
                        continue
                    except:
                        pass
                # Try float with comma decimal separator (common in Ukraine)
                if re.match(r'^-?\d+,\d+$', val_stripped):
                    try:
                        normalized_val = val_stripped.replace(',', '.')
                        cell.value = float(normalized_val)
                        if cell.number_format in [None, 'General', '@']:
                            cell.number_format = '0.00'
                        continue
                    except:
                        pass
            elif isinstance(val, bool):
                continue
            elif isinstance(val, (int, float)):
                if isinstance(val, int):
                    if cell.number_format in [None, 'General', '@']:
                        cell.number_format = '0'
                elif isinstance(val, float):
                    if cell.number_format in [None, 'General', '@']:
                        cell.number_format = '0.00'

def get_structure_key(path, ext):
    try:
        if ext == '.docx':
            doc = Document(path)
            if doc.tables:
                table_shapes = tuple((len(t.rows), len(t.columns)) for t in doc.tables)
                return ("word", "with_tables", len(doc.tables), table_shapes)
            else:
                return ("word", "no_tables", len(doc.paragraphs) // 5)
        elif ext == '.xlsx':
            wb = openpyxl.load_workbook(path, read_only=True)
            sheet_info = []
            
            def get_cell_style_tuple(cell):
                f = cell.font
                font_tuple = (
                    bool(f.bold) if f else False,
                    bool(f.italic) if f else False,
                    float(f.size) if (f and f.size) else 11.0,
                    str(f.name) if (f and f.name) else "Calibri",
                    serialize_color(f.color) if (f and f.color) else ""
                )
                fill_str = serialize_fill(cell.fill) if cell.fill else ""
                border_str = serialize_border(cell.border) if cell.border else ""
                
                # Normalize number format in-memory for the signature
                val = cell.value
                num_format = str(cell.number_format or "General")
                if isinstance(val, str):
                    val_stripped = val.strip()
                    if val_stripped:
                        if re.match(r'^-?\d+$', val_stripped):
                            if num_format in ['General', None, '@']:
                                num_format = '0'
                        elif re.match(r'^-?\d+\.\d+$', val_stripped) or re.match(r'^-?\d+,\d+$', val_stripped):
                            if num_format in ['General', None, '@']:
                                num_format = '0.00'
                elif isinstance(val, bool):
                    pass
                elif isinstance(val, (int, float)):
                    if isinstance(val, int):
                        if num_format in ['General', None, '@']:
                            num_format = '0'
                    elif isinstance(val, float):
                        if num_format in ['General', None, '@']:
                            num_format = '0.00'
                            
                return (font_tuple, fill_str, border_str, num_format)
                
            DEFAULT_STYLE = ((False, False, 11.0, 'Calibri', ''), '', '', 'General')
            
            def normalize_header_text(val):
                if val is None: return ""
                val_str = str(val).strip()
                if re.match(r'^\d+([.,]\d+)?$', val_str):
                    return ""
                return "".join(c for c in val_str.lower() if c.isalnum())
                
            for idx, sheetname in enumerate(wb.sheetnames):
                ws = wb[sheetname]
                formulas_header = []
                formula_cols_data = set()
                header_rows_cells = {}
                sheet_styles = []
                
                for row in ws.iter_rows(max_row=100, max_col=100):
                    for cell in row:
                        if not hasattr(cell, 'row'):
                            continue
                        val = cell.value
                        r, c = cell.row, cell.column
                        
                        # Capture cell style/formatting (ignoring default styles to save space)
                        style_tup = get_cell_style_tuple(cell)
                        if style_tup != DEFAULT_STYLE:
                            sheet_styles.append((r, c, style_tup))
                            
                        if val is not None:
                            is_formula = cell.data_type == 'f' or (isinstance(val, str) and val.startswith('='))
                            if is_formula:
                                if r <= 15:
                                    formulas_header.append((r, c))
                                else:
                                    formula_cols_data.add(c)
                            else:
                                if r <= 15:
                                    if r not in header_rows_cells:
                                        header_rows_cells[r] = []
                                    header_rows_cells[r].append((c, val))
                                    
                header_row_idx = 1
                max_pop = 0
                for r in range(1, 16):
                    cells = header_rows_cells.get(r, [])
                    pop_count = sum(1 for col, v in cells if normalize_header_text(v))
                    if pop_count > max_pop:
                        max_pop = pop_count
                        header_row_idx = r
                        
                header_cells = header_rows_cells.get(header_row_idx, [])
                normalized_headers = tuple(col for col, v in header_cells if normalize_header_text(v))
                
                sheet_info.append((
                    idx,
                    tuple(formulas_header),
                    tuple(sorted(formula_cols_data)),
                    normalized_headers,
                    tuple(sheet_styles)
                ))
            return ("excel", len(wb.sheetnames), tuple(sheet_info))
    except: return None

def sort_files_by_complexity(files, ext):
    def get_complexity(f):
        try:
            if ext == '.docx':
                doc = Document(f)
                return sum(len(p.text) for p in doc.paragraphs) + sum(len(c.text) for t in doc.tables for r in t.rows for c in r.cells)
            elif ext == '.xlsx':
                wb = openpyxl.load_workbook(f, data_only=True)
                return sum(len(str(cell.value or "")) for sheet in wb.worksheets for row in sheet.iter_rows() for cell in row)
        except:
            return 0
    return sorted(files, key=get_complexity, reverse=True)

def run_compare_two(f1, f2):
    ext = os.path.splitext(f1)[1].lower()
    if ext not in ['.docx', '.xlsx'] or os.path.splitext(f2)[1].lower() != ext:
        print("Помилка: Файли повинні мати однакове розширення (.docx або .xlsx).")
        return
    files = sort_files_by_complexity([f1, f2], ext)
    f1_a = os.path.abspath(files[0])
    f_dir, f_name = os.path.dirname(f1_a), os.path.splitext(os.path.basename(f1_a))[0]
    t_o = os.path.join(f_dir, f"template_{f_name}{ext}")
    m_x = os.path.join(f_dir, f"{f_name}_config.xlsx")
    b_o = os.path.join(f_dir, f"{f_name}_run_all.bat")
    print(f"Аналіз 2 файлів (Режим: Порівняння)...")
    try:
        if ext == '.docx': t_p, data, n_v = compare_word_group(files, t_o)
        else: t_p, data, n_v = compare_excel_group(files, t_o)
        if n_v == 0:
            print(f"\nКонфіг та BAT-файл не створено, оскільки змінних не знайдено.")
            if os.path.exists(t_p):
                try: os.remove(t_p)
                except: pass
        else:
            m_p = save_single_config(m_x, t_p, data, n_v, files)
            create_bat_file(b_o, m_p)
            print(f"\nУспішно!\nШаблон: {t_p}\nКонфіг: {m_p}")
    except Exception as e: print(f"Помилка: {e}")

def run_package(sample, folder):
    ext = os.path.splitext(sample)[1].lower()
    sample_key = get_structure_key(sample, ext)
    if not sample_key:
        print("Помилка: Не вдалося визначити структуру файлу-зразка.")
        return
    files = [sample]
    for root, _, fs in os.walk(folder):
        for f in fs:
            if f.lower().endswith(ext):
                if "template" in f.lower(): continue
                fp = os.path.join(root, f)
                if os.path.abspath(fp) == os.path.abspath(sample): continue
                if is_file_a_template(fp, ext): continue
                if get_structure_key(fp, ext) == sample_key:
                    files.append(fp)
    if len(files) < 2:
        print("Не знайдено подібних файлів для порівняння.")
        return
    files = sort_files_by_complexity(files, ext)
    sample_a = os.path.abspath(files[0])
    f_dir, f_name = os.path.dirname(sample_a), os.path.splitext(os.path.basename(sample_a))[0]
    t_o = os.path.join(f_dir, f"template_{f_name}{ext}")
    m_x = os.path.join(f_dir, f"{f_name}_config.xlsx")
    b_o = os.path.join(f_dir, f"{f_name}_run_all.bat")
    print(f"Аналіз {len(files)} файлів (Режим: Пакетний)...")
    try:
        if ext == '.docx': t_p, data, n_v = compare_word_group(files, t_o)
        else: t_p, data, n_v = compare_excel_group(files, t_o)
        if n_v == 0:
            print(f"\nКонфіг та BAT-файл не створено, оскільки змінних не знайдено.")
            if os.path.exists(t_p):
                try: os.remove(t_p)
                except: pass
        else:
            m_p = save_single_config(m_x, t_p, data, n_v, files, relative_to_folder=folder)
            create_bat_file(b_o, m_p)
            print(f"\nУспішно!\nШаблон: {t_p}\nКонфіг: {m_p}")
    except Exception as e: print(f"Помилка: {e}")

def get_name_prefix(filename):
    name = os.path.splitext(filename)[0]
    name = re.sub(r'\d+$', '', name)
    parts = name.split('_')
    parts = [p for p in parts if p]
    if len(parts) > 1:
        last = parts[-1].lower()
        if last.isdigit() or len(last) == 1 or last in ['a', 'b', 'c', 'd', 'x', 'y', 'z']:
            parts = parts[:-1]
        if parts:
            last = parts[-1].lower()
            if last in ['pair', 'triplet', 'v2', 'v1', 'v3', 'copy']:
                parts = parts[:-1]
    res = "_".join(parts)
    res = re.sub(r'[\-_]+$', '', res)
    return res

def wb_sheetname_by_idx(filepath, idx):
    try:
        wb = openpyxl.load_workbook(filepath, read_only=True)
        if idx < len(wb.sheetnames):
            return wb.sheetnames[idx]
    except:
        pass
    return f"Sheet {idx+1}"

def run_full_auto(folder):
    print(f"Сканування папки: {folder} (Режим: Повний автопілот)...")
    groups = {}
    for root, _, fs in os.walk(folder):
        for f in fs:
            ext = os.path.splitext(f)[1].lower()
            if ext not in ['.docx', '.xlsx']: continue
            if "template" in f.lower(): continue
            fp = os.path.join(root, f)
            if is_file_a_template(fp, ext): continue
            key = get_structure_key(fp, ext)
            if key:
                if key not in groups: groups[key] = []
                groups[key].append(fp)
                
    if not groups:
        print("Не знайдено підходящих документів.")
        return
        
    print(f"\n--- Результати групування (Знайдено унікальних сигнатур: {len(groups)}) ---")
    for idx, (key, files) in enumerate(groups.items()):
        ftype = key[0]
        criteria_parts = []
        if ftype == "excel":
            num_sheets = key[1]
            criteria_parts.append(f"Тип: Excel, кількість аркушів: {num_sheets}")
            for sh_idx, (idx_sh, form_h, form_d, headers, styles) in enumerate(key[2]):
                sh_details = []
                if len(headers) > 0: sh_details.append(f"стовпців: {len(headers)}")
                if len(form_h) > 0: sh_details.append(f"формул у заголовку: {len(form_h)}")
                if len(form_d) > 0: sh_details.append(f"колонки з формулами: {list(form_d)}")
                if len(styles) > 0: sh_details.append(f"стилізованих комірок: {len(styles)}")
                if sh_details:
                    criteria_parts.append(f"  Аркуш {idx_sh+1} ({wb_sheetname_by_idx(files[0], idx_sh)}): {', '.join(sh_details)}")
        elif ftype == "word":
            if key[1] == "with_tables":
                criteria_parts.append(f"Тип: Word, з таблицями ({key[2]} шт.), розміри таблиць: {key[3]}")
            else:
                criteria_parts.append(f"Тип: Word, без таблиць (орієнтовно {key[2]*5} абзаців)")
                
        criteria_str = "\n".join(criteria_parts)
        print(f"\n[Сигнатура {idx+1}] ({len(files)} файлів):\n{criteria_str}")
        print("Склад групи:")
        for f in sorted(files):
            print(f"  - {os.path.basename(f)} (відносно: {os.path.relpath(f, folder)})")
    print("\n" + "="*60 + "\n")
    
    results = []
    out_dir = os.getcwd()
    for key, files in groups.items():
        if len(files) < 2:
            print(f"Група з файлом {os.path.basename(files[0])} має лише один файл, пропускаємо.")
            continue
        ext = os.path.splitext(files[0])[1].lower()
        files = sort_files_by_complexity(files, ext)
        base_name = os.path.splitext(os.path.basename(files[0]))[0]
        t_o = os.path.join(out_dir, f"template_{base_name}{ext}")
        print(f"Обробка групи: {base_name} ({len(files)} файлів)...")
        try:
            if ext == '.docx': t_p, data, n_v = compare_word_group(files, t_o)
            else: t_p, data, n_v = compare_excel_group(files, t_o)
            print(f"  - Знайдено змінних: {n_v}")
            if n_v > 0:
                results.append((files, t_p, data, n_v))
            else:
                print(f"  - Аркуш пропущено: змінних не знайдено.")
                if os.path.exists(t_p):
                    try: os.remove(t_p)
                    except: pass
        except Exception as e:
            import traceback
            print(f"  - Помилка: {e}")
            traceback.print_exc()
    if not results:
        print("Не вдалося створити жодного шаблону (або змінні відсутні).")
        return
    m_p = save_master_config(os.path.join(out_dir, "Auto_Config.xlsx"), results, relative_to_folder=folder)
    create_bat_file(os.path.join(out_dir, "Auto_Run_All.bat"), m_p)
    print(f"\nГотово!\nСтворено конфіг: {m_p}\nОброблено груп: {len(results)}")
