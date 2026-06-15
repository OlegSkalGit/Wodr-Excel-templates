import os
import re
from datetime import datetime

def get_unique_path(full_path):
    directory = os.path.dirname(full_path)
    if directory and not os.path.exists(directory): 
        os.makedirs(directory, exist_ok=True)
    if not os.path.exists(full_path): return full_path
    filename = os.path.basename(full_path)
    name, ext = os.path.splitext(filename)
    counter = 1
    unique_path = full_path
    while os.path.exists(unique_path):
        unique_path = os.path.join(directory, f"{name}_{counter}{ext}")
        counter += 1
    return unique_path

def get_norm_key(v):
    if v is None: return ""
    if isinstance(v, (int, float)):
        try:
            if float(v).is_integer(): return str(int(v))
        except: pass
        return str(v)
    return str(v)

def tokenize(s):
    if not isinstance(s, str): return [str(s)]
    return re.findall(r'\w+|[^\w\s]+|\s+', s)

def resolve_path(base_dir, path):
    if os.path.isabs(path): 
        return os.path.abspath(path)
    return os.path.abspath(os.path.join(base_dir, path))

def render_string_template(template_str, variables):
    result = str(template_str)
    changed = True
    iters = 0
    max_iters = 20
    while changed and iters < max_iters:
        changed = False
        iters += 1
        for key, val in variables.items():
            pattern = r"\{\{\s*" + re.escape(key) + r"\s*\}\}"
            if re.search(pattern, result):
                result = re.sub(pattern, lambda _: str(val), result)
                changed = True
    if iters >= max_iters:
        print(f"Попередження: виявлено можливу циклічну рекурсію в шаблоні: {template_str}")
    return result

def get_now_vars():
    now = datetime.now()
    return {
        "YYYY": now.strftime("%Y"), "MM": now.strftime("%m"), "DD": now.strftime("%d"),
        "hh": now.strftime("%H"), "mm": now.strftime("%M"), "ss": now.strftime("%S")
    }

def clean_relative_path(path):
    if not isinstance(path, str):
        path = str(path or "")
    import re
    if re.match(r'^[a-zA-Z]:', path) or path.startswith('\\\\') or path.startswith('//'):
        return path
    p = path.replace('\\', '/')
    while p.startswith('/'):
        p = p[1:]
    while '//' in p:
        p = p.replace('//', '/')
    return p

def resolve_virtual_doc_name(pattern, row_data, template_path):
    """Resolves output document name using date/time variables and row data."""
    now_vars = get_now_vars()
    variables = {**now_vars, **row_data}
    
    result = str(pattern)
    for key, val in variables.items():
        pattern_re = r"\{\{\s*" + re.escape(key) + r"\s*\}\}"
        result = re.sub(pattern_re, lambda _, v=val: str(v), result)
        
    result = clean_relative_path(result)
    
    if template_path:
        ext = os.path.splitext(template_path)[1].lower()
        if not result.lower().endswith(ext):
            result += ext
    return result

