import sys
import os

from core.analyzer import run_full_auto, run_package, run_compare_two
from core.generator import run_generation

def heal_cmd_string(s):
    if not isinstance(s, str): return s
    for enc in ['cp1251', 'cp866', 'cp1252', 'utf-8']:
        try:
            candidate = s.encode(enc).decode('utf-8')
            if candidate and any(ord(c) > 127 for c in candidate):
                return candidate
        except: pass
    return s

def main():
    args = [heal_cmd_string(arg) for arg in sys.argv[1:]]
    
    ignore_single = False
    if "--ignore-single" in args:
        ignore_single = True
        args.remove("--ignore-single")
        
    if len(args) == 0:
        print("Використання СИСТЕМИ _templates_machine_:")
        print("  [СТВОРЕННЯ ШАБЛОНІВ ТА КОНФІГІВ]")
        print("  1. Повний автопілот: python _templates_machine_.py <папка> [--ignore-single]")
        print("  2. Пакетний режим:   python _templates_machine_.py <файл_зразок> <папка>")
        print("  3. Порівняння двох:  python _templates_machine_.py <файл_1> <файл_2>\n")
        print("  [ГЕНЕРАЦІЯ ГОТОВИХ ДОКУМЕНТІВ]")
        print("  4. Запуск генерації: python _templates_machine_.py <файл_конфігу.xlsx> [аркуш] [рядок] [папка_вихідних_файлів]")
        return
        
    arg1 = args[0]
    
    if len(args) == 1 and os.path.isdir(arg1):
        run_full_auto(arg1, ignore_single=ignore_single)
        return
        
    if len(args) == 2 and os.path.isfile(arg1) and os.path.isdir(args[1]):
        run_package(arg1, args[1])
        return
        
    if len(args) == 2 and os.path.isfile(arg1) and os.path.isfile(args[1]):
        run_compare_two(arg1, args[1])
        return
        
    if arg1.lower().endswith('.xlsx') and os.path.isfile(arg1):
        sheet = args[1] if len(args) > 1 else "all"
        row = args[2] if len(args) > 2 else "all"
        custom_out_dir = args[3] if len(args) > 3 else None
        run_generation(arg1, sheet, row, custom_out_dir)
        return
        
    print("Помилка: Невідома комбінація параметрів.")
    print("Запустіть скрипт без параметрів для виклику довідки.")

if __name__ == "__main__":
    main()
