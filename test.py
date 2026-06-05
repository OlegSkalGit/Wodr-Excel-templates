import sys

print("--- Running templates_machine core imports test ---")

try:
    import core.cmd_helpers
    import core.analyzer
    import core.generator
    import core.io_utils
    import core.text_processor
    import core.excel_styles
    print("[OK] All core modules imported successfully!")
except ImportError as e:
    print(f"[FAIL] Import failed: {e}")
    sys.exit(1)

print("[OK] Test execution finished successfully.")
