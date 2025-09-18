"""Quick environment verification for NVapps.
Run:  python verify_env.py
"""
from __future__ import annotations
import sys, importlib, platform

MODULES = ["openpyxl", "pandas", "matplotlib", "tkinter"]

def main() -> None:
    print("Interpreter:", sys.executable)
    print("Python version:", sys.version)
    print("Platform:", platform.platform())
    print("Checking modules:")
    for mod in MODULES:
        try:
            importlib.import_module(mod)
            print(f"  [OK] {mod}")
        except Exception as e:  # pragma: no cover
            print(f"  [FAIL] {mod}: {e}")
    print("Done.")

if __name__ == "__main__":
    main()
