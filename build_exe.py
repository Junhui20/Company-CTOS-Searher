"""
Build script to create CTOS Scraper executable.
Usage: python build_exe.py

Supports Windows, macOS, and Linux.
"""

import subprocess
import sys
import os
import platform


def get_package_path(package_name: str) -> str:
    """Get the installation path of a Python package."""
    import importlib
    mod = importlib.import_module(package_name)
    return mod.__path__[0]


def add_data_arg(src: str, dest: str) -> str:
    """Return --add-data value with the correct OS separator."""
    sep = ";" if platform.system() == "Windows" else ":"
    return f"{src}{sep}{dest}"


def build() -> None:
    ctk_path = get_package_path("customtkinter")

    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--noconfirm",
        "--onedir",
        "--windowed",
        "--name", "CTOS Scraper",
        "--icon", "NONE",
        # Bundle customtkinter assets (themes, JSON files)
        "--add-data", add_data_arg(ctk_path, "customtkinter"),
        # Include app package
        "--add-data", add_data_arg("app", "app"),
        # Hidden imports that PyInstaller might miss
        "--hidden-import", "customtkinter",
        "--hidden-import", "pandas",
        "--hidden-import", "openpyxl",
        "--hidden-import", "selenium",
        "--hidden-import", "webdriver_manager",
        "--hidden-import", "sqlite3",
        # Collect all submodules
        "--collect-all", "customtkinter",
        "--collect-all", "webdriver_manager",
        # Exclude unneeded heavy packages
        "--exclude-module", "torch",
        "--exclude-module", "torchvision",
        "--exclude-module", "scipy",
        "--exclude-module", "matplotlib",
        "--exclude-module", "numpy.f2py",
        "--exclude-module", "pytest",
        "--exclude-module", "sympy",
        "--exclude-module", "pyarrow",
        "--exclude-module", "tensorboard",
        # Entry point
        "main.py",
    ]

    print("Building CTOS Scraper executable...")
    print(f"Platform: {platform.system()} ({platform.machine()})")
    print(f"Command: {' '.join(cmd)}")
    print()

    result = subprocess.run(cmd, cwd=os.path.dirname(os.path.abspath(__file__)))

    if result.returncode == 0:
        print("\n" + "=" * 50)
        print("Build successful!")
        print(f"Output: dist/CTOS Scraper/")
        print("=" * 50)
    else:
        print("\nBuild failed!")
        sys.exit(1)


if __name__ == "__main__":
    build()
