
import sys
import pandas as pd
import os

print("Testing openpyxl installation...")
try:
    import openpyxl
    print(f"openpyxl version: {openpyxl.__version__}")
except ImportError:
    print("FAILED: openpyxl is not installed.")
    sys.exit(1)

print("Testing Excel export via pandas...")
try:
    data = [{"Name": "Test Company", "BRN": "12345", "Status": "Active"}]
    df = pd.DataFrame(data)
    test_file = "test_export.xlsx"
    df.to_excel(test_file, index=False)
    
    if os.path.exists(test_file):
        print(f"SUCCESS: Created {test_file}")
        os.remove(test_file)
    else:
        print("FAILED: File was not created.")
        sys.exit(1)

except Exception as e:
    print(f"FAILED: Error during export: {e}")
    sys.exit(1)

print("Verification complete.")
