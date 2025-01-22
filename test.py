import pandas as pd

file_path = r"C:\Users\jscheftic\OneDrive - spsi.com\Desktop\SPSI Code\SPSI Signatures\employees_spsi.xlsx"
try:
    df = pd.read_excel(file_path, engine='openpyxl')  # Explicitly use 'openpyxl' for .xlsx files
    print("File loaded successfully!")
except Exception as e:
    print(f"Error loading file: {e}")