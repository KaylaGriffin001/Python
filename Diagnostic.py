import pandas as pd

# Load the Excel file
file_path = 'C:/Users/kayla/Desktop/nonprofits_info.xlsx'
sheet_name = 'FL'

# Read the Excel sheet
df = pd.read_excel(file_path, sheet_name=sheet_name)

print("="*60)
print("EXCEL FILE DIAGNOSTICS")
print("="*60)
print()

print("COLUMN NAMES:")
for i, col in enumerate(df.columns):
    print(f"  Column {i} (Excel Column {chr(65+i)}): '{col}'")
print()

print("TOTAL ROWS:", len(df))
print()

print("FIRST 5 ROWS:")
print(df.head())
print()

print("="*60)
print("CHECKING COLUMNS E AND F (indices 4 and 5):")
print("="*60)

if len(df.columns) > 4:
    col_e_name = df.columns[4]
    print(f"\nColumn E name: '{col_e_name}'")
    print(f"First 3 values in Column E:")
    for i in range(min(3, len(df))):
        val = df.iloc[i, 4]
        print(f"  Row {i+1}: '{val}' (type: {type(val).__name__}, is null: {pd.isna(val)})")

if len(df.columns) > 5:
    col_f_name = df.columns[5]
    print(f"\nColumn F name: '{col_f_name}'")
    print(f"First 3 values in Column F:")
    for i in range(min(3, len(df))):
        val = df.iloc[i, 5]
        print(f"  Row {i+1}: '{val}' (type: {type(val).__name__}, is null: {pd.isna(val)})")
else:
    print("\nColumn F doesn't exist yet!")