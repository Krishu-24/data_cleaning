import pandas as pd
import re

# Step 1: Load Excel
file_path = 'your_excel_file.xlsx'  # Replace with actual file path
df = pd.read_excel(file_path)

# Step 2: Clean and flag function
def clean_number_full(num):
    if pd.isna(num):
        return num, '', False  # original, cleaned, flag

    num_str = str(num).strip()

    # Detect if alt numbers or junk exist
    is_flagged = False
    if '(' in num_str or ')' in num_str or '/' in num_str:
        is_flagged = True

    # Remove content inside brackets (and the brackets themselves)
    num_str_no_brackets = re.sub(r'\(.*?\)', '', num_str)

    # Remove +91, -, and spaces
    num_str_cleaned = num_str_no_brackets.replace('+91', '').replace('-', '').replace(' ', '').strip()

    return num, num_str_cleaned, is_flagged

# Step 3: Apply to both columns
tel_processed = df['Telephone'].apply(clean_number_full)
contact_processed = df['Contact person - Telephone'].apply(clean_number_full)

# Step 4: Assign split values into new columns
df['Telephone'] = tel_processed.apply(lambda x: x[0])
df['Cleaned Telephone'] = tel_processed.apply(lambda x: x[1])
df['Contact person - Telephone'] = contact_processed.apply(lambda x: x[0])
df['Cleaned Contact person - Telephone'] = contact_processed.apply(lambda x: x[1])
df['Flagged Reason'] = ['flagged' if t[2] or c[2] else '' for t, c in zip(tel_processed, contact_processed)]

# Step 5: If cleaned values are equal and not flagged, remove contact person number
mask = (df['Cleaned Telephone'] == df['Cleaned Contact person - Telephone']) & (df['Flagged Reason'] == '')
df.loc[mask, 'Contact person - Telephone'] = ''

# Step 6: Save result
df.to_excel('cleaned_excel_file.xlsx', index=False)

print("✨ File cleaned and saved with flagged rows and cleaned columns.")


[Running] python -u "c:\Users\krish\Desktop\cleaner.py"
Traceback (most recent call last):
  File "C:\Users\krish\AppData\Local\Programs\Python\Python39\lib\site-packages\pandas\compat\_optional.py", line 135, in import_optional_dependency
    module = importlib.import_module(name)
  File "C:\Users\krish\AppData\Local\Programs\Python\Python39\lib\importlib\__init__.py", line 127, in import_module
    return _bootstrap._gcd_import(name[level:], package, level)
  File "<frozen importlib._bootstrap>", line 1030, in _gcd_import
  File "<frozen importlib._bootstrap>", line 1007, in _find_and_load
  File "<frozen importlib._bootstrap>", line 984, in _find_and_load_unlocked
ModuleNotFoundError: No module named 'openpyxl'

During handling of the above exception, another exception occurred:

Traceback (most recent call last):
  File "c:\Users\krish\Desktop\cleaner.py", line 6, in <module>
    df = pd.read_excel(file_path)
  File "C:\Users\krish\AppData\Local\Programs\Python\Python39\lib\site-packages\pandas\io\excel\_base.py", line 495, in read_excel
    io = ExcelFile(
  File "C:\Users\krish\AppData\Local\Programs\Python\Python39\lib\site-packages\pandas\io\excel\_base.py", line 1567, in __init__
    self._reader = self._engines[engine](
  File "C:\Users\krish\AppData\Local\Programs\Python\Python39\lib\site-packages\pandas\io\excel\_openpyxl.py", line 552, in __init__
    import_optional_dependency("openpyxl")
  File "C:\Users\krish\AppData\Local\Programs\Python\Python39\lib\site-packages\pandas\compat\_optional.py", line 138, in import_optional_dependency
    raise ImportError(msg)
ImportError: Missing optional dependency 'openpyxl'.  Use pip or conda to install openpyxl.

[Done] exited with code=1 in 1.257 seconds




