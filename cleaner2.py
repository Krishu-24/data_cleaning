import pandas as pd
import re

# ---------- 1. Clean Numbers Function ----------
def clean_numbers(df):
    def clean_and_flag(num):
        if pd.isna(num):
            return '', False

        num_str = str(num).strip()

        # Flag if contains brackets or slash
        is_flagged = any(c in num_str for c in ['(', ')', '/'])

        # Remove anything in brackets
        num_str = re.sub(r'\(.*?\)', '', num_str)

        # Remove +91, -, and spaces
        cleaned = num_str.replace('+91', '').replace('-', '').replace(' ', '').strip()

        # Additional flag: if not exactly 10 digits or not all digits
        if not cleaned.isdigit() or len(cleaned) != 10:
            is_flagged = True

        return cleaned, is_flagged

    # Apply to Telephone & Contact columns
    tel_result = df['Telephone'].apply(clean_and_flag)
    contact_result = df['Contact person - Telephone'].apply(clean_and_flag)

    # Assign cleaned values
    df['Telephone'] = tel_result.apply(lambda x: x[0])
    df['Contact person - Telephone'] = contact_result.apply(lambda x: x[0])

    # Combine flags
    df['Flagged'] = ['flagged' if t[1] or c[1] else '' for t, c in zip(tel_result, contact_result)]

    # Clear contact number if it matches Telephone and is not flagged
    mask = (df['Telephone'] == df['Contact person - Telephone']) & (df['Flagged'] == '')
    df.loc[mask, 'Contact person - Telephone'] = ''

    return df

# ---------- 2. Clean Names Function ----------
def clean_names(df):
    def replace_mr_only(row):
        first = str(row['Contact person - First name']).strip()
        last = str(row['Contact person - Last name']).strip()

        if first.lower() in ['mr', 'mr.']:
            return pd.Series([last, ''])  # Move last to first, clear last
        return pd.Series([first, last])

    df[['Contact person - First name', 'Contact person - Last name']] = df.apply(replace_mr_only, axis=1)
    return df

# ---------- 3. Clean City Function ----------
def clean_city_names(df):
    def clean_city(city):
        if pd.isna(city):
            return city
        city = str(city).strip()
        match = re.match(r'^[A-Za-z]+', city)
        return match.group(0).capitalize() if match else ''
    
    df['City'] = df['City'].apply(clean_city)
    return df

# ---------- 4. Main Execution ----------
def main():
    file_path = 'your_excel_file.xlsx'  # <-- Update this to your actual file
    df = pd.read_excel(file_path)

    df = clean_numbers(df)
    df = clean_names(df)
    df = clean_city_names(df)

    df.to_excel('cleaned_excel_file.xlsx', index=False)
    print("✅ All cleaning complete. File saved as 'cleaned_excel_file.xlsx'.")

if __name__ == "__main__":
    main()
