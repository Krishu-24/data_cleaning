import pandas as pd

# ---------- 1. Clean Numbers ----------
def clean_numbers(df):
    flagged = []

    for index, row in df.iterrows():
        tel = str(row['Telephone']) if not pd.isna(row['Telephone']) else ''
        contact = str(row['Contact person - Telephone']) if not pd.isna(row['Contact person - Telephone']) else ''

        # Remove unwanted characters
        for target in ['+91', '-', ' ']:
            tel = tel.replace(target, '')
            contact = contact.replace(target, '')

        # Check for flags
        is_flagged = False
        if '(' in tel or ')' in tel or '/' in tel or not tel.isdigit() or len(tel) != 10:
            is_flagged = True
        if '(' in contact or ')' in contact or '/' in contact or not contact.isdigit() or len(contact) != 10:
            is_flagged = True

        # Save cleaned values back
        df.at[index, 'Telephone'] = tel
        df.at[index, 'Contact person - Telephone'] = contact

        # Clear contact if same as main number and not flagged
        if tel == contact and not is_flagged:
            df.at[index, 'Contact person - Telephone'] = ''

        # Set flag
        flagged.append('flagged' if is_flagged else '')

    df['Flagged'] = flagged
    return df

# ---------- 2. Clean Names ----------
def clean_names(df):
    for index, row in df.iterrows():
        first = str(row['Contact person - First name']) if not pd.isna(row['Contact person - First name']) else ''
        last = str(row['Contact person - Last name']) if not pd.isna(row['Contact person - Last name']) else ''

        # Check for exact "Mr" or "Mr." only
        if first.strip().lower() in ['mr', 'mr.']:
            df.at[index, 'Contact person - First name'] = last
            df.at[index, 'Contact person - Last name'] = ''
        else:
            # Clean NaNs anyway
            df.at[index, 'Contact person - First name'] = first
            df.at[index, 'Contact person - Last name'] = last

    return df

# ---------- 3. Clean City ----------
def clean_city_names(df):
    for index, row in df.iterrows():
        city = str(row['City']) if not pd.isna(row['City']) else ''

        # Remove everything after the first non-letter character
        clean_city = ''
        for char in city:
            if char.isalpha():
                clean_city += char
            else:
                break

        # Capitalize first letter
        if clean_city:
            clean_city = clean_city[0].upper() + clean_city[1:].lower()

        df.at[index, 'City'] = clean_city

    return df

# ---------- 4. Main ----------
def main():
    file_path = 'your_excel_file.xlsx'  # <- Change this to your actual Excel file
    df = pd.read_excel(file_path)

    df = clean_numbers(df)
    df = clean_names(df)
    df = clean_city_names(df)

    df.to_excel('cleaned_excel_file.xlsx', index=False)
    print("✅ File cleaned and saved as 'cleaned_excel_file.xlsx'")

if __name__ == "__main__":
    main()
