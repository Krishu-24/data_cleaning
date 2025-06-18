import pandas as pd

# ---------- 1. Clean Numbers ----------
def clean_numbers(df):
    flagged = []

    for index, row in df.iterrows():
        tel = str(row['Telephone']) if not pd.isna(row['Telephone']) else ''
        contact = str(row['Contact person -Telephone']) if not pd.isna(row['Contact person -Telephone']) else ''

        for target in ['+91', '-', ' ']:
            tel = tel.replace(target, '')
            contact = contact.replace(target, '')

        # Remove bracket content
        if '(' in tel and ')' in tel:
            tel = tel.split('(')[0]
        if '(' in contact and ')' in contact:
            contact = contact.split('(')[0]

        tel = tel.strip()
        contact = contact.strip()

        # Check validity
        tel_invalid = not tel.isdigit() or len(tel) != 10 or any(c in tel for c in ['/', '(', ')'])
        contact_invalid = not contact.isdigit() or len(contact) != 10 or any(c in contact for c in ['/', '(', ')'])
        contact_exists = contact != ''

        # Clear contact if same and not flagged
        if tel == contact and not contact_invalid:
            contact = ''
            contact_exists = False

        # Apply cleaned
        df.at[index, 'Telephone'] = tel
        df.at[index, 'Contact person -Telephone'] = contact

        is_flagged = tel_invalid or (contact_exists and contact_invalid)
        flagged.append('flagged' if is_flagged else '')

    df['Flagged'] = flagged
    return df

# ---------- 2. Clean Names ----------
def clean_names(df):
    for index, row in df.iterrows():
        first = str(row['Contact person -First name']) if not pd.isna(row['Contact person -First name']) else ''
        last = str(row['Contact person -Last name']) if not pd.isna(row['Contact person -Last name']) else ''

        first = first.strip()
        last = last.strip()

        lowered = first.lower().replace(' ', '')  # normalize for checking

        if lowered in ['mr', 'mr.']:
            first = last
            last = ''
        else:
            prefixes = ['mr.', 'mr', 'mrs.', 'mrs']
            for prefix in prefixes:
                if lowered.startswith(prefix):
                    cleaned = first.lower().replace(prefix, '', 1).lstrip()
                    first = cleaned.capitalize()
                    break

        df.at[index, 'Contact person -First name'] = first
        df.at[index, 'Contact person -Last name'] = last

    return df

# ---------- 3. Clean City ----------
def clean_city_names(df):
    for index, row in df.iterrows():
        city = str(row['City']) if not pd.isna(row['City']) else ''
        clean_city = ''

        for char in city:
            if char.isalpha():
                clean_city += char
            else:
                break

        if clean_city:
            clean_city = clean_city[0].upper() + clean_city[1:].lower()

        df.at[index, 'City'] = clean_city

    return df

# ---------- 4. Main ----------
def main():
    file_path = "C:/Users/krish/Downloads/Customer_list_06102025063059.xlsx"
    df = pd.read_excel(file_path)

    df = clean_numbers(df)
    df = clean_names(df)
    df = clean_city_names(df)

    df.to_excel("C:/Users/krish/Downloads/cleaned.xlsx", index=False)
    print("✅ All cleaning complete. File saved as 'cleaned.xlsx'.")

if __name__ == "__main__":
    main()
