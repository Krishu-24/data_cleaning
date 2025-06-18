import pandas as pd

# ---------- 1. Clean Numbers Function ----------
def clean_numbers(df):
    flagged = []

    for index, row in df.iterrows():
        tel = str(row['Telephone']) if not pd.isna(row['Telephone']) else ''
        contact = str(row['Contact person -Telephone']) if not pd.isna(row['Contact person -Telephone']) else ''

        # Remove unwanted characters
        for target in ['+91', '-', ' ']:
            tel = tel.replace(target, '')
            contact = contact.replace(target, '')

        # Remove bracket content if any
        if '(' in tel and ')' in tel:
            tel = tel.split('(')[0]
        if '(' in contact and ')' in contact:
            contact = contact.split('(')[0]

        # Check for flags
        is_flagged = False
        if any(c in tel for c in ['(', ')', '/']) or not tel.isdigit() or len(tel) != 10:
            is_flagged = True
        if any(c in contact for c in ['(', ')', '/']) or not contact.isdigit() or len(contact) != 10:
            is_flagged = True

        # Update cleaned values
        df.at[index, 'Telephone'] = tel
        df.at[index, 'Contact person -Telephone'] = contact

        # Clear contact if same as tel and not flagged
        if tel == contact and not is_flagged:
            df.at[index, 'Contact person -Telephone'] = ''

        flagged.append('flagged' if is_flagged else '')

    df['Flagged'] = flagged
    return df

# ---------- 2. Clean Names Function ----------
def clean_names(df):
    for index, row in df.iterrows():
        first = str(row['Contact person -First name']) if not pd.isna(row['Contact person -First name']) else ''
        last = str(row['Contact person -Last name']) if not pd.isna(row['Contact person -Last name']) else ''

        if first.strip().lower() in ['mr', 'mr.']:
            df.at[index, 'Contact person -First name'] = last
            df.at[index, 'Contact person -Last name'] = ''
        else:
            df.at[index, 'Contact person -First name'] = first
            df.at[index, 'Contact person -Last name'] = last

    return df

# ---------- 3. Clean City Function ----------
def clean_city_names(df):
    for index, row in df.iterrows():
        city = str(row['City']) if not pd.isna(row['City']) else ''

        # Keep only leading letters
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

# ---------- 4. Main Execution ----------
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
