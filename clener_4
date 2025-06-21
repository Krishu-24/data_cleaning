import pandas as pd

file_input = "C:/Users/krish/Downloads/Customer_list_06102025063059.xlsx"
file_output = "C:/Users/krish/Downloads/cleaned.xlsx"
file_flagged = "C:/Users/krish/Downloads/flagged_only.xlsx"
col_telephone = "Telephone"
col_contact_tel = "Contact person -Telephone"
col_first_name = "Contact person -First name"
col_last_name = "Contact person -Last name"
col_city = "City"

def clean_nums(df, col_n):
    flagged_list = []
    cleaned_col = []
    rem = ["-", "+91", " "]
    target = ["(", ".", ")", "&"]
    for i, row in df.iterrows():
        flag = False
        if pd.isna(row[col_n]):
            num = ""
        else:
            num = str(row[col_n])
        for r in rem:
            num = num.replace(r, "")
        if any(t in str(row[col_n]) for t in target):
            flag = True
        elif not num.isdigit() or len(num) != 10:
            flag = True
        flagged_list.append("flagged" if flag else "")
        cleaned_col.append(num)
    df[col_n] = cleaned_col
    df["flagged_" + col_n] = flagged_list
    return df

def compare_nums(df):
    cl = []
    for i, row in df.iterrows():
        t = str(row[col_telephone]).strip()
        c = str(row[col_contact_tel]).strip()
        if t == c:
            cl.append("")
        else:
            cl.append(c)
    df[col_contact_tel] = cl
    return df

def clean_names(df, fn, ln):
    p = ["mr", "mr.", "mr .", "mrs", "mrs.", "mrs ."]
    fn_l = []
    ln_l = []
    for i, row in df.iterrows():
        f = str(row[fn]).strip().lower() if not pd.isna(row[fn]) else ''
        l = str(row[ln]).strip().lower() if not pd.isna(row[ln]) else ''
        if f in p:
            fn_l.append(l)
            ln_l.append("")
        else:
            for x in p:
                if f.startswith(x):
                    f = f.replace(x, "").strip()
                    break
            fn_l.append(f)
            ln_l.append(l)
    df[fn] = fn_l
    df[ln] = ln_l
    return df

def clean_city_names(df):
    for i, row in df.iterrows():
        c = str(row[col_city]) if not pd.isna(row[col_city]) else ''
        clean = ''
        for ch in c:
            if ch.isalpha():
                clean += ch
            else:
                break
        if clean:
            clean = clean[0].upper() + clean[1:].lower()
        df.at[i, col_city] = clean
    return df

def export_flagged_rows(df):
    cols = [c for c in df.columns if c.startswith("flagged_")]
    if cols:
        flagged = df[df[cols].apply(lambda x: "flagged" in x.values, axis=1)]
        flagged.to_excel(file_flagged, index=False)

def main():
    df = pd.read_excel(file_input)
    df = clean_nums(df, col_telephone)
    df = clean_nums(df, col_contact_tel)
    df = compare_nums(df)
    df = clean_names(df, col_first_name, col_last_name)
    df = clean_city_names(df)
    df.to_excel(file_output, index=False)
    export_flagged_rows(df)

main()
