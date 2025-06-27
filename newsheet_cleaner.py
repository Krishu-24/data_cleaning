import pandas as pd

col = "GSTIN/UIN"
add = "Buyer Address"
name = "Buyer"
file_path = "your_file_path.xlsx"  # Update this

def flag(df):
    df["flag"] = df[col].apply(lambda x: "flag" if pd.isna(x) else "")
    return df

def check_gst(df):
    duplicates = []
    seen = {}

    for idx, row in df.iterrows():
        gstin = row[col]
        if pd.isna(gstin):
            continue
        key = (gstin, row[name], row[add])
        if gstin in seen:
            for prev_idx in seen[gstin]:
                prev_row = df.loc[prev_idx]
                if row[name] == prev_row[name] or row[add] == prev_row[add]:
                    duplicates.append(idx)
                    break
            seen[gstin].append(idx)
        else:
            seen[gstin] = [idx]

    df = df.drop(duplicates)
    return df

def main():
    df = pd.read_excel(file_path)
    df = flag(df)
    df = check_gst(df)
    df.to_excel("cleaned_output.xlsx", index=False)

main()
