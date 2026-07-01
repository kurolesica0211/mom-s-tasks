import re
import pandas as pd

file_path = "Список платежей Новоросс РКО.xlsx"
df = pd.read_excel(file_path)


def clean_number(num_str):
    num_str = str(num_str).replace(" ", "")
    num_str = num_str.replace(",", "")

    try:
        return float(num_str)
    except ValueError:
        return num_str


def clean_number_from_text(text, num_str):
    num_str = str(num_str).replace(" ", "")

    if isinstance(text, str) and text.startswith("//Реестр//"):
        num_str = num_str.replace(",", ".")
    else:
        num_str = num_str.replace(",", "")

    try:
        return float(num_str)
    except ValueError:
        return num_str


def extract_commission(text):
    if not isinstance(text, str):
        return None

    commission_patterns = [
        r"\(Комиссия\s*([\d\s.,]+)\s*руб,\s*в т\.ч\. НДС",
        r"комиссия:\s*([\d\s.,]+);",
        r"Комиссия\s*([\d\s.,]+?)(?=(?:\. Возврат|\(в т\.ч\. НДС|\. НДС не облагается\.|\. Без НДС\b|\s*руб\.?\s*Без НДС\b))",
    ]

    for pattern in commission_patterns:
        match = re.search(pattern, text, flags=re.IGNORECASE)
        if match:
            if pattern.startswith(r"\(Комиссия"):
                return clean_number_from_text(text, match.group(1))
            return clean_number_from_text(text, match.group(1))

    return None


def extract_return(row):
    text = row["Назначение платежа"]
    if not isinstance(text, str):
        return None

    return_patterns = [
        r"возврат\s*[:=]\s*([\d\s.,]+)",
        r"Возврат покупки\s*([\d\s.,]+)",
    ]

    for pattern in return_patterns:
        match = re.search(pattern, text, flags=re.IGNORECASE)
        if match:
            return clean_number_from_text(text, match.group(1).split("/")[0])

    return None


commission_values = df["Назначение платежа"].apply(extract_commission)
return_values = df.apply(extract_return, axis=1)

if "Комиссия" in df.columns:
    df = df.drop(columns=["Комиссия"])

if "Возврат" in df.columns:
    df = df.drop(columns=["Возврат"])

insert_at = df.columns.get_loc("Поступление") + 1
df.insert(insert_at, "Комиссия", commission_values)
df.insert(insert_at + 1, "Возврат", return_values)

output_path = file_path.split(".xlsx")[0] + "_ред.xlsx"
df.to_excel(output_path, index=False)

print(f"Done! The updated table has been saved to '{output_path}'")