import pandas as pd
import joblib
import openpyxl
import numpy as np


# ==============================
# UTILITAIRES FEATURES
# ==============================

def max_consecutive_empty(values):
    max_count = 0
    current = 0
    for v in values:
        if v is None:
            current += 1
            max_count = max(max_count, current)
        else:
            current = 0
    return max_count


def extract_row_features_from_row(row, total_cols, ws, row_idx):

    values = [cell.value for cell in row]
    non_empty = [v for v in values if v is not None]

    if not non_empty:
        return None

    fill_ratio = len(non_empty) / total_cols
    str_ratio = sum(isinstance(v, str) for v in non_empty) / len(non_empty)
    num_ratio = sum(isinstance(v, (int, float)) for v in non_empty) / len(non_empty)

    bold_ratio = sum(
        1 for c in row if getattr(c, "font", None) and c.font.bold
    ) / total_cols

    colored_ratio = sum(
        1 for c in row
        if getattr(c, "fill", None)
        and c.fill.fgColor
        and c.fill.fgColor.rgb != "00000000"
    ) / total_cols

    row_position = row_idx / ws.max_row

    str_lengths = [len(str(v)) for v in non_empty if isinstance(v, str)]
    avg_str_len = np.mean(str_lengths or [0])
    std_str_len = np.std(str_lengths or [0])

    header_keywords = ["date", "name", "amount", "total", "id", "ref"]
    keyword_hits = sum(
        any(k in str(v).lower() for k in header_keywords)
        for v in non_empty if isinstance(v, str)
    )
    keyword_ratio = keyword_hits / len(non_empty)

    max_empty_streak = max_consecutive_empty(values) / total_cols
    unique_ratio = len(set(non_empty)) / len(non_empty)

    upper_ratio = sum(
        str(v).isupper() for v in non_empty if isinstance(v, str)
    ) / len(non_empty)

    special_char_ratio = sum(
        any(c in str(v) for c in ",.;:()")
        for v in non_empty if isinstance(v, str)
    ) / len(non_empty)

    num_to_str_ratio = num_ratio - str_ratio

    return [
        fill_ratio,
        str_ratio,
        num_ratio,
        bold_ratio,
        colored_ratio,
        row_position,
        avg_str_len,
        std_str_len,
        keyword_ratio,
        max_empty_streak,
        unique_ratio,
        upper_ratio,
        special_char_ratio,
        num_to_str_ratio
    ]


# ==============================
# PREDICTION HEADER
# ==============================

def predict_header_row(filepath, model):
    wb = openpyxl.load_workbook(filepath, data_only=True)
    results = {}

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        total_cols = ws.max_column
        candidates = []

        for row_idx, row in enumerate(ws.iter_rows(max_row=30), start=1):
            features = extract_row_features_from_row(row, total_cols, ws, row_idx)

            if features is not None:
                proba = model.predict_proba([features])[0][1]
                candidates.append((row_idx, proba))

        if candidates:
            best_row = max(candidates, key=lambda x: x[1])
            results[sheet_name] = best_row[0]

    return results


# ==============================
# UTILISATION DU MODELE
# ==============================

model = joblib.load("header_detector.pkl")

target = target="C:/Users/docs.xlsm"

header_predictions = predict_header_row(target, model)

print("Header détectés :")
print(header_predictions)


# ==============================
# LECTURE PROPRE DU FICHIER
# ==============================

dfs = {}

for sheet_name, header_row in header_predictions.items():
    try:
        df = pd.read_excel(
            target,
            sheet_name=sheet_name,
            header=header_row - 1  # pandas commence à 0
        )
        dfs[sheet_name] = df
        print(f"✅ {sheet_name} -> header ligne {header_row}")

    except Exception as e:
        print(f"❌ Erreur lecture {sheet_name} :", e)

