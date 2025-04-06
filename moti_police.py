import pandas as pd
import streamlit as st
import re as re

# המרת אות לאינדקס (בסיס 0) לפי Excel
def letter_to_number(col):
    col = col.upper()
    result = 0
    for c in col:
        result = result * 26 + (ord(c) - ord('A') + 1)
    return result - 1

# חילוץ מילון ת"ז -> שם מלא
def extract_id_name_dict_from_column(series):
    id_name_dict = {}
    for entry in series.dropna():
        parts = str(entry).strip().split()
        if len(parts) >= 3:
            id_number = parts[0]
            full_name = " ".join(parts[1:])
            id_name_dict[id_number] = full_name
    return id_name_dict

# ספירת הופעות לפי ת"ז (גם אם תת־מחרוזת בתא)
def count_id_occurrences_exact_in_text(id_dict, id_column_series):
    id_column_series = id_column_series.astype(str)
    result = {full_name: 0 for full_name in id_dict.values()}

    for cell in id_column_series:
        for id_key, full_name in id_dict.items():
            pattern = r'\b{}\b'.format(re.escape(str(id_key)))
            if re.search(pattern, cell):
                result[full_name] += 1

    return result


# Streamlit UI
st.set_page_config(page_title="דו\"ח נוכחות", layout="centered", page_icon="📊")
st.title("📊 דו\"ח נוכחות לפי שם")

uploaded_file = st.file_uploader("בחר קובץ Excel (.xlsx)", type=["xls", "xlsx"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)

        # שמות עמודות לפי אותיות: C, L, V
        df = df.iloc[:, [letter_to_number('C'), letter_to_number('L'), letter_to_number('V')]]

        # חילוץ ת"ז → שם מלא
        col_c = df.iloc[:, 0]
        id_name_dict = extract_id_name_dict_from_column(col_c)

        # ספירה כללית (עמודה V)
        attendance_all = count_id_occurrences_exact_in_text(id_name_dict, df.iloc[:, 2])

        # פילוח לפי משמרת
        yes_shift = df[df.iloc[:, 1].astype(str) == "משמרת"].iloc[:, 0]
        no_shift = df[df.iloc[:, 1].astype(str) != "משמרת"].iloc[:, 0]

        attendance_shift = count_id_occurrences_exact_in_text(id_name_dict, yes_shift)
        attendance_no_shift = count_id_occurrences_exact_in_text(id_name_dict, no_shift)

        def dict_to_df(d):
            # סינון שמות שמספר ההופעות שלהם גדול מ-0
            filtered = {name: count for name, count in d.items() if count > 0}
            df_out = pd.DataFrame(list(filtered.items()), columns=["שם", "מספר הופעות"]).reset_index(drop=True)
            return df_out.sort_values(by="מספר הופעות", ascending=False)



        st.subheader("סהכ מספר אחמושים לאחמש")
        df_clean = dict_to_df(attendance_all).copy()
        df_clean.index = [''] * len(df_clean)
        st.dataframe(df_clean)

        st.subheader("נוכחות במשמרת")
        df_clean = dict_to_df(attendance_shift).copy()
        df_clean.index = [''] * len(df_clean)
        st.dataframe(df_clean)

        st.subheader("נוכחות שלא במשמרת")
        df_clean = dict_to_df(attendance_no_shift).copy()
        df_clean.index = [''] * len(df_clean)
        st.dataframe(df_clean)

    except Exception as e:
        st.error(f"שגיאה בטעינת הקובץ: {e}")
