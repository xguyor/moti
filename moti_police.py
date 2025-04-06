import pandas as pd
import streamlit as st

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
def count_id_occurrences(id_dict, id_column_series):
    id_column_series = id_column_series.astype(str)
    result = {full_name: 0 for full_name in id_dict.values()}
    for cell in id_column_series:
        for id_key, full_name in id_dict.items():
            if str(id_key) in cell:
                result[full_name] += 1
    return result

# Streamlit UI
st.set_page_config(page_title="דו\"ח נוכחות", layout="centered", page_icon="📊")
st.title("📊 דו\"ח נוכחות לפי ת\"ז")

uploaded_file = st.file_uploader("בחר קובץ Excel (.xks / .xlsx)", type=["xls", "xlsx", "xks"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)

        # שמות עמודות לפי אותיות: C, L, V
        df = df.iloc[:, [letter_to_number('C'), letter_to_number('L'), letter_to_number('V')]]

        # חילוץ ת"ז → שם מלא
        col_c = df.iloc[:, 0]
        id_name_dict = extract_id_name_dict_from_column(col_c)

        # ספירה כללית (עמודה V)
        attendance_all = count_id_occurrences(id_name_dict, df.iloc[:, 2])

        # פילוח לפי משמרת
        yes_shift = df[df.iloc[:, 1].astype(str) == "משמרת"].iloc[:, 0]
        no_shift = df[df.iloc[:, 1].astype(str) != "משמרת"].iloc[:, 0]

        attendance_shift = count_id_occurrences(id_name_dict, yes_shift)
        attendance_no_shift = count_id_occurrences(id_name_dict, no_shift)

        # שורת חיפוש
        search = st.text_input("🔍 חפש לפי שם")

        def dict_to_df(d):
            df_out = pd.DataFrame(list(d.items()), columns=["שם", "מספר הופעות"])
            if search:
                df_out = df_out[df_out["שם"].str.contains(search, case=False)]
            return df_out.sort_values(by="מספר הופעות", ascending=False)

        st.subheader("עמודה 500 ")
        st.dataframe(dict_to_df(attendance_all), use_container_width=True)

        st.subheader("נוכחות במשמרת")
        st.dataframe(dict_to_df(attendance_shift), use_container_width=True)

        st.subheader("נוכחות שלא במשמרת")
        st.dataframe(dict_to_df(attendance_no_shift), use_container_width=True)

    except Exception as e:
        st.error(f"שגיאה בטעינת הקובץ: {e}")
