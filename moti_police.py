import pandas as pd
import streamlit as st
import re as re
# from tkinter import Tk, filedialog


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

    # במקום לשמור מפתחות לפי "full_name", שומרים אותם לפי "id_key" (הת"ז)
    result = {id_key: 0 for id_key in id_dict.keys()}

    for cell in id_column_series:
        for id_key in id_dict.keys():
            pattern = r'\b{}\b'.format(re.escape(str(id_key)))
            if re.search(pattern, cell):
                result[id_key] += 1

    return result

def dict_to_df(d, id_name_dict):
    # סינון שמות שמספר ההופעות שלהם גדול מ-0 (רק אם המספר חוקי)
    filtered = {
        the_id: int(count)
        for the_id, count in d.items()
        if isinstance(count, (int, float)) and count > 0
    }

    # בניית טבלה בסיסית: עמודה 'ID' ו'מספר אירועים'
    df_out = pd.DataFrame(list(filtered.items()), columns=["ID", "מספר אירועים"]).reset_index(drop=True)

    # ממפים את ID -> שם (או "לא ידוע" אם חסר במילון)
    df_out["שם"] = df_out["ID"].map(id_name_dict).fillna("לא ידוע")

    # נרצה רק את עמודות "שם" ו"מספר אירועים"
    df_out = df_out[["שם", "מספר אירועים"]].sort_values(by="מספר אירועים", ascending=False).reset_index(drop=True)

    return df_out
def dict_to_df_col(d, col_name):
    items = list(d.items())  # [(tz, value), ...]
    return pd.DataFrame(items, columns=["tz", col_name])

def extend_df_with_columns(df_shift, attendance_dict, id_name_dict):
    df_extended = pd.DataFrame()
    df_extended["שם"] = df_shift.iloc[:, 0].map(id_name_dict)
    df_extended["מספר אירועים"] = df_shift.iloc[:, 0].map(
        lambda id_val: attendance_dict.get(id_name_dict.get(id_val, ""), 0)
    )

    # מצרף את העמודות 1–5 (שהן העמודות הנוספות O..S)
    df_extended = pd.concat([df_extended, df_shift.iloc[:, 1:]], axis=1)

    # מסנן הופעות אפס אם רוצים
    df_extended = df_extended[df_extended["מספר אירועים"] > 0]

    return df_extended.reset_index(drop=True)

def sum_by_id_in_text_column(df):

    # המרת העמודה השנייה למספרים (NaN -> 0)
    text_series = df.iloc[:, 0].astype(str)
    numeric_series = pd.to_numeric(df.iloc[:, 1], errors="coerce").fillna(0)

    sums = {}  # מילון: key=ID (מחרוזת), value=סכום

    for i in range(len(df)):
        text_val = text_series.iloc[i]
        num_val = numeric_series.iloc[i]

        # חיפוש כל מופע של מספרים (ת"ז וכדומה) עם גבולות מילה
        found_ids = re.findall(r'\b(\d+)\b', text_val)
        # לדוגמה, במחרוזת "blah 123 hello 939" -> ["123", "939"]

        # אם יכולות להיות כפילויות באותה שורה (למשל "123 123"),
        # ונרצה להוסיף רק פעם אחת - נמיר את found_ids לסט
        # אם דווקא כן חשוב לספור כל מופע, נשאיר כמו שהוא.
        found_ids = set(found_ids)

        for f_id in found_ids:
            # עדכן את הסכום
            sums[f_id] = sums.get(f_id, 0) + num_val

    return sums

column_dict = {
    "O" : "פרטי/מסחרי",
    "P" : "משאית עד 15 טון",
    "Q" : "משאית מעל 15 טון",
    "R" : "פולטריילר",
    "S" : "אוטובוס",
    "T" : "הפוך"
}
def merge_attendance_dicts(
        count_dict,
        dict_O,
        dict_P,
        dict_Q,
        dict_R,
        dict_S,
        dict_T,
        id_name_dict
):

    def dict_to_df_col(d, col_name):
        # ניקוי רווחים סביב המפתחות (ת"ז)
        cleaned = {}
        for k, v in d.items():
            k_str = str(k).strip()
            cleaned[k_str] = v

        items = list(cleaned.items())  # [(tz, value), ...]
        df_out = pd.DataFrame(items, columns=["tz", col_name])
        return df_out

    # הופכים כל מילון ל-DataFrame עם שם עמודה מתאים
    dfCount = dict_to_df_col(count_dict, "מספר אירועים")
    dfO = dict_to_df_col(dict_O, column_dict["O"])
    dfP = dict_to_df_col(dict_P, column_dict["P"])
    dfQ = dict_to_df_col(dict_Q, column_dict["Q"])
    dfR = dict_to_df_col(dict_R, column_dict["R"])
    dfS = dict_to_df_col(dict_S, column_dict["S"])
    dfT = dict_to_df_col(dict_S, column_dict["T"])

    # איחוד (merge) בכל השלבים
    df_merged = (
        dfCount
        .merge(dfO, how="outer", on="tz")
        .merge(dfP, how="outer", on="tz")
        .merge(dfQ, how="outer", on="tz")
        .merge(dfR, how="outer", on="tz")
        .merge(dfS, how="outer", on="tz")
        .merge(dfT, how="outer", on="tz")
    )

    # מילוי NaN ב-0 כדי לא לקבל ערכים חסרים בעמודות המספריות
    df_merged = df_merged.fillna(0)

    # אם יש לנו מילון ת"ז -> שם
    if id_name_dict is not None:
        # הוספת עמודת 'שם' מתוך ת"ז
        df_merged["שם"] = df_merged["tz"].map(id_name_dict).fillna("לא ידוע")
    else:
        # אחרת, נשתמש בעמודת tz כבסיס ל-"שם"
        df_merged["שם"] = df_merged["tz"]

    # המרת עמודות מספריות ל-int (אחרי fillna(0))
    numeric_cols = ["מספר אירועים", column_dict["O"], column_dict["P"], column_dict["Q"], column_dict["R"], column_dict["S"], column_dict["T"]]
    for col in numeric_cols:
        df_merged[col] = pd.to_numeric(df_merged[col], errors="coerce").fillna(0).astype(int)

    # סידור העמודות: [שם, מספר אירועים, O, P, Q, R, S]
    df_merged = df_merged[["שם"] + numeric_cols]

    # מיון לפי 'מספר אירועים' בסדר יורד
    df_merged = df_merged.sort_values(by="מספר אירועים", ascending=False).reset_index(drop=True)

    return df_merged

# def get_file_from_tkinter():
#     root = Tk()
#     root.withdraw()  # לא להציג את החלון הראשי
#     file_path = filedialog.askopenfilename(
#         title="בחר קובץ Excel",
#         filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
#     )
#     return file_path

# uploaded_file = get_file_from_tkinter()
# Streamlit UI
st.set_page_config(page_title="דו\"ח נוכחות", layout="wide", page_icon="📊")
# הפעלת RTL ויישור לימין על כל הדף
st.markdown(
    """
    <style>
        html, body, [class*="css"]  {
            direction: rtl;
            text-align: right;
        }
    </style>
    """,
    unsafe_allow_html=True
)

# כותרת ראשית
st.markdown('<h1>📊 ניתוח דוחות יחפ"צ</h1>', unsafe_allow_html=True)
uploaded_file = st.file_uploader("בחר קובץ Excel (.xlsx)", type=["xlsx"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)

        # שמות עמודות לפי אותיות: C, L, V
        df = df.iloc[:, [letter_to_number('C'), letter_to_number('L'), letter_to_number('V'),
                         letter_to_number('O'), letter_to_number('P'), letter_to_number('Q'),
                         letter_to_number('R'),   letter_to_number('S'), letter_to_number('T')]]

        # חילוץ ת"ז → שם מלא
        col_c = df.iloc[:, 0]

        id_name_dict = extract_id_name_dict_from_column(col_c)

        # ספירה כללית (עמודה V)
        col_v_no_shift = df[df.iloc[:, 1].astype(str) != "משמרת"].iloc[:, 2]
        attendance_all = count_id_occurrences_exact_in_text(id_name_dict, col_v_no_shift)

        # פילוח לפי משמרת
        yes_shift = df[df.iloc[:, 1].astype(str) == "משמרת"].iloc[:, [0, 3, 4, 5, 6, 7, 8]]
        attendance_shift = count_id_occurrences_exact_in_text(id_name_dict, yes_shift.iloc[:, 0])
        attendance_shift_O = sum_by_id_in_text_column(yes_shift.iloc[:, [0,1]])
        attendance_shift_P = sum_by_id_in_text_column(yes_shift.iloc[:, [0,2]])
        attendance_shift_Q = sum_by_id_in_text_column(yes_shift.iloc[:, [0,3]])
        attendance_shift_R = sum_by_id_in_text_column(yes_shift.iloc[:, [0,4]])
        attendance_shift_S = sum_by_id_in_text_column(yes_shift.iloc[:, [0,5]])
        attendance_shift_T = sum_by_id_in_text_column(yes_shift.iloc[:, [0,6]])

        df_merged_shift = merge_attendance_dicts(
            attendance_shift,
            attendance_shift_O,
            attendance_shift_P,
            attendance_shift_Q,
            attendance_shift_R,
            attendance_shift_S,
            attendance_shift_T,
            id_name_dict
        )

        no_shift = df[df.iloc[:, 1].astype(str) != "משמרת"].iloc[:, [0, 3, 4, 5, 6, 7, 8]]
        attendance_no_shift = count_id_occurrences_exact_in_text(id_name_dict, no_shift.iloc[:, 0])
        attendance_no_shift_O = sum_by_id_in_text_column(no_shift.iloc[:, [0,1]])
        attendance_no_shift_P = sum_by_id_in_text_column(no_shift.iloc[:, [0,2]])
        attendance_no_shift_Q = sum_by_id_in_text_column(no_shift.iloc[:, [0,3]])
        attendance_no_shift_R = sum_by_id_in_text_column(no_shift.iloc[:, [0,4]])
        attendance_no_shift_S = sum_by_id_in_text_column(no_shift.iloc[:, [0,5]])
        attendance_no_shift_T = sum_by_id_in_text_column(no_shift.iloc[:, [0,6]])

        df_merged_no_shift = merge_attendance_dicts(
            attendance_no_shift,
            attendance_no_shift_O,
            attendance_no_shift_P,
            attendance_no_shift_Q,
            attendance_no_shift_R,
            attendance_no_shift_S,
            attendance_no_shift_T,
            id_name_dict
        )

        # הצגת תוצאות
        st.markdown(
            '<h3 style="text-align: right;">סה&quot;כ אחמושים לאחמש</h3>',
            unsafe_allow_html=True
        )
        df_clean = dict_to_df(attendance_all, id_name_dict).copy()
        df_clean.index = [''] * len(df_clean)
        st.dataframe(df_clean, use_container_width=False)


        st.markdown(
            '<h3 style="text-align: right;">סה&quot;כ משמרות ואירועים במשמרת לפי סוג</h3>',
            unsafe_allow_html=True
        )
        df_merged_shift.index = [''] * len(df_merged_shift)
        st.dataframe(df_merged_shift, use_container_width=False)

        st.markdown(
            '<h3 style="text-align: right;">סה&quot;כ אירועים לפי סוג לא במשמרת</h3>',
            unsafe_allow_html=True
        )
        df_merged_no_shift.index = [''] * len(df_merged_no_shift)
        st.dataframe(df_merged_no_shift, use_container_width=False)

    except Exception as e:
        st.error(f"שגיאה בטעינת הקובץ: {e}")
