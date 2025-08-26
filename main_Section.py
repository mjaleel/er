import streamlit as st
import pandas as pd
import hashlib
import re
from rapidfuzz import fuzz

def normalize_name(name):
    if pd.isnull(name):
        return ""
    name = name.strip().replace("ه", "ة").replace("أ", "ا").replace("إ", "ا").replace("آ", "ا")
    name = re.sub(r'(عبد)([^\s])', r'\1 \2', name)
    return " ".join(name.split()).lower()

def is_first_three_words_match(name1, name2):
    words1 = name1.split()
    words2 = name2.split()
    length = min(len(words1), len(words2), 3)
    return all(words1[i] == words2[i] for i in range(length))

def match_names(names_df, database_df):
    names_df["normalized_name"] = names_df["اسم الموظف"].apply(normalize_name)
    database_df["normalized_name"] = database_df["اسم الموظف"].apply(normalize_name)

    database_df = database_df.drop_duplicates(subset=["normalized_name"])
    database_map = database_df.set_index("normalized_name")[["اسم الموظف", "Iban"]].to_dict(orient="index")

    matched_results = []

    for original_name, normalized_name in zip(names_df["اسم الموظف"], names_df["normalized_name"]):
        best_match = None
        best_score = 0
        for db_name in database_map.keys():
            score = fuzz.ratio(normalized_name, db_name)
            if score > best_score:
                best_score = score
                best_match = db_name

        match_data = None
        if best_score >= 85 and (is_first_three_words_match(normalized_name, best_match) or best_match.startswith(normalized_name)):
            match_data = database_map[best_match]
        else:
            for db_name in database_map.keys():
                if db_name.startswith(normalized_name):
                    match_data = database_map[db_name]
                    best_match = db_name
                    best_score = fuzz.ratio(normalized_name, best_match)
                    break

        if match_data:
            matched_results.append({
                "الاسم الأصلي": original_name,
                "الاسم المطابق": match_data["اسم الموظف"],
                "الآيبان": match_data["Iban"],
                "نسبة التطابق": f"{round(best_score)}%",
                "ملاحظة": "✅ تطابق دقيق"
            })
        else:
            matched_results.append({
                "الاسم الأصلي": original_name,
                "الاسم المطابق": "",
                "الآيبان": "",
                "نسبة التطابق": "",
                "ملاحظة": "❌ لم يتم العثور على تطابق"
            })

    results_df = pd.DataFrame(matched_results)

    # تنبيه التكرار
    iban_counts = results_df["الآيبان"].value_counts()
    results_df["تنبيه"] = results_df["الآيبان"].apply(lambda x: "⚠️ مكرر" if pd.notnull(x) and iban_counts[x] > 1 else "")

    return results_df

def match_sections(names_df, database_df):
    names_df["normalized_name"] = names_df["اسم الموظف"].apply(normalize_name)
    database_df["normalized_name"] = database_df["اسم الموظف"].apply(normalize_name)

    database_df = database_df.drop_duplicates(subset=["normalized_name"])
    database_map = database_df.set_index("normalized_name")[[
        "اسم الموظف",
        "Operator Id",
        "الدرجة الوظيفية",
        "العنوان الوظيفي",
        "المدرسة",
        "الدائرة"
    ]].to_dict(orient="index")

    matched_results = []

    for original_name, normalized_name in zip(names_df["اسم الموظف"], names_df["normalized_name"]):
        best_match = None
        best_score = 0
        for db_name in database_map.keys():
            score = fuzz.ratio(normalized_name, db_name)
            if score > best_score:
                best_score = score
                best_match = db_name

        match_data = None
        if best_score >= 85 and (is_first_three_words_match(normalized_name, best_match) or best_match.startswith(normalized_name)):
            match_data = database_map[best_match]
        else:
            for db_name in database_map.keys():
                if db_name.startswith(normalized_name):
                    match_data = database_map[db_name]
                    best_match = db_name
                    best_score = fuzz.ratio(normalized_name, best_match)
                    break

        if match_data:
            matched_results.append({
                "الاسم الأصلي": original_name,
                "الاسم المطابق": match_data["اسم الموظف"],
                "Operator Id": match_data.get("Operator Id", ""),
                "الدرجة الوظيفية": match_data.get("الدرجة الوظيفية", ""),
                "العنوان الوظيفي": match_data.get("العنوان الوظيفي", ""),
                "المدرسة": match_data.get("المدرسة", ""),
                "الدائرة": match_data.get("الدائرة", ""),
                "نسبة التطابق": f"{round(best_score)}%",
                "ملاحظة": "✅ تطابق دقيق"
            })
        else:
            matched_results.append({
                "الاسم الأصلي": original_name,
                "الاسم المطابق": "",
                "Operator Id": "",
                "الدرجة الوظيفية": "",
                "العنوان الوظيفي": "",
                "المدرسة": "",
                "الدائرة": "",
                "نسبة التطابق": "",
                "ملاحظة": "❌ لم يتم العثور على تطابق"
            })

    return pd.DataFrame(matched_results)

def main():
    st.set_page_config(page_title="برنامج المطابقة", layout="wide")
    st.title("برنامج المطابقة للأسماء والأقسام")

    tab_names, tab_sections = st.tabs(["مطابقة الأسماء", "مطابقة الأقسام"])

    with tab_names:
        st.header("مطابقة الأسماء")
        uploaded_names = st.file_uploader("اختر ملف الأسماء (Excel)", type=["xlsx"])
        uploaded_db = st.file_uploader("اختر ملف قاعدة البيانات (Excel)", type=["xlsx"])
        if uploaded_names and uploaded_db:
            if st.button("بدء المطابقة (الأسماء)"):
                names_df = pd.read_excel(uploaded_names)
                db_df = pd.read_excel(uploaded_db)
                with st.spinner("جاري المطابقة..."):
                    results_df = match_names(names_df, db_df)
                st.success("تمت المطابقة!")
                st.dataframe(results_df)

                csv = results_df.to_csv(index=False).encode('utf-8-sig')
                st.download_button(
                    label="تحميل النتائج بصيغة CSV",
                    data=csv,
                    file_name="نتائج_مطابقة_الأسماء.csv",
                    mime="text/csv",
                )

    with tab_sections:
        st.header("مطابقة الأقسام")
        uploaded_names_sec = st.file_uploader("اختر ملف الأسماء (Excel)", type=["xlsx"], key="names_sec")
        uploaded_db_sec = st.file_uploader("اختر ملف قاعدة البيانات (Excel)", type=["xlsx"], key="db_sec")
        if uploaded_names_sec and uploaded_db_sec:
            if st.button("بدء المطابقة (الأقسام)"):
                names_df_sec = pd.read_excel(uploaded_names_sec)
                db_df_sec = pd.read_excel(uploaded_db_sec)
                with st.spinner("جاري المطابقة..."):
                    results_sec = match_sections(names_df_sec, db_df_sec)
                st.success("تمت المطابقة!")
                st.dataframe(results_sec)

                csv_sec = results_sec.to_csv(index=False).encode('utf-8-sig')
                st.download_button(
                    label="تحميل النتائج بصيغة CSV",
                    data=csv_sec,
                    file_name="نتائج_مطابقة_الأقسام.csv",
                    mime="text/csv",
                )

if __name__ == "__main__":
    main()
