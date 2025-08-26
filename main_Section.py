import streamlit as st
import pandas as pd
import re
from rapidfuzz import fuzz

def normalize_text(text):
    if pd.isnull(text):
        return ""
    text = str(text).strip().replace("ه", "ة").replace("أ", "ا").replace("إ", "ا").replace("آ", "ا")
    text = re.sub(r'(عبد)([^\s])', r'\1 \2', text)
    return " ".join(text.split()).lower()

def is_first_three_words_match(name1, name2):
    words1 = name1.split()
    words2 = name2.split()
    length = min(len(words1), len(words2), 3)
    return all(words1[i] == words2[i] for i in range(length))

def match_names(names_df, database_df):
    names_df["normalized_name"] = names_df["اسم الموظف"].apply(normalize_text)
    database_df["normalized_name"] = database_df["اسم الموظف"].apply(normalize_text)

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
        if best_match and best_score >= 85 and (is_first_three_words_match(normalized_name, best_match) or best_match.startswith(normalized_name)):
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
                "الاسم المطابق": match_data.get("اسم الموظف", ""),
                "الآيبان": match_data.get("Iban", ""),
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
    return results_df

def add_accountant_info(results_df, accountants_df):
    # تنظيف عمود الدائرة في كلا الملفين
    results_df["normalized_section"] = results_df["الدائرة"].apply(normalize_text)
    accountants_df["normalized_section"] = accountants_df["الدائرة"].apply(normalize_text)

    # إضافة عمود فارغ
    results_df["اليوزر"] = ""
    results_df["اسم المحاسب"] = ""

    # مطابقة تقريبية على الدائرة
    for i, res_row in results_df.iterrows():
        best_score = 0
        best_index = None
        for j, acc_row in accountants_df.iterrows():
            score = fuzz.partial_ratio(res_row["normalized_section"], acc_row["normalized_section"])
            if score > best_score:
                best_score = score
                best_index = j
        if best_score >= 80:  # نسبة ثقة عالية
            results_df.at[i, "اليوزر"] = accountants_df.at[best_index, "اليوزر"]
            results_df.at[i, "اسم المحاسب"] = accountants_df.at[best_index, "اسم المحاسب"]

    results_df.drop(columns=["normalized_section"], inplace=True)
    return results_df

def main():
    st.set_page_config(page_title="برنامج المطابقة", layout="wide")
    st.title("برنامج المطابقة للأسماء والأقسام مع ربط المحاسبين")

    uploaded_names = st.file_uploader("اختر ملف الأسماء (Excel)", type=["xlsx"])
    uploaded_db = st.file_uploader("اختر ملف قاعدة البيانات (Excel)", type=["xlsx"])
    uploaded_accounts = st.file_uploader("اختر ملف المحاسبين (Excel)", type=["xlsx"])

    if uploaded_names and uploaded_db and uploaded_accounts:
        if st.button("بدء المطابقة"):
            names_df = pd.read_excel(uploaded_names)
            db_df = pd.read_excel(uploaded_db)
            accounts_df = pd.read_excel(uploaded_accounts)

            with st.spinner("جاري المطابقة..."):
                results_df = match_names(names_df, db_df)
                final_df = add_accountant_info(results_df, accounts_df)

            st.success("تمت المطابقة وإضافة المحاسبين!")
            st.dataframe(final_df)

            csv = final_df.to_csv(index=False).encode('utf-8-sig')
            st.download_button(
                label="تحميل النتائج بصيغة CSV",
                data=csv,
                file_name="نتائج_مطابقة_مع_المحاسبين.csv",
                mime="text/csv",
            )

if __name__ == "__main__":
    main()
