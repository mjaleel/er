import streamlit as st
import pandas as pd
from rapidfuzz import fuzz
import re

# دالة لتنظيف النصوص (الأسماء والدائرة)
def normalize_text(text):
    if pd.isnull(text):
        return ""
    text = str(text).strip().replace("ه", "ة").replace("أ", "ا").replace("إ", "ا").replace("آ", "ا")
    text = re.sub(r'(عبد)([^\s])', r'\1 \2', text)
    return " ".join(text.split()).lower()

# دالة لإضافة معلومات المحاسب
def add_accountant_info(results_df, accountants_df):
    # تنظيف عمود الدائرة
    results_df["normalized_section"] = results_df["الدائرة"].apply(normalize_text)
    accountants_df["normalized_section"] = accountants_df["الدائرة"].apply(normalize_text)

    # إضافة عمود فارغ للنتائج
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

    # إزالة العمود المؤقت
    results_df.drop(columns=["normalized_section"], inplace=True)
    return results_df

# واجهة Streamlit
def main():
    st.set_page_config(page_title="برنامج المطابقة مع المحاسبين", layout="wide")
    st.title("برنامج إضافة المحاسب لنتائج المطابقة")

    uploaded_results = st.file_uploader("اختر ملف نتائج المطابقة (Excel)", type=["xlsx"])
    uploaded_accounts = st.file_uploader("اختر ملف المحاسبين (Excel)", type=["xlsx"])

    if uploaded_results and uploaded_accounts:
        if st.button("إضافة معلومات المحاسب"):
            results_df = pd.read_excel(uploaded_results)
            accounts_df = pd.read_excel(uploaded_accounts)

            with st.spinner("جاري إضافة معلومات المحاسب..."):
                final_df = add_accountant_info(results_df, accounts_df)

            st.success("تمت الإضافة!")
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
