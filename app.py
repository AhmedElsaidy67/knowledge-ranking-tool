import streamlit as st
import pandas as pd
import io

# 1. إعدادات الصفحة
st.set_page_config(page_title="Knowledge Prioritization Tool", layout="wide")

# صف علوي لتوزيع العنوان وزر التمبلت
header_col, btn_col = st.columns([4, 1])

with header_col:
    st.title("📊 Knowledge Priority Ranking Tool")
    st.markdown("Upload the Excel file and adjust the points to get the best ranking for the review.")

# دالة مساعدة لتحويل الملف لـ Excel (XLSX) للتحميل
def to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
    return output.getvalue()

# --- 2. القائمة الجانبية (Sidebar) ---
st.sidebar.header("⚙️ Configure Points")

with st.sidebar.expander("Views Points", expanded=True):
    p_v_1k = st.number_input("1K views or more", value=2.0)
    p_v_100 = st.number_input("100 or more", value=1.0)
    p_v_50 = st.number_input("50 or more", value=0.50)

with st.sidebar.expander("Scores Points"):
    p_nssd = st.number_input("NSSD (1,2,3)", value=2.0)
    p_fcr = st.number_input("FCR (If 'No')", value=2.0)  # تم توضيح أنها في حالة No
    p_feedback = st.number_input("Feedback (If 'Yes')", value=1.0)
    p_top3 = st.number_input("Search Top 3", value=1.0)
    p_top10 = st.number_input("Search Top 10", value=0.5)
    p_qa = st.number_input("QA/RCA Issue (If 'Yes')", value=1.0)

with st.sidebar.expander("➕ Add Custom Field"):
    custom_col_name = st.text_input("Column Name in Excel", placeholder="e.g. Critical_Error")
    custom_col_points = st.number_input("Points for 'Yes'", value=0.0)

# --- 3. زر التمبلت (موضع جديد في أعلى اليمين) ---
template_cols = ['Knowledge ID', 'Views', 'NSSD', 'FCR', 'Feedback"Yes or No"', 'Search Accuracy', 'QA_RCA_Issues"Yes or No"']
if custom_col_name:
    template_cols.append(custom_col_name)

template_df = pd.DataFrame(columns=template_cols)
with btn_col:
    st.write("") # لإزاحة الزر قليلاً للأسفل ليتساوى مع العنوان
    st.download_button(
        label="📥 Template",
        data=to_excel(template_df),
        file_name="Knowledge_Template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        help="Download an empty template with correct headers"
    )

# --- 4. رفع الملف ---
st.subheader("📤 Upload & Process")
uploaded_file = st.file_uploader("Upload your filled Excel file", type=["xlsx"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)
        df.columns = [c.strip() for c in df.columns]

        def calculate_priority(row):
            score = 0.0 # ضمان التعامل مع الأرقام العشرية
            
            # ① Views
            try:
                v = float(row.get('Views', 0))
                if v >= 1000: score += p_v_1k
                elif v >= 100: score += p_v_100
                elif v >= 50: score += p_v_50
            except: pass

            # ② NSSD (Points if 1, 2, or 3)
            if str(row.get('NSSD')).strip() in ['1', '2', '3', 1, 2, 3]:
                score += p_nssd

            # ③ FCR (Important: Points if 'No')
            fcr_val = str(row.get('FCR', '')).strip().lower()
            if fcr_val == "no":
                score += p_fcr

            # ④ Feedback (Points if 'Yes')
            fb_val = str(row.get('Feedback"Yes or No"', '')).strip().lower()
            if fb_val == "yes":
                score += p_feedback

            # ⑤ Search Accuracy
            acc = str(row.get('Search Accuracy', '')).strip().lower()
            if "top 3" in acc: score += p_top3
            elif "top 10" in acc: score += p_top10

            # ⑥ QA_RCA_Issues (Points if 'Yes')
            qa_val = str(row.get('QA_RCA_Issues"Yes or No"', '')).strip().lower()
            if qa_val == "yes":
                score += p_qa

            # ⑦ Custom Field (Points if 'Yes')
            if custom_col_name in df.columns:
                c_val = str(row.get(custom_col_name, '')).strip().lower()
                if c_val == "yes":
                    score += custom_col_points

            return score

        if st.button("🚀 Run Ranking Analysis"):
            df['Final_Score'] = df.apply(calculate_priority, axis=1)
            df_sorted = df.sort_values(by='Final_Score', ascending=False).reset_index(drop=True)
            
            def get_rank_label(i):
                if i < 50: return "High (Top 50)"
                elif i < 100: return "Medium (Top 100)"
                elif i < 200: return "Normal (Top 200)"
                else: return "Low (Over 200)"
            
            df_sorted['Category'] = [get_rank_label(i) for i in range(len(df_sorted))]

            st.balloons()
            t1, t2, t3, t4 = st.tabs(["🔴 Top 50", "🟠 Top 100", "🟡 Top 200", "📄 All Data"])
            
            with t1: st.dataframe(df_sorted[df_sorted['Category'] == "High (Top 50)"])
            with t2: st.dataframe(df_sorted[df_sorted['Category'] == "Medium (Top 100)"])
            with t3: st.dataframe(df_sorted[df_sorted['Category'] == "Normal (Top 200)"])
            with t4: st.dataframe(df_sorted)

            st.download_button(
                label="📥 Download Final Ranked Report (XLSX)",
                data=to_excel(df_sorted),
                file_name="Final_Ranking_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Error: {e}")