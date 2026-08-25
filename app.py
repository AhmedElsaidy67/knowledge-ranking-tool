import streamlit as st
import pandas as pd
import io
import re
from datetime import datetime, timedelta

# 1. Page Configuration
st.set_page_config(page_title="Knowledge Prioritization Tool", layout="wide")

header_col, btn_col = st.columns([4, 1])

with header_col:
    st.title("📊 Knowledge Priority Ranking Tool")
    st.markdown("Upload multiple Excel files, automatically extract embedded Knowledge IDs, aggregate unique problem occurrences for FCR & NSSD, and generate rankings.")

# Helper function to convert DataFrames to downloadable XLSX
def to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
    return output.getvalue()

# --- 2. Sidebar Configuration ---
st.sidebar.header("⚙️ Configure Points & Thresholds")

# Views Thresholds Cards (5 Ranges UI)
with st.sidebar.expander("👁️ Views Thresholds & Points (5 Ranges)", expanded=True):
    st.markdown("---")
    
    # 1. Very High Views Range Card (> 10,000)
    with st.container(border=True):
        st.markdown("#### 🚀 Very High Range (> 10,000)")
        v_vh_thresh = st.number_input("Threshold (Min Views)", value=10000, step=500, key="vvh_t")
        p_v_10k = st.number_input("Points Awarded", value=10.0, key="vvh_p")

    # 2. High Views Range Card (5,000 - 10,000)
    with st.container(border=True):
        st.markdown("#### 🔥 High Range (5,000 - 10,000)")
        v_h_thresh = st.number_input("Threshold (Min Views)", value=5000, step=250, key="vh_t")
        p_v_5k = st.number_input("Points Awarded", value=5.0, key="vh_p")
    
    # 3. Medium Views Range Card (1,000 - 5,000)
    with st.container(border=True):
        st.markdown("#### 📈 Medium Range (1,000 - 5,000)")
        v_m_thresh = st.number_input("Threshold (Min Views)", value=1000, step=100, key="vm_t")
        p_v_1k = st.number_input("Points Awarded", value=3.0, key="vm_p")
    
    # 4. Low-Medium Views Range Card (500 - 1,000)
    with st.container(border=True):
        st.markdown("#### 📊 Low-Medium Range (500 - 1,000)")
        v_lm_thresh = st.number_input("Threshold (Min Views)", value=500, step=50, key="vlm_t")
        p_v_500 = st.number_input("Points Awarded", value=2.0, key="vlm_p")

    # 5. Low Views Range Card (< 500)
    with st.container(border=True):
        st.markdown("#### 📉 Low Range (< 500)")
        p_v_below500 = st.number_input("Points Awarded", value=1.0, key="vl_p")
        
    st.markdown("---")

with st.sidebar.expander("Scores Points"):
    p_nssd = st.number_input("NSSD Point (Per Issue Occurrence)", value=5.0)
    p_fcr = st.number_input("FCR Point (Per 'N'/'No' Occurrence)", value=5.0)
    p_feedback = st.number_input("Feedback (If 'Yes')", value=3.0)
    p_top3 = st.number_input("Search Top 3", value=3.0)
    p_top10 = st.number_input("Search Top 10", value=1.0)
    p_qa = st.number_input("QA Issue (If 'Yes')", value=3.0)
    p_rca = st.number_input("RCA Issue (If 'Yes')", value=3.0)

with st.sidebar.expander("➕ Add Custom Field"):
    custom_col_name = st.text_input("Column Name in Excel", placeholder="e.g. Critical_Error")
    custom_col_points = st.number_input("Points for 'Yes'", value=0.0)

with st.sidebar.expander("🔄 Column Name Mapping (Alternative Names)"):
    st.caption("Provide comma-separated alternative column names if header names differ across sheets.")
    alt_title = st.text_input("Alternative names for 'Knowledge Title'", placeholder="Title, Article_Title")
    alt_type = st.text_input("Alternative names for 'Knowledge Type'", placeholder="Type, Article_Type")
    alt_date = st.text_input("Alternative names for 'Last update date'", placeholder="Update_Date, Modified_Date, Last_Modified")
    alt_views = st.text_input("Alternative names for 'Views'", placeholder="Page_Views, View_Count")
    alt_nssd = st.text_input("Alternative names for 'NSSD'", placeholder="NSSD_Score, Priority")
    alt_fcr = st.text_input("Alternative names for 'FCR'", placeholder="FCR_Status, First_Call")
    alt_feedback = st.text_input("Alternative names for 'Feedback'", placeholder="User_Feedback, Feedback_Given")
    alt_search = st.text_input("Alternative names for 'Search Accuracy'", placeholder="Search_Rank, Accuracy")
    alt_qa = st.text_input("Alternative names for 'QA Issues'", placeholder="QA_Issue, QA_Flag")
    alt_rca = st.text_input("Alternative names for 'RCA Issues'", placeholder="RCA_Issue, RCA_Flag")

# --- 3. Base Template Columns (شامل Last update date) ---
base_template_cols = [
    'Knowledge ID', 
    'Knowledge Title', 
    'Knowledge Type', 
    'Last update date',
    'Views', 
    'NSSD', 
    'FCR', 
    'Feedback"Yes or No"', 
    'Search Accuracy', 
    'QA Issues"Yes or No"', 
    'RCA Issues"Yes or No"'
]

if custom_col_name:
    base_template_cols.append(custom_col_name)

template_df = pd.DataFrame(columns=base_template_cols)

with btn_col:
    st.write("")
    st.download_button(
        label="📥 Download Template",
        data=to_excel(template_df),
        file_name="Knowledge_Template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        help="Download clean template"
    )

# --- 4. File Upload & Processing Trigger ---
st.subheader("📤 Step 1: Upload Excel Files")
uploaded_files = st.file_uploader("Upload all Excel files", type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    if st.button("⚡ Extract & Merge Data"):
        try:
            all_extracted_rows = []
            id_pattern = re.compile(r'en-us(?:-[a-zA-Z0-9]+)?\d+', re.IGNORECASE)

            total_files = len(uploaded_files)
            progress_bar = st.progress(0)
            status_text = st.empty()

            for file_idx, uploaded_file in enumerate(uploaded_files):
                xls = pd.ExcelFile(uploaded_file)
                num_sheets = len(xls.sheet_names)

                for sheet_idx, sheet_name in enumerate(xls.sheet_names):
                    current_progress = (file_idx + (sheet_idx + 1) / num_sheets) / total_files
                    progress_bar.progress(min(current_progress, 1.0))
                    status_text.text(f"Processing File {file_idx + 1}/{total_files}: '{uploaded_file.name}' | Sheet: '{sheet_name}'...")

                    df_temp = pd.read_excel(xls, sheet_name=sheet_name)
                    df_temp.columns = [str(c).strip() for c in df_temp.columns]

                    # استخراج الـ IDs الفريدة في كل صف لحماية الحساب لو الـ ID اتكرر ف نفس الصف بـ أعمدة مختلفة
                    for idx, row in df_temp.iterrows():
                        row_ids = set()
                        for col in df_temp.columns:
                            cell_value = str(row[col])
                            matches = id_pattern.findall(cell_value)
                            for match in matches:
                                row_ids.add(match.lower())
                        
                        # الصف كله يعامل كـ Entry واحد فقط لكل ID فريد تم إيجاده فيه
                        for found_id in row_ids:
                            row_data = row.to_dict()
                            row_data['Knowledge ID'] = found_id
                            all_extracted_rows.append(row_data)

            progress_bar.progress(1.0)
            status_text.success("🎉 File extraction and scanning completed successfully!")

            if not all_extracted_rows:
                st.error("❌ No matching Knowledge IDs (e.g., en-us, en-us-vol, en-us-map) were found in the uploaded files.")
            else:
                raw_df = pd.DataFrame(all_extracted_rows)

                # Column name unification
                mapping_rules = {
                    'Knowledge Title': [x.strip() for x in alt_title.split(',') if x.strip()],
                    'Knowledge Type': [x.strip() for x in alt_type.split(',') if x.strip()],
                    'Last update date': [x.strip() for x in alt_date.split(',') if x.strip()],
                    'Views': [x.strip() for x in alt_views.split(',') if x.strip()],
                    'NSSD': [x.strip() for x in alt_nssd.split(',') if x.strip()],
                    'FCR': [x.strip() for x in alt_fcr.split(',') if x.strip()],
                    'Feedback"Yes or No"': [x.strip() for x in alt_feedback.split(',') if x.strip()],
                    'Search Accuracy': [x.strip() for x in alt_search.split(',') if x.strip()],
                    'QA Issues"Yes or No"': [x.strip() for x in alt_qa.split(',') if x.strip()],
                    'RCA Issues"Yes or No"': [x.strip() for x in alt_rca.split(',') if x.strip()]
                }

                for target_col, alt_cols in mapping_rules.items():
                    if target_col not in raw_df.columns:
                        raw_df[target_col] = None
                    for alt in alt_cols:
                        if alt in raw_df.columns:
                            raw_df[target_col] = raw_df[target_col].fillna(raw_df[alt])

                # تجميع البيانات وحساب تكرارات المشاكل فقط بناءً على الحالات الفريدة (الصفوف)
                def aggregate_problems(group):
                    res = group.iloc[0].to_dict()

                    # 1. عد تكرارات مشاكل FCR (الصفوف الفريدة التي احتوت على N أو No)
                    fcr_vals = [str(x).strip().lower() for x in group['FCR'].dropna()]
                    fcr_issues_count = sum(1 for x in fcr_vals if x in ['n', 'no'])
                    res['FCR_Issue_Count'] = fcr_issues_count
                    res['FCR'] = 'No' if fcr_issues_count > 0 else (fcr_vals[0] if fcr_vals else None)

                    # 2. عد تكرارات مشاكل NSSD والحفاظ على القيمة النصية الأصلية
                    bad_nssd_list = ['1', '2', '3', 'very unsatisfied', 'unsatisfactory', 'normal']
                    
                    raw_nssd_list = [x for x in group['NSSD'].dropna() if str(x).strip()]
                    
                    nssd_issues_count = 0
                    first_bad_val = None
                    
                    for val in raw_nssd_list:
                        if str(val).strip().lower() in bad_nssd_list:
                            nssd_issues_count += 1
                            if first_bad_val is None:
                                first_bad_val = val # الاحتفاظ بالنص الأصلي لـ NSSD بدلاً من كلمة Issue

                    res['NSSD_Issue_Count'] = nssd_issues_count
                    # وضع النص الأصلي للمشكلة عند وجودها
                    res['NSSD'] = first_bad_val if first_bad_val is not None else (raw_nssd_list[0] if raw_nssd_list else None)

                    # 3. الفحص عن أية إجابات Yes في الخانات الأخرى
                    for bool_col in ['Feedback"Yes or No"', 'QA Issues"Yes or No"', 'RCA Issues"Yes or No"']:
                        vals = [str(x).strip().lower() for x in group[bool_col].dropna()]
                        if any(x == 'yes' for x in vals):
                            res[bool_col] = 'Yes'

                    # أقصى عدد Views
                    try:
                        res['Views'] = group['Views'].astype(float).max()
                    except: pass

                    # استخراج أحدث تاريخ متاح في حالة وجود أكثر من سطر
                    date_vals = group['Last update date'].dropna()
                    if not date_vals.empty:
                        res['Last update date'] = date_vals.iloc[0]

                    return pd.Series(res)

                clean_df = raw_df.groupby('Knowledge ID', as_index=False).apply(aggregate_problems).reset_index(drop=True)

                for col in base_template_cols:
                    if col not in clean_df.columns:
                        clean_df[col] = None

                st.session_state['clean_df'] = clean_df

        except Exception as e:
            st.error(f"An error occurred during file extraction: {e}")

# --- 5. Data Ranking & Output Sections ---
if 'clean_df' in st.session_state:
    clean_df = st.session_state['clean_df']
    st.success(f"✅ Extracted and merged ({len(clean_df)}) unique Knowledge IDs successfully!")

    with st.expander("📋 Download Clean Merged Data (Pre-Ranking)"):
        st.info("This file contains the consolidated raw data with unified columns prior to calculating points.")
        st.download_button(
            label="📥 Download Clean Merged Sheet (.xlsx)",
            data=to_excel(clean_df[base_template_cols]),
            file_name="Clean_Merged_Data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    # Calculation logic (5 Views Ranges + شرط الاستبعاد)
    def calculate_priority(row):
        # فحص شرط الاستبعاد: لو تم تحديثه خلال شهر (30 يوم) ونوعه Services -> تصفير السكور
        k_type = str(row.get('Knowledge Type', '')).strip().lower()
        date_val = row.get('Last update date')
        
        if 'services' in k_type and pd.notna(date_val):
            try:
                parsed_date = pd.to_datetime(date_val)
                # فحص ما إذا كان التاريخ خلال الـ 30 يوماً الماضية
                if parsed_date >= (datetime.now() - timedelta(days=30)):
                    return 0.0 # تصفير النقاط تماماً
            except:
                pass

        score = 0.0
        
        # Views مع النطاقات الخمسة الجديدة
        try:
            v = float(row.get('Views', 0))
            if v >= v_vh_thresh: 
                score += p_v_10k            # > 10,000
            elif v >= v_h_thresh: 
                score += p_v_5k             # 5,000 - 10,000
            elif v >= v_m_thresh: 
                score += p_v_1k             # 1,000 - 5,000
            elif v >= v_lm_thresh: 
                score += p_v_500            # 500 - 1,000
            else: 
                score += p_v_below500       # < 500
        except: pass

        # NSSD: ضرب النقاط في عدد الصفوف (الكيسات) الفريدة التي طرأت فيها المشكلة فقط
        nssd_issues = float(row.get('NSSD_Issue_Count', 0))
        score += (nssd_issues * p_nssd)

        # FCR: ضرب النقاط في عدد الصفوف (الكيسات) الفريدة التي طرأت فيها المشكلة فقط
        fcr_issues = float(row.get('FCR_Issue_Count', 0))
        score += (fcr_issues * p_fcr)

        # Feedback
        fb_val = str(row.get('Feedback"Yes or No"', '')).strip().lower()
        if fb_val == "yes":
            score += p_feedback

        # Search Accuracy
        acc = str(row.get('Search Accuracy', '')).strip().lower()
        if "top 3" in acc: score += p_top3
        elif "top 10" in acc: score += p_top10

        # QA Issues
        qa_val = str(row.get('QA Issues"Yes or No"', '')).strip().lower()
        if qa_val == "yes":
            score += p_qa

        # RCA Issues
        rca_val = str(row.get('RCA Issues"Yes or No"', '')).strip().lower()
        if rca_val == "yes":
            score += p_rca

        # Custom Field
        if custom_col_name in clean_df.columns:
            c_val = str(row.get(custom_col_name, '')).strip().lower()
            if c_val == "yes":
                score += custom_col_points

        return score

    st.subheader("🚀 Step 2: Run Prioritization")
    if st.button("🚀 Calculate & Rank Priority"):
        clean_df['Final_Score'] = clean_df.apply(calculate_priority, axis=1)
        
        # Sort descending
        df_sorted = clean_df.sort_values(by='Final_Score', ascending=False).reset_index(drop=True)

        def get_rank_label(row, i):
            if row['Final_Score'] == 0.0:
                return "Low (Score 0)"
            if i < 50: return "High (Top 50)"
            elif i < 100: return "Medium (Top 100)"
            elif i < 200: return "Normal (Top 200)"
            else: return "Low (Over 200)"
        
        df_sorted['Category'] = [get_rank_label(df_sorted.iloc[i], i) for i in range(len(df_sorted))]

        # Reorder columns to show essential identifier, title, type & date first
        first_cols = ['Knowledge ID', 'Knowledge Title', 'Knowledge Type', 'Last update date', 'Final_Score', 'Category']
        export_columns = first_cols + [c for c in base_template_cols if c not in first_cols]
        df_final_export = df_sorted[export_columns]

        st.balloons()
        t1, t2, t3, t4 = st.tabs(["🔴 Top 50", "🟠 Top 100", "🟡 Top 200", "📄 All Data"])
        
        with t1: st.dataframe(df_final_export[df_final_export['Category'] == "High (Top 50)"])
        with t2: st.dataframe(df_final_export[df_final_export['Category'] == "Medium (Top 100)"])
        with t3: st.dataframe(df_final_export[df_final_export['Category'] == "Normal (Top 200)"])
        with t4: st.dataframe(df_final_export)

        st.download_button(
            label="📥 Download Final Clean Ranked Report (.xlsx)",
            data=to_excel(df_final_export),
            file_name="Final_Clean_Ranked_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
