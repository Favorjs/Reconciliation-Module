import pandas as pd
from difflib import SequenceMatcher
import streamlit as st
from io import BytesIO

# --- Page Setup ---
st.set_page_config(
    page_title="Excel Reconciliation Toolkit",
    page_icon="🔄",
    layout="centered"
)

# --- Shared Functions ---
def clean_name(name):
    return str(name).strip().lower()

def has_sufficient_match(name1, name2, min_words=2, threshold=0.6):
    words1 = clean_name(name1).split()
    words2 = clean_name(name2).split()
    common = 0
    for w1 in words1:
        for w2 in words2:
            if SequenceMatcher(None, w1, w2).ratio() >= threshold:
                common += 1
                if common >= min_words:
                    return True
    return False

# --- Module 1: Flexible Name Matching ---
def module_flexible_match(df1, df2):
    # Auto-detect name columns
    name_col1 = [col for col in df1.columns if 'name' in col.lower()][0]
    name_col2 = [col for col in df2.columns if 'name' in col.lower()][0]
    
    # Process matching
    matched = []
    for _, row2 in df2.iterrows():
        best_match = None
        best_score = 0
        for _, row1 in df1.iterrows():
            if SequenceMatcher(None, clean_name(row1[name_col1]), clean_name(row2[name_col2])).ratio() >= 0.7:
                score = SequenceMatcher(None, clean_name(row1[name_col1]), clean_name(row2[name_col2])).ratio()
                if score > best_score:
                    best_score = score
                    best_match = row1
        
        if best_match:
            result = {'Name_Matched': row2[name_col2]}
            # Add all columns from both sheets
            for col in df2.columns:
                result[f"Sheet2_{col}"] = row2[col]
            for col in df1.columns:
                result[f"Sheet1_{col}"] = best_match[col]
            matched.append(result)
    
    return pd.DataFrame(matched)

# --- Module 2: CHN-RIN Matching with Units Reconciliation ---
def module_chn_rin_match(df1, df2):
    # Auto-detect CHN/RIN columns
    chn_col = [col for col in df1.columns if 'chn' in col.lower()][0]
    rin_col = [col for col in df2.columns if 'rin' in col.lower()][0]
    
    # Find units columns
    units_col1 = [col for col in df1.columns if 'unit' in col.lower()][0]
    units_col2 = [col for col in df2.columns if 'unit' in col.lower()][0]
    
    # Merge data
    merged = pd.merge(df1, df2, 
                     left_on=chn_col, 
                     right_on=rin_col, 
                     how='outer',
                     suffixes=('_Sheet1', '_Sheet2'))
    
    # Calculate variance
    merged['Units_Variance'] = merged[units_col1+'_Sheet1'] - merged[units_col2+'_Sheet2']
    
    return merged

# --- Module 3: Strict 2-Word Name Matching ---
def module_strict_match(df1, df2):
    name_col1 = [col for col in df1.columns if 'name' in col.lower()][0]
    name_col2 = [col for col in df2.columns if 'name' in col.lower()][0]
    
    matched = []
    unmatched_df1 = df1.copy()
    
    for _, row2 in df2.iterrows():
        best_match = None
        best_score = 0
        for _, row1 in df1.iterrows():
            if has_sufficient_match(row1[name_col1], row2[name_col2]):
                score = SequenceMatcher(None, 
                                      clean_name(row1[name_col1]), 
                                      clean_name(row2[name_col2])).ratio()
                if score > best_score:
                    best_score = score
                    best_match = row1
        
        if best_match:
            result = {f'Name_Matched': row2[name_col2]}
            for col in df2.columns:
                result[f"Sheet2_{col}"] = row2[col]
            for col in df1.columns:
                result[f"Sheet1_{col}"] = best_match[col]
            matched.append(result)
            unmatched_df1 = unmatched_df1[unmatched_df1[name_col1] != best_match[name_col1]]
    
    # Add unmatched
    unmatched = []
    for _, row in unmatched_df1.iterrows():
        record = {f'Name_Matched': ''}
        for col in df1.columns:
            record[f"Sheet1_{col}"] = row[col]
        for col in df2.columns:
            record[f"Sheet2_{col}"] = ''
        unmatched.append(record)
    
    return pd.DataFrame(matched + unmatched)

# --- Main App ---
st.title("🔄 Excel Reconciliation Toolkit")
st.markdown("Match and reconcile data from different Excel sheets")

# Module selection
module = st.selectbox("Select Module", [
    "1. Flexible Name Matching",
    "2. CHN-RIN Units Reconciliation",
    "3. Strict 2-Word Name Matching"
])

uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx", "xls"])

if uploaded_file:
    # Read both sheets
    df1 = pd.read_excel(uploaded_file, sheet_name=0)
    df2 = pd.read_excel(uploaded_file, sheet_name=1)
    
    # Process based on module
    if "Flexible" in module:
        result = module_flexible_match(df1, df2)
    elif "CHN-RIN" in module:
        result = module_chn_rin_match(df1, df2)
    else:
        result = module_strict_match(df1, df2)
    
    # Show results
    st.success(f"Processed {len(result)} records")
    st.dataframe(result.head(10))
    
    # Download
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        result.to_excel(writer, index=False)
    
    st.download_button(
        label="📥 Download Results",
        data=output.getvalue(),
        file_name="reconciliation_results.xlsx",
        mime="application/vnd.ms-excel"
    )

# --- Module Descriptions ---
st.sidebar.markdown("""
### Module Guide
1. **Flexible Name Matching**  
   - Auto-detects name columns  
   - 70% similarity threshold  
   - Preserves all original columns

2. **CHN-RIN Units Reconciliation**  
   - Matches CHN (Sheet1) with RIN (Sheet2)  
   - Compares Units side-by-side  
   - Calculates variance column

3. **Strict 2-Word Matching**  
   - Requires ≥2 similar words  
   - Shows unmatched records  
   - Case-insensitive comparison
""")