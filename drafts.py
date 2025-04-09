import pandas as pd
from difflib import SequenceMatcher
import streamlit as st
from io import BytesIO

# --- Page Configuration ---
st.set_page_config(
    page_title="Ultra-Strict Data Matcher Pro",
    page_icon="🔒",
    layout="centered"
)

# --- Initialize Session State ---
if 'matching_levels' not in st.session_state:
    st.session_state.matching_levels = [{'col1': None, 'col2': None, 'threshold': 85}]

# --- Matching Functions ---
def clean_name(name):
    return str(name).strip().lower()

def ultra_strict_match(name1, name2):
    """Ultra-strict matching requiring 2+ words with 85%+ similarity"""
    words1 = clean_name(name1).split()
    words2 = clean_name(name2).split()
    
    if len(words1) < 2 or len(words2) < 2:
        return 0  # Not enough words for ultra-strict matching
    
    strong_matches = 0
    total_similarity = 0
    
    # Check first 3 words maximum
    for i in range(min(3, len(words1))):
        for j in range(min(3, len(words2))):
            if abs(i-j) > 1:  # Only allow reasonable position differences
                continue
                
            similarity = SequenceMatcher(None, words1[i], words2[j]).ratio()
            if similarity >= 0.85:  # 85% similarity threshold
                strong_matches += 1
                total_similarity += similarity
                if strong_matches >= 2:
                    return (total_similarity / strong_matches) * 100  # Average match percentage
    return 0

# --- Core Matching Logic ---
def ultra_strict_matching(df1, df2, name_col1, name_col2):
    units_col1 = next((col for col in df1.columns if 'unit' in col.lower()), None)
    units_col2 = next((col for col in df2.columns if 'unit' in col.lower()), None)
    
    results = []
    matched_indices1 = set()
    matched_indices2 = set()
    
    # First pass: ultra-strict matches
    for idx2, rec2 in enumerate(df2.to_dict('records')):
        best_score = 0
        best_match = None
        best_idx1 = None
        
        for idx1, rec1 in enumerate(df1.to_dict('records')):
            if idx1 in matched_indices1:
                continue
                
            score = ultra_strict_match(rec1[name_col1], rec2[name_col2])
            if score > best_score:
                best_score = score
                best_match = rec1
                best_idx1 = idx1
        
        if best_score >= 85:  # Only consider if meets ultra-strict threshold
            results.append({
                'Type': 'Ultra-Strict Match',
                'Match_Score': best_score,
                'Account_Number': rec2.get('Account Number', ''),
                'Name_Sheet1': best_match[name_col1],
                'Name_Sheet2': rec2[name_col2],
                'Units_Sheet1': best_match.get(units_col1, '') if units_col1 else '',
                'Units_Sheet2': rec2.get(units_col2, '') if units_col2 else '',
                'Match_Status': 'Verified' if best_score >= 90 else 'Confirmed'
            })
            matched_indices1.add(best_idx1)
            matched_indices2.add(idx2)
    
    # Second pass: strict-but-not-ultra matches
    for idx2, rec2 in enumerate(df2.to_dict('records')):
        if idx2 not in matched_indices2:
            best_score = 0
            best_match = None
            
            for idx1, rec1 in enumerate(df1.to_dict('records')):
                if idx1 in matched_indices1:
                    continue
                    
                score = SequenceMatcher(None, 
                                      clean_name(rec1[name_col1]), 
                                      clean_name(rec2[name_col2])).ratio() * 100
                if 60 <= score < 85 and score > best_score:  # Strict but not ultra-strict
                    best_score = score
                    best_match = rec1
                    best_idx1 = idx1
            
            if best_match:
                results.append({
                    'Type': 'Strict Match',
                    'Match_Score': best_score,
                    'Account_Number': rec2.get('Account Number', ''),
                    'Name_Sheet1': best_match[name_col1],
                    'Name_Sheet2': rec2[name_col2],
                    'Units_Sheet1': best_match.get(units_col1, '') if units_col1 else '',
                    'Units_Sheet2': rec2.get(units_col2, '') if units_col2 else '',
                    'Match_Status': 'Review Recommended'
                })
                matched_indices1.add(best_idx1)
                matched_indices2.add(idx2)
    
    # Third pass: possible matches (below 60%)
    for idx2, rec2 in enumerate(df2.to_dict('records')):
        if idx2 not in matched_indices2:
            results.append({
                'Type': 'Possible Match',
                'Match_Score': 0,
                'Account_Number': rec2.get('Account Number', ''),
                'Name_Sheet1': '',
                'Name_Sheet2': rec2[name_col2],
                'Units_Sheet1': '',
                'Units_Sheet2': rec2.get(units_col2, '') if units_col2 else '',
                'Match_Status': 'Manual Review Needed'
            })
    
    # Fourth pass: complete non-matches
    for idx1, rec1 in enumerate(df1.to_dict('records')):
        if idx1 not in matched_indices1:
            results.append({
                'Type': 'No Match',
                'Match_Score': 0,
                'Account_Number': '',
                'Name_Sheet1': rec1[name_col1],
                'Name_Sheet2': '',
                'Units_Sheet1': rec1.get(units_col1, '') if units_col1 else '',
                'Units_Sheet2': '',
                'Match_Status': 'No Match Found'
            })
    
    return pd.DataFrame(results).sort_values(
        by=['Match_Score', 'Type'], 
        ascending=[False, True]
    )

# --- Streamlit UI ---
st.title("🔒 Ultra-Strict Data Matcher Pro")
st.markdown("""
**Merge datasets with ultra-strict matching:**
- Requires **2+ words with 85%+ similarity**  
- Position-aware matching  
- Multi-tier verification system  
""")

# --- File Upload ---
file_cols = st.columns(2)
with file_cols[0]:
    file1 = st.file_uploader("Primary Dataset", type=["xlsx", "xls"])
with file_cols[1]:
    file2 = st.file_uploader("Secondary Dataset", type=["xlsx", "xls"])

if file1 and file2:
    try:
        df1 = pd.read_excel(file1)
        df2 = pd.read_excel(file2)
        
        # --- Column Selection ---
        st.subheader("Select Matching Columns")
        cols = st.columns(2)
        with cols[0]:
            name_col1 = st.selectbox(
                "Name column (Primary Dataset)",
                df1.columns,
                index=next((i for i, col in enumerate(df1.columns) if 'name' in col.lower()), 0)
            )
        with cols[1]:
            name_col2 = st.selectbox(
                "Name column (Secondary Dataset)",
                df2.columns,
                index=next((i for i, col in enumerate(df2.columns) if 'name' in col.lower()), 0)
            )
        
        # --- Run Matching ---
        if st.button("🚀 Run Ultra-Strict Matching", type="primary"):
            with st.spinner("Applying matching rules..."):
                result = ultra_strict_matching(df1, df2, name_col1, name_col2)
                
                # --- Color Coding ---
                def color_cells(val):
                    if val == 'Verified':
                        return 'background-color: #a5d6a7'  # Strong green
                    elif val == 'Confirmed':
                        return 'background-color: #c8e6c9'  # Light green
                    elif val == 'Review Recommended':
                        return 'background-color: #fff9c4'  # Yellow
                    elif val == 'Manual Review Needed':
                        return 'background-color: #ffcc80'  # Orange
                    else:
                        return 'background-color: #eeeeee'   # Gray
                
                # --- Display Results ---
                st.dataframe(
                    result.style.applymap(color_cells, subset=['Match_Status']),
                    height=700,
                    column_config={
                        "Match_Score": st.column_config.ProgressColumn(
                            "Confidence",
                            format="%.0f%%",
                            min_value=0,
                            max_value=100,
                        )
                    }
                )
                
                # --- Download ---
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    result.to_excel(writer, index=False)
                st.download_button(
                    label="📥 Download Matches",
                    data=output.getvalue(),
                    file_name="ultra_strict_matches.xlsx",
                    mime="application/vnd.ms-excel",
                    type="primary"
                )
                
                # --- Summary Stats ---
                st.subheader("Match Summary")
                stat_cols = st.columns(4)
                ultra_strict = len(result[result['Type'] == 'Ultra-Strict Match'])
                verified = len(result[result['Match_Status'] == 'Verified'])
                review_needed = len(result[result['Match_Status'] == 'Manual Review Needed'])
                no_match = len(result[result['Type'] == 'No Match'])
                
                stat_cols[0].metric("Ultra-Strict", ultra_strict)
                stat_cols[1].metric("Verified", verified)
                stat_cols[2].metric("Needs Review", review_needed)
                stat_cols[3].metric("No Match", no_match)
    
    except Exception as e:
        st.error(f"Error: {str(e)}")

# --- Sidebar ---
with st.sidebar:
    st.markdown("### Matching Rules")
    st.markdown("""
    **Ultra-Strict Requirements:**
    - Minimum **2 matching words**
    - Each word must have **≥85% similarity**
    - Words must be in **similar positions**
    
    **Verification Tiers:**
    - ✅ **Verified (90%+)** - Perfect matches
    - 🟢 **Confirmed (85-89%)** - Strong matches
    - 🟡 **Review Recommended (60-84%)** - Good matches
    - 🟠 **Manual Review Needed (<60%)** - Possible matches
    - ⬜ **No Match** - No similarity found
    """)

# --- Footer ---
st.markdown("---")
st.caption("🔒 85%+ similarity matching | 📊 Multi-tier verification | ✅ Professional-grade results")