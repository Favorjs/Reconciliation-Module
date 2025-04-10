import pandas as pd
from difflib import SequenceMatcher
import streamlit as st
from io import BytesIO

try:
    import openpyxl
except ImportError:
    st.error("""
        Critical Error: The 'openpyxl' package is required but not installed.
        
        For local development, run:
        `pip install openpyxl`
        
        For Streamlit Cloud, add 'openpyxl' to requirements.txt
        """)
    st.stop()

# Increase pandas styler max elements
pd.set_option("styler.render.max_elements", 2**31-1)



# --- Page Configuration ---
st.set_page_config(
    page_title="Ultra-Strict Sheet Matcher",
    page_icon="🔍",
    layout="centered"
)

# --- Initialize Session State ---
if 'matching_levels' not in st.session_state:
    st.session_state.matching_levels = [{'col1': None, 'col2': None, 'threshold': 50}]

# --- Matching Functions ---
def clean_name(name):
    return str(name).strip().lower()

def ultra_strict_match(name1, name2):
    """Ultra-strict matching requiring 2+ words with 50%+ similarity"""
    words1 = clean_name(name1).split()
    words2 = clean_name(name2).split()
    
    if len(words1) < 2 or len(words2) < 2:
        return 0
    
    strong_matches = 0
    total_similarity = 0
    
    for i in range(min(3, len(words1))):
        for j in range(min(3, len(words2))):
            if abs(i-j) > 1:
                continue
                
            similarity = SequenceMatcher(None, words1[i], words2[j]).ratio()
            if similarity >= 0.50:
                strong_matches += 1
                total_similarity += similarity
                if strong_matches >= 2:
                    return (total_similarity / strong_matches) * 100
    return 0

def values_match(val1, val2, threshold):
    if pd.isna(val1) or pd.isna(val2):
        return False
    return SequenceMatcher(None, str(val1), str(val2)).ratio() >= (threshold/100)

# --- Core Matching Logic ---
def ultra_strict_matching(df1, df2, match_cols):
    # Rename columns to avoid conflicts
    df1 = df1.add_suffix('_Sheet1')
    df2 = df2.add_suffix('_Sheet2')
    
    matched_cols_sheet1 = [rule['col1'] + '_Sheet1' for rule in match_cols]
    matched_cols_sheet2 = [rule['col2'] + '_Sheet2' for rule in match_cols]
    
    other_cols_sheet1 = [col for col in df1.columns if col not in matched_cols_sheet1]
    other_cols_sheet2 = [col for col in df2.columns if col not in matched_cols_sheet2]
    
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
                
            all_match = True
            match_scores = []
            
            for rule in match_cols:
                col1 = rule['col1'] + '_Sheet1'
                col2 = rule['col2'] + '_Sheet2'
                val1 = rec1.get(col1, None)
                val2 = rec2.get(col2, None)
                
                if not values_match(val1, val2, rule['threshold']):
                    all_match = False
                    break
                
                match_scores.append(SequenceMatcher(None, str(val1), str(val2)).ratio())
            
            if all_match and match_scores:
                current_score = (sum(match_scores)/len(match_scores)) * 100
                if current_score > best_score:
                    best_score = current_score
                    best_match = rec1
                    best_idx1 = idx1
        
        if best_score >= 50:
            combined = {
                **{k: v for k, v in best_match.items() if k in other_cols_sheet1},
                **{k: v for k, v in best_match.items() if k in matched_cols_sheet1},
                **{k: v for k, v in rec2.items() if k in matched_cols_sheet2},
                **{k: v for k, v in rec2.items() if k in other_cols_sheet2},
                'Match_Score': best_score,
                'Match_Status': 'Verified' if best_score >= 90 else 'Confirmed',
                'Match_Type': 'Ultra-Strict Match'
            }
            results.append(combined)
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
                    
                all_match = True
                match_scores = []
                
                for rule in match_cols:
                    col1 = rule['col1'] + '_Sheet1'
                    col2 = rule['col2'] + '_Sheet2'
                    val1 = rec1.get(col1, None)
                    val2 = rec2.get(col2, None)
                    
                    similarity = SequenceMatcher(None, str(val1), str(val2)).ratio() * 100
                    if similarity < rule['threshold']:
                        all_match = False
                        break
                    
                    match_scores.append(similarity)
                
                if all_match and match_scores:
                    current_score = sum(match_scores)/len(match_scores)
                    if 50 <= current_score < 85 and current_score > best_score:
                        best_score = current_score
                        best_match = rec1
                        best_idx1 = idx1
            
            if best_match:
                combined = {
                    **{k: v for k, v in best_match.items() if k in other_cols_sheet1},
                    **{k: v for k, v in best_match.items() if k in matched_cols_sheet1},
                    **{k: v for k, v in rec2.items() if k in matched_cols_sheet2},
                    **{k: v for k, v in rec2.items() if k in other_cols_sheet2},
                    'Match_Score': best_score,
                    'Match_Status': 'Review Recommended',
                    'Match_Type': 'Strict Match'
                }
                results.append(combined)
                matched_indices1.add(best_idx1)
                matched_indices2.add(idx2)

    # Third pass: possible matches
    for idx2, rec2 in enumerate(df2.to_dict('records')):
        if idx2 not in matched_indices2:
            combined = {
                **{k: None for k in other_cols_sheet1},
                **{k: None for k in matched_cols_sheet1},
                **{k: v for k, v in rec2.items()},
                'Match_Score': 0,
                'Match_Status': 'Manual Review Needed',
                'Match_Type': 'Possible Match'
            }
            results.append(combined)

    # Fourth pass: non-matches
    for idx1, rec1 in enumerate(df1.to_dict('records')):
        if idx1 not in matched_indices1:
            combined = {
                **{k: v for k, v in rec1.items()},
                **{k: None for k in matched_cols_sheet2},
                **{k: None for k in other_cols_sheet2},
                'Match_Score': 0,
                'Match_Status': 'No Match Found',
                'Match_Type': 'No Match'
            }
            results.append(combined)

    ordered_columns = (
        other_cols_sheet1 + 
        matched_cols_sheet1 + 
        matched_cols_sheet2 + 
        other_cols_sheet2 + 
        ['Match_Score', 'Match_Status', 'Match_Type']
    )
    
    return pd.DataFrame(results)[ordered_columns].sort_values(
        by=['Match_Score', 'Match_Type'], 
        ascending=[False, True]
    )





# --- UI Components ---
def matching_level_ui(level_idx, sheet1_cols, sheet2_cols):
    cols = st.columns([3, 3, 2, 1])
    
    with cols[0]:
        col1 = st.selectbox(
            f"Sheet1 Column (Level {level_idx+1})",
            sheet1_cols,
            key=f"col1_{level_idx}"
        )
    with cols[1]:
        col2 = st.selectbox(
            f"Sheet2 Column (Level {level_idx+1})",
            sheet2_cols,
            key=f"col2_{level_idx}"
        )
    with cols[2]:
        threshold = st.slider(
            f"Similarity % (Level {level_idx+1})",
            50, 100, 50,
            key=f"threshold_{level_idx}"
        )
    with cols[3]:
        if st.button("❌", key=f"remove_{level_idx}"):
            if len(st.session_state.matching_levels) > 1:
                st.session_state.matching_levels.pop(level_idx)
                st.rerun()
    
    return {'col1': col1, 'col2': col2, 'threshold': threshold}

# --- Streamlit UI ---
st.title("🔍 XModule: Matching Toolkit")
st.markdown("""
**Match records across sheets with:**
- **2+ words at 50%+ similarity**  
- **Multi-level matching**  
- **Upload you data in a file with two sheets** 
- **Download and verify data for duplicates or remove duplicates before matching**  
""")

uploaded_file = st.file_uploader("Upload Excel file with two sheets:", type=["xlsx", "xls"])

if uploaded_file:
    try:
        sheet_names = pd.ExcelFile(uploaded_file).sheet_names
        if len(sheet_names) < 2:
            st.error("The uploaded file must contain at least two sheets.")
            st.stop()
            
        df_sheet1 = pd.read_excel(uploaded_file, sheet_name=0)
        df_sheet2 = pd.read_excel(uploaded_file, sheet_name=1)
        
        # Handle duplicate column names
        df_sheet1 = df_sheet1.add_suffix('_Sheet1')
        df_sheet2 = df_sheet2.add_suffix('_Sheet2')
        
        st.subheader("📋 Original Sheets Preview")
        col1, col2 = st.columns(2)
        with col1:
            st.write(f"📄 {sheet_names[0]}", df_sheet1.head())
        with col2:
            st.write(f"📄 {sheet_names[1]}", df_sheet2.head())

        if not st.session_state.matching_levels[0]['col1']:
            st.session_state.matching_levels[0] = {
                'col1': df_sheet1.columns[0].replace('_Sheet1', ''),
                'col2': df_sheet2.columns[0].replace('_Sheet2', ''),
                'threshold': 50
            }
        
        st.subheader("🔧 Configure Matching Levels")
        if st.button("➕ Add Matching Level"):
            new_col1 = df_sheet1.columns[
                min(len(st.session_state.matching_levels), len(df_sheet1.columns)-1)
            ].replace('_Sheet1', '')
            new_col2 = df_sheet2.columns[
                min(len(st.session_state.matching_levels), len(df_sheet2.columns)-1)
            ].replace('_Sheet2', '')
            st.session_state.matching_levels.append({
                'col1': new_col1,
                'col2': new_col2,
                'threshold': 50
            })
            st.rerun()
        
        matching_rules = []
        for i in range(len(st.session_state.matching_levels)):
            matching_rules.append(
                matching_level_ui(i, 
                                [col.replace('_Sheet1', '') for col in df_sheet1.columns],
                                [col.replace('_Sheet2', '') for col in df_sheet2.columns])
            )

        if st.button("🚀 Run Super Smart Matching", type="primary"):
            with st.spinner("🔍 Hunting for matches..."):
                result = ultra_strict_matching(
                    df_sheet1.rename(columns=lambda x: x.replace('_Sheet1', '')),
                    df_sheet2.rename(columns=lambda x: x.replace('_Sheet2', '')),
                    matching_rules
                )
                
                def color_status(val):
                    colors = {
                        'Verified': '#a5d6a7',
                        'Confirmed': '#c8e6c9',
                        'Review Recommended': '#fff9c4',
                        'Manual Review Needed': '#ffcc80',
                        'No Match Found': '#eeeeee'
                    }
                    return f'background-color: {colors.get(val, "#ffffff")}'
                
                st.subheader("🎯 Matching Results")
                
                # Get column groups
                sheet1_cols = [col for col in result.columns if '_Sheet1' in col]
                sheet2_cols = [col for col in result.columns if '_Sheet2' in col]
                matched_cols = [rule['col1'] + '_Sheet1' for rule in matching_rules] + \
                             [rule['col2'] + '_Sheet2' for rule in matching_rules]
                
                tab1, tab2 = st.tabs(["Full View", "Comparison View"])
                
                with tab1:
                    st.dataframe(
                        result.style.applymap(color_status, subset=['Match_Status']),
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
                
                with tab2:
                    cols = st.columns([2,3,2])
                    with cols[0]:
                        st.write("### Sheet1 Data")
                        st.dataframe(result[[c for c in sheet1_cols if c not in matched_cols]])
                    with cols[1]:
                        st.write("### Matched Columns")
                        st.dataframe(result[matched_cols])
                    with cols[2]:
                        st.write("### Sheet2 Data")
                        st.dataframe(result[[c for c in sheet2_cols if c not in matched_cols]])
                
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    result.to_excel(writer, index=False, sheet_name="Matches")
                    df_sheet1.to_excel(writer, index=False, sheet_name=f"{sheet_names[0]}_original")
                    df_sheet2.to_excel(writer, index=False, sheet_name=f"{sheet_names[1]}_original")
                
                st.download_button(
                    label="📥 Download Full Report",
                    data=output.getvalue(),
                    file_name="smart_matches.xlsx",
                    mime="application/vnd.ms-excel",
                    type="primary"
                )
                
                st.subheader("📊 Match Summary")
                cols = st.columns(4)
                cols[0].metric("🔒 Ultra-Strict", len(result[result['Match_Type'] == 'Ultra-Strict Match']))
                cols[1].metric("✅ Verified", len(result[result['Match_Status'] == 'Verified']))
                cols[2].metric("🔄 Needs Review", len(result[result['Match_Status'] == 'Manual Review Needed']))
                cols[3].metric("❌ No Match", len(result[result['Match_Type'] == 'No Match']))

    except Exception as e:
        st.error(f"🚨 Error: {str(e)}")
with st.sidebar:
    st.markdown("## 🎯 Super Smart Matcher")
    
    # Fun dropdown explanation
    with st.expander("🤔 How This Works"):
        st.markdown("""
        ### ✨ The Matching Wizardry
        
        This cures the irregular name matchup to an extent
        
        1. **🕵️‍♂️ Name Detective**  
           - Breaks names into words  
           - Hunts for similar words in similar positions  
           - Requires **2+ words with 50%+ similarity**
        
        2. **🔍 Multi-Level Matching**  
           - Match on multiple columns (add as many as you want!)  
           - Set different similarity thresholds for each  
        
        3. **🎯 Smart Scoring**  
           - **90%+** = ✅ Verified (Perfect match!)  
           - **85-89%** = 🟢 Confirmed (Pretty sure!)  
           - **50-84%** = 🟡 Review Recommended (Looks familiar...)  
           - **<50%** = 🟠 Manual Review (Hmm, maybe?)  
           - **0%** = ⬜ No Match (Total strangers!)  
        
        4. **📊 Results That Make Sense**  
           - Side-by-side comparison  
           - Confidence scores  
           - Color-coded status  
        
        ### 💡 Pro Tip:
        Start with 50% similarity and adjust up if you get too many matches!
        """)
    
    st.markdown("---")
    st.markdown("### 🛠️ Matching Rules")
    st.markdown("""
    - Minimum **2 matching words**  
    - Each word must have **≥50% similarity**  
    - Words in **similar positions** score higher  
    - **Multiple columns** can be matched  
    - Each column can have **custom thresholds**  
    """)

# --- Footer ---
st.markdown("---")
st.caption("🔍 50%+ similarity matching | 📊 Multi-tier verification | 🎯 Results that make sense")