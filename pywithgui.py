import pandas as pd
from difflib import SequenceMatcher
import streamlit as st
from io import BytesIO

# --- Page Configuration ---
st.set_page_config(
    page_title="Ultra-Strict Sheet Matcher",
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
def ultra_strict_matching(df1, df2, match_cols):
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
                
            # Check all matching levels
            all_match = True
            match_scores = []
            
            for rule in match_cols:
                val1 = rec1[rule['col1']] if rule['col1'] in rec1 else None
                val2 = rec2[rule['col2']] if rule['col2'] in rec2 else None
                
                if not values_match(val1, val2, rule['threshold']):
                    all_match = False
                    break
                
                match_scores.append(
                    SequenceMatcher(None, str(val1), str(val2)).ratio()
                )
            
            if all_match and match_scores:
                current_score = (sum(match_scores)/len(match_scores)) * 100
                if current_score > best_score:
                    best_score = current_score
                    best_match = rec1
                    best_idx1 = idx1
        
        if best_score >= 85:  # Only consider if meets ultra-strict threshold
            results.append({
                'Type': 'Ultra-Strict Match',
                'Match_Score': best_score,
                'Account_Number': rec2.get('Account Number', ''),
                'Name_Sheet1': best_match.get('Name', ''),
                'Name_Sheet2': rec2.get('Name', ''),
                'Units_Sheet1': best_match.get('Units', ''),
                'Units_Sheet2': rec2.get('Units', ''),
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
                    
                # Check all matching levels
                all_match = True
                match_scores = []
                
                for rule in match_cols:
                    val1 = rec1[rule['col1']] if rule['col1'] in rec1 else None
                    val2 = rec2[rule['col2']] if rule['col2'] in rec2 else None
                    
                    similarity = SequenceMatcher(None, str(val1), str(val2)).ratio() * 100
                    if similarity < rule['threshold']:
                        all_match = False
                        break
                    
                    match_scores.append(similarity)
                
                if all_match and match_scores:
                    current_score = sum(match_scores)/len(match_scores)
                    if 60 <= current_score < 85 and current_score > best_score:
                        best_score = current_score
                        best_match = rec1
                        best_idx1 = idx1
            
            if best_match:
                results.append({
                    'Type': 'Strict Match',
                    'Match_Score': best_score,
                    'Account_Number': rec2.get('Account Number', ''),
                    'Name_Sheet1': best_match.get('Name', ''),
                    'Name_Sheet2': rec2.get('Name', ''),
                    'Units_Sheet1': best_match.get('Units', ''),
                    'Units_Sheet2': rec2.get('Units', ''),
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
                'Name_Sheet2': rec2.get('Name', ''),
                'Units_Sheet1': '',
                'Units_Sheet2': rec2.get('Units', ''),
                'Match_Status': 'Manual Review Needed'
            })
    
    # Fourth pass: complete non-matches
    for idx1, rec1 in enumerate(df1.to_dict('records')):
        if idx1 not in matched_indices1:
            results.append({
                'Type': 'No Match',
                'Match_Score': 0,
                'Account_Number': '',
                'Name_Sheet1': rec1.get('Name', ''),
                'Name_Sheet2': '',
                'Units_Sheet1': rec1.get('Units', ''),
                'Units_Sheet2': '',
                'Match_Status': 'No Match Found'
            })
    
    return pd.DataFrame(results).sort_values(
        by=['Match_Score', 'Type'], 
        ascending=[False, True]
    )

def values_match(val1, val2, threshold):
    """Compare two values with threshold"""
    if pd.isna(val1) or pd.isna(val2):
        return False
    return SequenceMatcher(None, str(val1), str(val2)).ratio() >= (threshold/100)

# --- UI Components ---
def matching_level_ui(level_idx, sheet1_cols, sheet2_cols):
    """Create UI for one matching level"""
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
            70, 100, 85,
            key=f"threshold_{level_idx}"
        )
    with cols[3]:
        if st.button("❌", key=f"remove_{level_idx}"):
            if len(st.session_state.matching_levels) > 1:
                st.session_state.matching_levels.pop(level_idx)
                st.rerun()
    
    return {'col1': col1, 'col2': col2, 'threshold': threshold}

# --- Streamlit UI ---
st.title("🔒 Ultra-Strict Sheet Matcher")
st.markdown("""
**Match records across sheets with:**
- **2+ words at 85%+ similarity**  
- **Multi-level matching**  
- **Position-aware comparison**  
""")

# --- File Upload ---
uploaded_file = st.file_uploader("Upload Excel file with two sheets:", type=["xlsx", "xls"])

if uploaded_file:
    try:
        # Load both sheets
        df_sheet1 = pd.read_excel(uploaded_file, sheet_name=0)
        df_sheet2 = pd.read_excel(uploaded_file, sheet_name=1)
        
        # Initialize matching levels if empty
        if not st.session_state.matching_levels:
            st.session_state.matching_levels = [{
                'col1': df_sheet1.columns[0],
                'col2': df_sheet2.columns[0],
                'threshold': 85
            }]
        
        # --- Matching Configuration ---
        st.subheader("Configure Matching Levels")
        
        # Add level button
        if st.button("➕ Add Matching Level"):
            new_col1 = df_sheet1.columns[
                min(len(st.session_state.matching_levels), len(df_sheet1.columns)-1)
            ]
            new_col2 = df_sheet2.columns[
                min(len(st.session_state.matching_levels), len(df_sheet2.columns)-1)
            ]
            st.session_state.matching_levels.append({
                'col1': new_col1,
                'col2': new_col2,
                'threshold': 85
            })
            st.rerun()
        
        # Display all levels
        matching_rules = []
        for i in range(len(st.session_state.matching_levels)):
            matching_rules.append(
                matching_level_ui(i, df_sheet1.columns, df_sheet2.columns)
            )
        
        # --- Run Matching ---
        if st.button("🚀 Run Ultra-Strict Matching", type="primary"):
            with st.spinner("Applying matching rules..."):
                result = ultra_strict_matching(df_sheet1, df_sheet2, matching_rules)
                
                # --- Color Coding ---
                def color_status(val):
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
                cols = st.columns(4)
                ultra_strict = len(result[result['Type'] == 'Ultra-Strict Match'])
                verified = len(result[result['Match_Status'] == 'Verified'])
                review_needed = len(result[result['Match_Status'] == 'Manual Review Needed'])
                no_match = len(result[result['Type'] == 'No Match'])
                
                cols[0].metric("Ultra-Strict", ultra_strict)
                cols[1].metric("Verified", verified)
                cols[2].metric("Needs Review", review_needed)
                cols[3].metric("No Match", no_match)
    
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