import pandas as pd
import difflib

# Load both sheets (adjust sheet names if needed)
df_units = pd.read_excel('C:/Users/fadebowale/Documents/Chapelhill recon.xlsx', sheet_name='Sheet1')  # Columns: [Name, Units]
df_accounts = pd.read_excel('C:/Users/fadebowale/Documents/Chapelhill recon.xlsx', sheet_name='Sheet2')  # Columns: [Account Number, Name]

# Clean names for better matching
def clean_name(name):
    return str(name).strip().lower()

df_units['clean_name'] = df_units['Name'].apply(clean_name)
df_accounts['clean_name'] = df_accounts['Name'].apply(clean_name)

# Find best matches (adjust cutoff=0.6 for stricter/looser matches)
sheet1_names = df_units['clean_name'].tolist()

def find_best_match(name):
    matches = difflib.get_close_matches(name, sheet1_names, n=1, cutoff=0.6)
    return matches[0] if matches else None

df_accounts['matched_name'] = df_accounts['clean_name'].apply(find_best_match)

# Merge matched accounts with units (left join to keep all accounts)
merged_df = pd.merge(
    df_accounts,
    df_units,
    left_on='matched_name',
    right_on='clean_name',
    how='left',
    suffixes=('_account', '_units')
)

# Get unmatched units (people in Sheet1 not linked to any account)
unmatched_units = df_units[~df_units['clean_name'].isin(merged_df['matched_name'].dropna())]

# Combine matched and unmatched data
final_df = pd.concat([
    merged_df[['Account Number', 'Name_account', 'Units']].rename(columns={
        'Name_account': 'Name',
        'Units': 'Units'
    }),
    pd.DataFrame({
        'Account Number': [None] * len(unmatched_units),
        'Name': unmatched_units['Name'],
        'Units': unmatched_units['Units']
    })
], ignore_index=True)

# Save to Excel
final_df.to_excel('final_output.xlsx', index=False)

print("Done! Check 'final_output.xlsx'")