import pandas as pd

def transform_rxn(data):
    result = {}
    for index, row in data.iterrows():
        human_gem = row.iloc[0]  # First column: HUMAN-GEM
        hepato_id = row.iloc[1]  # Second column: rxnHepatoNET1ID
        recon_id = row.iloc[2]   # Third column: rxnRecon3DID
        hmr_id = row.iloc[3]     # Fourth column: rxnHMR2ID
        
        if pd.notna(recon_id) and recon_id != "":
            result[human_gem] = recon_id
        elif pd.notna(hmr_id) and hmr_id != "":
            result[human_gem] = hmr_id
        elif pd.notna(hepato_id) and hepato_id != "":
            result[human_gem] = hepato_id
        else:
            result[human_gem] = None
    return result

# Read the Excel file
file_path = r"C:\Users\omer\TCGADATA\reactionnames.xlsx"
df = pd.read_excel(file_path)

# Transform the data
transformed = transform_rxn(df)

# Print the results
for human_gem, new_id in transformed.items():
    print(f"{human_gem} -> {new_id}")