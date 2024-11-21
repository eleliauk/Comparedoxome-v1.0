import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment

# Set error tolerance range
tolerance = 0.0030

# Load data
normal_df = pd.read_excel('data/normal.xlsx')
cancer_df = pd.read_excel('data/cancer.xlsx')

# Ensure m/z data is of float type
normal_mz = normal_df.iloc[:, 0].astype(float)
cancer_mz = cancer_df.iloc[:, 0].astype(float)

# Get column names
normal_cols = [f"Normal_{col}" for col in normal_df.columns.tolist()]
cancer_cols = [f"Cancer_{col}" for col in cancer_df.columns.tolist()]

# Initialize result lists and matched indices
matched_rows = []
remaining_normal = []
remaining_cancer = []

matched_cancer_indices = set()  # To store matched cancer indices

# Find matched items
for i, mz_normal in normal_mz.items():
    found_match = False
    for j, mz_cancer in cancer_mz.items():
        if abs(mz_normal - mz_cancer) <= tolerance and j not in matched_cancer_indices:
            # Combine the data of normal and cancer
            combined_row = pd.concat([normal_df.iloc[i], cancer_df.iloc[j]], ignore_index=True)
            combined_row['m/z'] = mz_normal  # Add m/z column for grouping
            matched_rows.append(combined_row)
            matched_cancer_indices.add(j)  # Mark cancer row as matched
            found_match = True
            break
    if not found_match:
        # If no match, add to remaining_normal
        remaining_normal.append(normal_df.iloc[i])

# Find remaining cancer items
remaining_cancer = [
    cancer_df.iloc[j] for j in range(len(cancer_mz))
    if j not in matched_cancer_indices
]

# Create DataFrame and set appropriate column names
if matched_rows:
    matched_df = pd.DataFrame(matched_rows)
    matched_df.sort_values(by='m/z', inplace=True)  # Sort by m/z

    # Use the column names from normal and cancer data
    column_names = normal_cols + cancer_cols
    matched_df.columns = column_names + ["m/z"]  # Add "m/z" column for following operations

# Write results to Excel file
with pd.ExcelWriter('data/matched.xlsx', engine='openpyxl') as writer:
    if 'matched_df' in locals() and not matched_df.empty:
        matched_df.to_excel(writer, index=False, sheet_name='Matched')
    else:
        pd.DataFrame(columns=normal_cols + cancer_cols + ["m/z"]).to_excel(writer, index=False, sheet_name='Matched')  # Empty file

# Load saved file to merge cells
wb = load_workbook('data/matched.xlsx')
ws = wb['Matched']

# Merge cells with the same m/z
for row in range(2, ws.max_row + 1):
    if ws[f'A{row}'].value == ws[f'A{row - 1}'].value:
        ws.merge_cells(f'A{row - 1}:A{row}')
        ws[f'A{row - 1}'].alignment = Alignment(horizontal="center", vertical="center")

wb.save('data/matched.xlsx')

# Generate remaining files
with pd.ExcelWriter('data/remaining_normal.xlsx') as writer:
    if remaining_normal:
        pd.DataFrame(remaining_normal).to_excel(writer, index=False, sheet_name='Remaining_Normal', header=normal_cols)
    else:
        pd.DataFrame(columns=normal_cols).to_excel(writer, index=False, sheet_name='Remaining_Normal')

with pd.ExcelWriter('data/remaining_cancer.xlsx') as writer:
    if remaining_cancer:
        pd.DataFrame(remaining_cancer).to_excel(writer, index=False, sheet_name='Remaining_Cancer', header=cancer_cols)
    else:
        pd.DataFrame(columns=cancer_cols).to_excel(writer, index=False, sheet_name='Remaining_Cancer')

print("Comparison file generation complete. Generated files include matched.xlsx, remaining_normal.xlsx, "
      "and remaining_cancer.xlsx.")