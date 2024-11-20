import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment

# Set error tolerance range
tolerance = 0.0030

# Load data
normal_df = pd.read_excel('data/normal.xlsx')
cancer_df = pd.read_excel('data/cancer.xlsx')

# Extract the m/z data from the first column
normal_mz = normal_df.iloc[:, 0]
cancer_mz = cancer_df.iloc[:, 0]

# Get column names
normal_cols = normal_df.columns.tolist()
cancer_cols = cancer_df.columns.tolist()

# Initialize result lists
matched_rows = []
remaining_normal = []
remaining_cancer = []

# Find matched items
for i, mz_normal in normal_mz.items():
    found_match = False
    for j, mz_cancer in cancer_mz.items():
        if abs(mz_normal - mz_cancer) <= tolerance:
            # Combine the data of normal and cancer
            combined_row = pd.concat([normal_df.iloc[i], cancer_df.iloc[j]], ignore_index=True)
            combined_row = combined_row.rename(lambda x: f"C{x}")  # Temporarily name columns as C0, C1, etc.
            combined_row['m/z'] = mz_normal  # Add m/z column for grouping
            matched_rows.append(combined_row)
            found_match = True
            break
    if not found_match:
        # If no match, add to remaining_normal
        remaining_normal.append(normal_df.iloc[i])

# Find remaining cancer items
for j, mz_cancer in cancer_mz.items():
    if not any(abs(mz_cancer - mz_normal) <= tolerance for mz_normal in normal_mz):
        remaining_cancer.append(cancer_df.iloc[j])

# Create DataFrame and set appropriate column names
if matched_rows:
    matched_df = pd.DataFrame(matched_rows)
    matched_df.sort_values(by='m/z', inplace=True)  # Sort by m/z 

    # Use the column names from normal and cancer data
    column_names = normal_cols + cancer_cols
    matched_df.columns = column_names + ["m/z"]  # Add "m/z" column for following operations

    # Delete the last "m/z" column
    matched_df.drop(columns=["m/z"], inplace=True)

# Write results to Excel file
with pd.ExcelWriter('data/common.xlsx', engine='openpyxl') as writer:
    if not matched_df.empty:
        matched_df.to_excel(writer, index=False, sheet_name='Common')
    else:
        pd.DataFrame().to_excel(writer, index=False, sheet_name='Common')  # Empty file

# Load saved file to merge cells
wb = load_workbook('data/common.xlsx')
ws = wb['Common']

# Merge cells with the same m/z
for row in range(2, ws.max_row + 1):
    if ws[f'A{row}'].value == ws[f'A{row - 1}'].value:
        ws.merge_cells(f'A{row - 1}:A{row}')
        ws[f'A{row - 1}'].alignment = Alignment(horizontal="center", vertical="center")

wb.save('data/common.xlsx')

# Generate remaining files
with pd.ExcelWriter('data/normal_only.xlsx') as writer:
    if remaining_normal:
        pd.DataFrame(remaining_normal).to_excel(writer, index=False, sheet_name='Normal_Only')
    else:
        pd.DataFrame().to_excel(writer, index=False, sheet_name='Normal_Only')  # Empty file

with pd.ExcelWriter('data/cancer_only.xlsx') as writer:
    if remaining_cancer:
        pd.DataFrame(remaining_cancer).to_excel(writer, index=False, sheet_name='Cancer_Only')
    else:
        pd.DataFrame().to_excel(writer, index=False, sheet_name='Cancer_Only')  # Empty file

print("Comparison file generation complete. Generated files include common.xlsx, normal_only.xlsx, and cancer_only.xlsx.")