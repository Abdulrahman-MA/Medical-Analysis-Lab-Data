import pandas as pd

# Define folder paths
OutputMax = r"D:\Desktop\Hba1c\Hba1c Filtered.xlsx"

file_path = r"D:\Desktop\Hba1c\Hba1c After cleaning and categorization.xlsx"

# Load the Excel file
df = pd.read_excel(file_path)

# Use column index instead of names
sample_id_col = df.columns[0]  # Assuming 'Sample Research ID' is at index 0
reg_date_col = df.columns[4]  # Assuming 'Reg Date' is at index 6
val_col = df.columns[-2]  # Assuming 'Alanine Aminotransferase (ALT)' is 2nd last column
cat_col = df.columns[-1]  # Assuming 'ALT Cat' is last column

# Convert 'Reg Date' to datetime
df[reg_date_col] = pd.to_datetime(df[reg_date_col], errors='coerce')

# Define filtering function per patient
def filter_patient_data(group):
    group = group.dropna(subset=[val_col])  # Remove rows where ALT is null
    
    if group.empty:
        return None  # Skip this patient if all values are null

    if "Abnormal" in group[cat_col].values:
        return group.sort_values(by=[val_col], ascending=False).iloc[0]  # Lowest ALT value if abnormal exists
    else:
        return group.sort_values(by=[reg_date_col], ascending=False).iloc[0]  # Most recent valid Reg Date otherwise

# Apply filtering logic per patient
df_unique = df.groupby(sample_id_col, group_keys=False).apply(filter_patient_data)

# Ensure 'Research ID' column is of string type
df_unique['Research ID'] = df_unique['Research ID'].astype(str)

# Extract the Patient ID
df_unique['Patient ID'] = df_unique['Research ID'].apply(lambda x: x.split('-')[1] if '-' in x else '')

# Move 'Patient ID' to the first column
cols = df_unique.columns.tolist()
cols = cols[-1:] + cols[:-1]
df_unique = df_unique[cols]

# Save the processed file
df_unique.to_excel(OutputMax, index=False)
print(f"Filtered file saved as: {OutputMax}")
