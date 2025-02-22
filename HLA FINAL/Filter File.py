import pandas as pd
import os

# Define folder paths
Maximum = "D:/Desktop/Maximum"  # Folder with original files
Minimum = "D:/Desktop/Minimum"
OutputMin = "D:/Desktop/Min Filtered Data"  # Folder for filtered files
OutputMax = "D:/Desktop/Max Filtered Data"

# Create output folder if it doesn't exist
os.makedirs(OutputMin, exist_ok=True)
os.makedirs(OutputMax, exist_ok=True)

# List of files to skip
fourfiles = [
    'Platelet Count_Platelet Count Cat_Platelet Count Cat (High)_Platelet Count Cat (Low).xlsx',
]

# Loop through all Excel files in the input folder
for file_name in os.listdir(Minimum):
    if file_name.endswith(".xlsx") and file_name not in fourfiles:  # Process only Excel files not in `fourfiles`
        file_path = os.path.join(Minimum, file_name)
        
        # Load the Excel file
        df = pd.read_excel(file_path)

        # Use column index instead of names
        sample_id_col = df.columns[0]  # Assuming 'Sample Research ID' is at index 0
        reg_date_col = df.columns[6]  # Assuming 'Reg Date' is at index 6
        alt_col = df.columns[-2]  # Assuming 'Alanine Aminotransferase (ALT)' is 2nd last column
        alt_cat_col = df.columns[-1]  # Assuming 'ALT Cat' is last column

        # Convert 'Reg Date' to datetime
        df[reg_date_col] = pd.to_datetime(df[reg_date_col], errors='coerce')

        # Define filtering function per patient
        def filter_patient_data(group):
            group = group.dropna(subset=[alt_col])  # Remove rows where ALT is null
            
            if group.empty:
                return None  # Skip this patient if all values are null

            if "Abnormal" in group[alt_cat_col].values:
                return group.sort_values(by=[alt_col], ascending=True).iloc[0]  # Lowest ALT value if abnormal exists
            else:
                return group.sort_values(by=[reg_date_col], ascending=False).iloc[0]  # Most recent valid Reg Date otherwise

        # Apply filtering logic per patient
        df_unique = df.groupby(sample_id_col, group_keys=False).apply(filter_patient_data)

        # Generate new file name (first three letters + "_filtered.xlsx")
        new_file_name = file_name[:3] + "_filtered.xlsx"
        OutputMin_file_path = os.path.join(OutputMin, new_file_name)

        # Save the processed file
        df_unique.to_excel(OutputMin_file_path, index=False)
        print(f"Filtered file saved as: {new_file_name}")

# Loop through all Excel files in the input folder
for file_name in os.listdir(Maximum):
    if file_name.endswith(".xlsx") and file_name not in fourfiles:  # Process only Excel files not in `fourfiles`
        file_path = os.path.join(Maximum, file_name)
        
        # Load the Excel file
        df = pd.read_excel(file_path)

        # Use column index instead of names
        sample_id_col = df.columns[0]  # Assuming 'Sample Research ID' is at index 0
        reg_date_col = df.columns[6]  # Assuming 'Reg Date' is at index 6
        alt_col = df.columns[-2]  # Assuming 'Alanine Aminotransferase (ALT)' is 2nd last column
        alt_cat_col = df.columns[-1]  # Assuming 'ALT Cat' is last column

        # Convert 'Reg Date' to datetime
        df[reg_date_col] = pd.to_datetime(df[reg_date_col], errors='coerce')

        # Define filtering function per patient
        def filter_patient_data(group):
            group = group.dropna(subset=[alt_col])  # Remove rows where ALT is null
            
            if group.empty:
                return None  # Skip this patient if all values are null

            if "Abnormal" in group[alt_cat_col].values:
                return group.sort_values(by=[alt_col], ascending=True).iloc[0]  # Lowest ALT value if abnormal exists
            else:
                return group.sort_values(by=[reg_date_col], ascending=False).iloc[0]  # Most recent valid Reg Date otherwise

        # Apply filtering logic per patient
        df_unique = df.groupby(sample_id_col, group_keys=False).apply(filter_patient_data)

        # Generate new file name (first three letters + "_filtered.xlsx")
        new_file_name = file_name[:3] + "_filtered.xlsx"
        OutputMax_file_path = os.path.join(OutputMax, new_file_name)

        # Save the processed file
        df_unique.to_excel(OutputMax_file_path, index=False)
        print(f"Filtered file saved as: {new_file_name}")

for file_name in os.listdir(Minimum):
    if file_name.endswith(".xlsx") and file_name in fourfiles:  # Process only Excel files in fourfiles
        file_path = os.path.join(Minimum, file_name)
        
        # Load the Excel file
        df = pd.read_excel(file_path)

        # Use column index instead of names
        sample_id_col = df.columns[0]  # 'Sample Research ID' (index 0)
        reg_date_col = df.columns[6]  # 'Reg Date' (index 6)
        alt_col = df.columns[-4]  # ALT Test Column (4th last)
        alt_cat_col = df.columns[-3]  # ALT Category Column (3rd Last)

        # Convert 'Reg Date' to datetime
        df[reg_date_col] = pd.to_datetime(df[reg_date_col], errors='coerce')

        # Convert ALT values to numeric (for sorting)
        df[alt_col] = pd.to_numeric(df[alt_col], errors='coerce')

        # Define filtering function per patient
        def filter_patient_data(group):
            group = group.dropna(subset=[alt_col])  # Remove rows where ALT is null
            
            if group.empty:
                return None  # Skip this patient if all values are null

            # If all ALT Category values are "High", skip this patient
            if all(group[alt_cat_col] == "High"):
                return None

            # If "Low" exists, return the row with the lowest ALT value
            if "Low" in group[alt_cat_col].values:
                return group[group[alt_cat_col] == "Low"].sort_values(by=[alt_col], ascending=True).iloc[0]

            # If "Normal" exists, return the latest valid "Normal" value
            if "Normal" in group[alt_cat_col].values:
                return group[group[alt_cat_col] == "Normal"].sort_values(by=[reg_date_col], ascending=False).iloc[0]

            # If neither "Low" nor "Normal" exists, return None
            return None

        # Apply filtering logic per patient
        df_unique = df.groupby(sample_id_col, group_keys=False).apply(filter_patient_data)

        # Rename columns in df_unique
        df_unique.rename(columns={alt_col: alt_col + " #Low", alt_cat_col: alt_cat_col + "(3L)"}, inplace=True)

        # Drop the last column (possibly unnecessary)
        df_unique.drop(df.columns[-2], axis=1, inplace=True)

        # Define the renamed category column
        alt_cat_col_new = alt_cat_col + "(3L)"

        # Ensure the column exists before filtering
        if alt_cat_col_new in df_unique.columns:
            df_unique = df_unique[df_unique[alt_cat_col_new] != "High"]

        # Generate new file name (first three letters + "_filtered.xlsx")
        new_file_name = file_name[:3] + "_filtered.xlsx"
        OutputMin_file_path = os.path.join(OutputMin, new_file_name)

        # Save the processed file
        df_unique.to_excel(OutputMin_file_path, index=False)
        print(f"Filtered file saved as: {new_file_name}")

for file_name in os.listdir(Maximum):
    if file_name.endswith(".xlsx") and file_name in fourfiles:  # Process only Excel files in fourfiles
        file_path = os.path.join(Maximum, file_name)
        
        # Load the Excel file
        df = pd.read_excel(file_path)

        # Use column index instead of names
        sample_id_col = df.columns[0]  # 'Sample Research ID' (index 0)
        reg_date_col = df.columns[6]  # 'Reg Date' (index 6)
        alt_col = df.columns[-4]  # ALT Test Column (4th last)
        alt_cat_col = df.columns[-3]  # ALT Category Column (3rd Last)

        # Convert 'Reg Date' to datetime
        df[reg_date_col] = pd.to_datetime(df[reg_date_col], errors='coerce')

        # Convert ALT values to numeric (for sorting)
        df[alt_col] = pd.to_numeric(df[alt_col], errors='coerce')

        # Define filtering function per patient
        def filter_patient_data(group):
            group = group.dropna(subset=[alt_col])  # Remove rows where ALT is null
            
            if group.empty:
                return None  # Skip this patient if all values are null

            # If all ALT Category values are "Low", skip this patient
            if all(group[alt_cat_col] == "Low"):
                return None

            # If "High" exists, return the row with the highest ALT value
            if "High" in group[alt_cat_col].values:
                return group[group[alt_cat_col] == "High"].sort_values(by=[alt_col], ascending=False).iloc[0]

            # If "Normal" exists, return the latest valid "Normal" value
            if "Normal" in group[alt_cat_col].values:
                return group[group[alt_cat_col] == "Normal"].sort_values(by=[reg_date_col], ascending=False).iloc[0]

            # If neither "High" nor "Normal" exists, return None
            return None

        # Apply filtering logic per patient
        df_unique = df.groupby(sample_id_col, group_keys=False).apply(filter_patient_data)

        # Rename columns in df_unique
        df_unique.rename(columns={alt_col: alt_col + " #HIGH", alt_cat_col: alt_cat_col + "(3H)"}, inplace=True)

        # Drop the last column (possibly unnecessary)
        df_unique.drop(df.columns[-1], axis=1, inplace=True)

        # Define the renamed category column
        alt_cat_col_new = alt_cat_col + "(3H)"

        # Ensure the column exists before filtering
        if alt_cat_col_new in df_unique.columns:
            df_unique = df_unique[df_unique[alt_cat_col_new] != "Low"]

        # Generate new file name (first three letters + "_filtered.xlsx")
        new_file_name = file_name[:3] + "_filtered.xlsx"
        OutputMax_file_path = os.path.join(OutputMax, new_file_name)

        # Save the processed file
        df_unique.to_excel(OutputMax_file_path, index=False)
        print(f"Filtered file saved as: {new_file_name}")
