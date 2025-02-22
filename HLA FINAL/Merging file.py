import pandas as pd
import os

# Define file paths
main_file_path = r"D:\Desktop\HLA FINAL\HLA-B27 Typing by PCR.xlsx"
output_file = r"D:\Desktop\merged_output.xlsx"
ana_file = r'D:\Desktop\HLA FINAL\ANA no IFA.xlsx'
hba1c_file = r"D:\Desktop\Hba1c\Hba1c Filtered.xlsx"
Extra_file = r"D:\Desktop\Extra Data\Extra Data After Categorization.xlsx"

# Directories to process
folders = [r"D:\Desktop\Max Filtered Data", r"D:\Desktop\Min Filtered Data"]

# List of files where we need the last 4 columns instead of 2
four_files = {'Tot_filtered', 'Thy_filtered', 'T3__filtered', 'MCV_filtered', 
              'T4__filtered', 'Pla_filtered', 'Fer_filtered'}

# Load the main Excel file
main_df = pd.read_excel(main_file_path, engine='openpyxl')

# Identify Patient ID column dynamically (assuming it's the second column)
patient_id_main = main_df.columns[1]

# Normalize Patient IDs for consistency
main_df[patient_id_main] = main_df[patient_id_main].astype(str).str.strip().str.lower()

# Process and merge each file
for folder in folders:
    for file_name in os.listdir(folder):
        if file_name.endswith(".xlsx"):
            file_path = os.path.join(folder, file_name)
            try:
                # Load Excel file
                df = pd.read_excel(file_path, engine='openpyxl')

                # Identify Patient ID column (assuming it's the first column)
                patient_id_folder = df.columns[0]

                # Normalize Patient IDs
                df[patient_id_folder] = df[patient_id_folder].astype(str).str.strip().str.lower()

                # Determine columns to merge
                num_cols = 3 if any(name in file_name for name in four_files) else 2
                columns_to_merge = [patient_id_folder] + list(df.columns[-num_cols:])

                # Merge only necessary columns
                main_df = main_df.merge(df[columns_to_merge], left_on=patient_id_main, 
                                        right_on=patient_id_folder, how="left")

                print(f"Merged {file_name}")

            except Exception as e:
                print(f"Error processing file {file_name}: {e}")

# Function to merge a secondary file (ANA, Hba1c) with main DataFrame
import pandas as pd

def merge_secondary_file(df_main, file_path, file_label):
    try:
        df_secondary = pd.read_excel(file_path, engine='openpyxl')
        
        # Identify Patient ID column (assuming it's the first column)
        patient_id_col = df_secondary.columns[0]

        if file_label == "Extra Data":
            # Last 6 columns including Patient ID
            last_columns = list(df_secondary.columns[-6:])

        else:
            # Only last column including Patient ID
            last_columns = [df_secondary.columns[-1]]

        # Select only the Patient ID and required columns
        df_secondary_filtered = df_secondary[[patient_id_col] + last_columns]

        # Merge based on Patient ID
        df_main = df_main.merge(df_secondary_filtered, on=patient_id_col, how="left")
        print(f"Merged: {file_label}")

        return df_main
    
    except Exception as e:
        print(f"Error merging {file_label}: {e}")
        return df_main  # Return unchanged DataFrame in case of error

# Merge ANA and Hba1c files
main_df = merge_secondary_file(main_df, ana_file, "ANA")
main_df = merge_secondary_file(main_df, hba1c_file, "Hba1c")
main_df = merge_secondary_file(main_df, Extra_file, "Extra Data")

# Save the final merged DataFrame
main_df.to_excel(output_file, index=False)
print(f"Merged file saved at: {output_file}")
