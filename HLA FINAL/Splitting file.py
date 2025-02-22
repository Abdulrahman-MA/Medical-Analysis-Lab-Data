import pandas as pd
import time
start_time = time.time()
print(f'Starting the process {start_time} ')

# Load the data
data = pd.read_excel("D:/Desktop/Data after cleaning and categorization.xlsx")  # Adjust path and file format if needed
df = pd.DataFrame(data)

# Convert 'Reg Date' and 'Patient Birth Date' to datetime
df['Reg Date'] = pd.to_datetime(df['Reg Date'], errors='coerce')
df['Patient Birth Date'] = pd.to_datetime(df['Patient Birth Date'], errors='coerce')

# Function to calculate age
def calculate_age(birth_date, visit_date):
    if pd.isnull(birth_date) or pd.isnull(visit_date):
        return None  # Return None if either date is missing
    age = abs(visit_date.year - birth_date.year)
    return age

# Calculate 'Age Years'
df['Age Years'] = df.apply(
    lambda row: calculate_age(row['Patient Birth Date'], row['Reg Date']), axis=1
)

# Extract Patient ID from 'Research ID'
df['Patient ID'] = df['Research ID'].str.split('-').str[-1]

# Filter rows for valid 'HLA-B27 Typing by PCR' values (Positive, Negative, None)
df = df[
    (df['HLA-B27 Typing by PCR'].isin(['Positive', 'Negative'])) |
    (df['HLA-B27 Typing by PCR'].isnull())
]

# Define a function to aggregate test results for each patient
def merge_tests(group):
    # Sort by registration date (most recent first)
    group = group.sort_values('Reg Date', ascending=False)
    
    # Columns to be merged (all except specified ones)
    exclude_columns = ['Research ID', 'Patient ID', 'Reg Date', 'City Name', 'Sample Research ID',
                       'Classification (B2B or B2C)', 'REG_YEAR', 'REG_MONTH', 'Patient Birth Date', 'Age Years']
    test_columns = [col for col in group.columns if col not in exclude_columns]

# Define the columns to check
fourcolumns = {'Ferritin In Serum', 'MCV', 'T3', 'T4', 'Thyroid Stimulating Hormone (TSH)', 'Platelet Count', 'Total Leucocytic Count(WBC)'}
Minimum = {'Hemoglobin','HDL Cholesterol','Apolipoprotein A1','Vitamin D (25 OH-Vit D -Total)2','Albumin in Serum'}
except_column = {'HLA-B27 Typing by PCR','Patient ID'}

# Define fixed columns (first 10 columns)
fixed_columns = df.columns[:10]
''''
# Start iteration from column 11
i = 10
while i < len(df.columns):
    if df.columns[i] in except_column:  # Skip excluded column
        i += 1
        continue
    
    if df.columns[i] in fourcolumns:
        additional_columns = df.columns[i:i+4]  # Take 4 columns
        i += 4  # Move forward by 4

        additional_columns = [col for col in additional_columns if col in df.columns and col not in except_column]

        file_name = f"D:/Desktop/Minimum/{'_'.join(additional_columns)}.xlsx"
        df[list(fixed_columns) + additional_columns].to_excel(file_name, index=False) 

        file_name = f"D:/Desktop/Maximum/{'_'.join(additional_columns)}.xlsx"
        df[list(fixed_columns) + additional_columns].to_excel(file_name, index=False)
    
    
    elif df.columns[i] in Minimum:
        additional_columns = df.columns[i:i+2]  # Take 2 columns
        i += 2  # Move forward by 2

        additional_columns = [col for col in additional_columns if col in df.columns and col not in except_column]

        file_name = f"D:/Desktop/Minimum/{'_'.join(additional_columns)}.xlsx"
        df[list(fixed_columns) + additional_columns].to_excel(file_name, index=False) 

    else:
        additional_columns = df.columns[i:i+2]  # Take 2 columns
        i += 2  # Move forward by 2

        additional_columns = [col for col in additional_columns if col in df.columns and col not in except_column]

        file_name = f"D:/Desktop/Maximum/{'_'.join(additional_columns)}.xlsx"
        df[list(fixed_columns) + additional_columns].to_excel(file_name, index=False) '''

# Selecting specific columns
fixed_columns = df.columns[:10].tolist()  # List the first 10 columns as fixed columns
additional_columns = ['HLA-B27 Typing by PCR']  # List the additional columns

# Check if the columns exist in the DataFrame
for col in additional_columns + ['Patient ID', 'Reg Date']:
    if col not in df.columns:
        raise KeyError(f"Column '{col}' does not exist in DataFrame.")

# Sort the DataFrame by 'Patient ID' and 'Reg Date' in descending order
df_sorted = df.sort_values(by=['Patient ID', 'Reg Date'], ascending=[True, False])

# Drop duplicates to keep the most recent 'HLA-B27 Typing by PCR' value for each 'Patient ID'
df_filtered = df_sorted.drop_duplicates(subset=['Patient ID'], keep='first')

# Filter out rows with null values in the specified additional columns
df_filtered = df_filtered.dropna(subset=additional_columns + ['Patient ID'])

# Relocate the Patient ID column to the first position
columns_order = ['Patient ID'] + fixed_columns + additional_columns

# Ensure no duplicates in columns_order
columns_order = list(dict.fromkeys(columns_order))

# Construct the file name
file_name = f"D:/Desktop/{'_'.join(additional_columns)}.xlsx"

# Export the selected columns to an Excel file
df_filtered[columns_order].to_excel(file_name, index=False)

print("File has been created successfully!")

