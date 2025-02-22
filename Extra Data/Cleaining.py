import pandas as pd
import numpy as np
import re

def load_excel(file_path):
    """Load an Excel file into a pandas DataFrame."""
    return pd.read_excel(file_path)

def remove_except_in_brackets(text, allowed_chars="0123456789/:."):
    """Remove all characters except numbers and special symbols like / and :."""
    return re.sub(f"[^{re.escape(allowed_chars)}]", "", str(text))

def extract_numbers(text):
    """Extract and convert numbers from text, handling invalid cases."""
    
    if isinstance(text, (int, float)):  # Already numeric
        return text
    if pd.isna(text) or str(text).strip().lower() in ["nan", ".", ".."]:  # Handle NaN values & dots
        return np.nan  

    text = remove_except_in_brackets(str(text).strip())
    text = text.replace("..", ".")  # Fix double dots

    if '/' in text or ':' in text:
        parts = re.split(r'[/|:]', text)
        if len(parts) >= 2:
            try:
                return float(parts[0]) / float(parts[1])
            except (ValueError, ZeroDivisionError):
                print(f"Failed to convert '{text}' to float")
                return np.nan
        else:
            print(f"Skipping invalid fraction: '{text}'")
            return np.nan

    try:
        return float(text) if text else np.nan
    except ValueError:
        print(f"Failed to convert '{text}' to float")
        return np.nan

def clean_columns(df, column_indices):
    
    actual_columns = []

    # Convert negative indices to column names
    for index in column_indices:
        if isinstance(index, range):  # Handle range() objects
            actual_columns.extend([df.columns[i] for i in index])
        elif isinstance(index, int):  # Handle single integer indices
            actual_columns.append(df.columns[index])

    # Apply number extraction to the selected columns
    for column in actual_columns:
        df[column] = df[column].apply(extract_numbers)
    
    return df

def save_excel(df, file_path):
    """Save the cleaned DataFrame to an Excel file."""
    df.to_excel(file_path, index=False)

def process_file(input_path, output_path, column_indices, name):
    """Load, clean, and save an Excel file."""
    df = load_excel(input_path)
    df = clean_columns(df, column_indices)
    save_excel(df, output_path)
    print(f"Cleaning of file '{name}' completed successfully!")

# List of files to process with corresponding columns to clean
files_to_process = [
    (r"D:\Desktop\HLA B27 extradata final .xlsx", r"D:\Desktop\extradata Clean.xlsx", [-4, -2, -3], "Extradata"),
    (r"D:\Desktop\Hba1c.xlsx", r"D:\Desktop\Hba1c Clean.xlsx", [-1], "Hba1c"),
    (r"D:\Desktop\ANA no IFA.xlsx", r"D:\Desktop\ANA Clean.xlsx", [-1], "ANA"),
    (r"D:\Desktop\HLA B27 final data.xlsx", r"D:\Desktop\HLA-B27 Clean.xlsx", [range(10, 25), range(26, 39)], "HLA-B27 Typing by PCR"),
]

for input_path, output_path, column_indices, name in files_to_process:
    process_file(input_path, output_path, column_indices, name)
