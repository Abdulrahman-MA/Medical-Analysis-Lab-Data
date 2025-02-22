import pandas as pd
import numpy as np
import re

# Load the Excel file
data = pd.read_excel(r"D:\Desktop\Hba1c\Hba1c.xlsx")
df = pd.DataFrame(data)

# Function to remove all characters except those in allowed_chars
def remove_except_in_brackets(text, allowed_chars="0123456789/:."):

    allowed_pattern = f"[^{re.escape(allowed_chars)}]"
    return re.sub(allowed_pattern, "", text)

# Function to extract numbers from text
def extract_numbers(text):
    if isinstance(text, (int, float)):  # Already numeric
        return text
    if pd.isna(text) or str(text).strip().lower() == "nan":  # Handle NaN values
        return np.nan

    text = str(text).strip()

    # Clean the text by removing all characters except allowed ones (digits, /, :)
    text = remove_except_in_brackets(text)

    # Handle fractions (e.g., '1/2' or '1:2')
    if '/' in text or ':' in text:
        parts = re.split(r'[/|:]', text)  # Split by / or :
        try:
            return float(parts[0]) / float(parts[1])
        except ValueError:
            print(f"Failed to convert '{text}' to float")
            return np.nan  # Return NaN if conversion fails

    # If text contains only digits and decimal points, convert it to float
    try:
        return float(text) if text else np.nan
    except ValueError:
        print(f"Failed to convert '{text}' to float")
        return np.nan  # Return NaN if conversion fails

# Apply the function to the last column
df.iloc[:, -1] = df.iloc[:, -1].apply(extract_numbers)

# Save cleaned data
df.to_excel(r"D:\Desktop\Hba1c\Hba1c After Cleaning.xlsx", index=False)

print("Cleaning completed successfully!")
