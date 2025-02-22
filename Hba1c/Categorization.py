import pandas as pd

# Load the Excel file
data = pd.read_excel(r"D:\Desktop\Hba1c\Hba1c After Cleaning.xlsx")
df = pd.DataFrame(data)

# Convert the second-to-last column to numeric, coercing errors to NaN
df.iloc[:, -1] = pd.to_numeric(df.iloc[:, -1], errors='coerce')

# Insert the new category column at the second-to-last position
df.insert(df.shape[1] , df.columns[-1] + '_Cat',
          df.iloc[:, -1].apply(lambda x: 'Normal' if pd.notna(x) and x < 6.5 else 'Abnormal'))

df.to_excel(r"D:\Desktop\Hba1c\Hba1c After cleaning and categorization.xlsx", index=False)
