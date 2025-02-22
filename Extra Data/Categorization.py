import pandas as pd

# Load the Excel file
data = pd.read_excel(r"D:\Desktop\Extra Data\Extra Data After Cleaning.xlsx")
df = pd.DataFrame(data)

# Drop the last column
df.drop(df.columns[-1], axis=1, inplace=True)

# Convert the second-to-last column to numeric, coercing errors to NaN
df.iloc[:, -1] = pd.to_numeric(df.iloc[:, -1], errors='coerce')

# Insert the new category column at the second-to-last position
df.insert(df.shape[1], df.columns[-1] + '_Cat', None)

for i in range(df.shape[0]):
        if df.iloc[i, -2] < 1/80:
                df.iloc[i, -1] = 'Normal'
        elif df.iloc[i, -2] >= 1/80:
                df.iloc[i, -1] = 'Abnormal'
        else:
                df.iloc[i, -1] = None

# Convert the second-to-last column to numeric, coercing errors to NaN
df.iloc[:, -3] = pd.to_numeric(df.iloc[:, -3], errors='coerce')

# Insert the new category column at the second-to-last position
df.insert(df.shape[1], df.columns[-3] + '_Cat', None)

for i in range(df.shape[0]):
        if df.iloc[i, -4] <= 20:
                df.iloc[i, -3] = 'Normal'
        elif df.iloc[i, -4] > 20:
                df.iloc[i, -3] = 'Abnormal'
        else:
                df.iloc[i, -3] = None

# Convert the second-to-last column to numeric, coercing errors to NaN
df.iloc[:, -5] = pd.to_numeric(df.iloc[:, -5], errors='coerce')

# Insert the new category column at the second-to-last position
df.insert(df.shape[1], df.columns[-5] + '_Cat', None)

for i in range(df.shape[0]):
        if df.iloc[i, -6] >= 60 :
                df.iloc[i, -5] = 'Normal'
        elif df.iloc[i, -6] < 60 :
                df.iloc[i, -5] = 'Abnormal'
        else:
                df.iloc[i, -5] = None
                
df.to_excel(r"D:\Desktop\Extra Data\Extra Data After Categorization.xlsx", index=False)
