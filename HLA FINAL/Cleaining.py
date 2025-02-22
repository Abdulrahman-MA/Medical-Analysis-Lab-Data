import pandas as pd

data = pd.read_excel(r"D:\Desktop\HLA FINAL\HLA B27 final data.xlsx")
df = pd.DataFrame(data)

def extract_numbers(text):
    text = str(text) if not isinstance(text, str) else text
    if '/' in text:
        parts = text.split('/')
        try:
            return float(parts[0]) / float(parts[1])
        except ValueError:
            return text
    elif ':' in text:
        parts = text.split(':')
        try:
            return float(parts[0]) / float(parts[1])
        except ValueError:
            return text
    else:
        allowed_chars = set('0123456789.')
        return ''.join([char for char in text if char in allowed_chars])

columns = ['Estimated Glomerular Filtration Rate(eGFR)', 'Rheumatoid Factor (quantitative)', 'Titre on Hep2 cells']

for column in columns:
    df[column] = df[column].apply(extract_numbers)

df.to_excel("D:/Desktop/test.xlsx", index=False)

data = pd.read_excel("D:/Desktop/test.xlsx")
df = pd.DataFrame(data)

for column in columns:
    df[column] = df[column].apply(extract_numbers)
    try:
        df[column] = df[column].astype(float)
    except ValueError:
        pass

df.to_excel(r"D:\Desktop\HLA FINAL\HLA After Cleaning.xlsx", index=False)
