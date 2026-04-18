import pandas as pd

datf = pd.read_excel('data.xlsx', engine='openpyxl')

result = []



with open ('items.txt', 'r') as f:
    trgt = [line.strip() for line in f if line.strip()]


for item in trgt:
    found = []
    
    for column in datf.columns:
        if item in datf[column].values:
            found.append(column)
    
    if found:
        formatted_cols = []

        for col in found:
            if hasattr(col, 'strftime'):
                formatted_cols.append(col.strftime('%m-%Y'))
            else:
                formatted_cols.append(str(col))
        
        result.append(f"{item} found in columns: {', '.join([str(formatted_cols)])}")
        
    else:
        result.append(f"{item} not found in any column")


with open('results.txt', 'w') as f:

    for line in result:
        f.write(line + '\n')


with open('results.txt', 'r') as f:
    print(f.read())

    