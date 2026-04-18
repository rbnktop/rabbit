import pandas as pd

datf = pd.read_excel('data.xlsx')

result = []



with open ('items.txt', 'r') as f:
    trgt = [line.strip() for line in f if line.strip()]


for item in trgt:
    found = []
    
    for column in datf.columns:
        if item in datf[column].values:
            found.append(column)
    
    if found:
        result.append(f"{item} found in columns: {', '.join(found)}")
    else:
        result.append(f"{item} not found in any column")

with open('results.txt', 'w') as f:
    for line in result:
        f.write(line + '\n')
    content = f.read()
    print(content)
    