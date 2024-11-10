import pandas as pd

file_path = 'Moroccan mosques - Final.xlsx'
df = pd.read_excel(file_path, sheet_name='Sheet1')

def check_existing_mosque(df, item):
    existing_entry = df[df['item'] == item]
    if not existing_entry.empty:
        return existing_entry.iloc[0]['item']
    return 'NEW'

df['QID'] = df.apply(lambda row: check_existing_mosque(df, row['item']), axis=1)

quickstatements = []
for _, row in df.iterrows():
    if row['QID'] == 'NEW':
        quickstatements.append(f'CREATE')
        quickstatements.append(f'LAST\tLen\t"{row["En"]}"')
        quickstatements.append(f'LAST\tLar\t"{row["Ar"]}"')
        quickstatements.append(f'LAST\tLary\t"{row["Ary"]}"')
        quickstatements.append(f'LAST\tP31\tQ32815') 
        quickstatements.append(f'LAST\tP17\tQ1028') 
        quickstatements.append(f'LAST\tP131\t{row["Place ID"]}') 
        quickstatements.append(f'LAST\tDen\t"Mosque in Morocco"')
        quickstatements.append(f'LAST\tDar\t"مسجد في المغرب"')
        quickstatements.append(f'LAST\tDary\t"جامع في المغرب"')
    else:
        quickstatements.append(f'{row["item"]}\tLen\t"{row["En"]}"')
        quickstatements.append(f'{row["item"]}\tLar\t"{row["Ar"]}"') 
        quickstatements.append(f'{row["item"]}\tLary\t"{row["Ary"]}"')
        quickstatements.append(f'{row["item"]}\tP31\tQ32815')
        quickstatements.append(f'{row["item"]}\tP17\tQ1028')
        quickstatements.append(f'{row["item"]}\tP131\t{row["Place ID"]}') 
        quickstatements.append(f'{row["item"]}\tDen\t"Mosque in Morocco"')
        quickstatements.append(f'{row["item"]}\tDar\t"مسجد في المغرب"')
        quickstatements.append(f'{row["item"]}\tDary\t"جامع في المغرب"')
try:
    with open('quickstatements.txt', 'w', encoding='utf-8') as file:
        for statement in quickstatements:
            file.write(statement + '\n')
    print("QuickStatements have been saved to 'quickstatements.txt'")
except IOError as e:
    print(f"Error occurred while writing to file: {e}")
