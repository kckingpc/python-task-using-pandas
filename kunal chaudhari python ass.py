import pandas as pd
from openpyxl import load_workbook

# Load the input sheets
sheet1 = pd.read_excel('/content/kc.xlsx', sheet_name='Sheet2')
sheet2 = pd.read_excel('/content/kc.xlsx', sheet_name='Sheet3')

#first ans
merged = pd.merge(sheet1, sheet2, left_on='User ID', right_on='uid')
result = merged.groupby('Team Name')['total_statements','total_reasons'].mean().reset_index()
result.index += 1
result.index.name = 'Team rank'

result1=result.sort_values(by='total_statements', ascending=False)

#2nd ans
result2=sheet2.sort_values(by='total_statements', ascending=False)

book = load_workbook('/content/kc.xlsx')
writer = pd.ExcelWriter('/content/kc.xlsx', engine='openpyxl')
writer.book=book
writer.sheet = dict((ws.title, ws) for ws in book.worksheets)

output = pd.DataFrame(result1)
output.to_excel(writer, sheet_name='Sheet4', index=True)


output = pd.DataFrame(result2)
output.to_excel(writer, sheet_name='Sheet5', index=False)

writer.save()