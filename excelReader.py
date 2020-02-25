import pandas as pd

excel_file = 'D:\DotnetCore\python\movies.xls'
movies_sheet1 = pd.read_excel(excel_file)
print(movies_sheet1.head())
print('Hello World')

