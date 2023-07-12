import pandas as pd

from openpyxl import load_workbook

#объединить все листы sales в один
def to_one_sheet(sheet_path):
    new = pd.DataFrame()
    sheets = pd.ExcelFile(sheet_path)
    for name in sheets.sheet_names[2:]:
        #print(name)
        sheet = pd.read_excel(sheet_path,sheet_name=name)
        #if sheet.empty:
            #print(name, " is empty")
        #else:
        if not sheet.empty:
            if 'Unnamed: 0' in sheet.columns.tolist():

                df = sheet.dropna()

                df.columns = df.iloc[0]
                df = df.iloc[1:]
            else:

                df = sheet
            new = pd.concat([new, df])


    return new

#объединяем все листы sales в один
sheet_path = "./Copy of Copy of Export Store_Sales Jan'18 – Nov'19 FOR OMD final (1).xlsx"
new = to_one_sheet(sheet_path)

stores = pd.read_excel("./Copy of Copy of Export Store_Sales Jan'18 – Nov'19 FOR OMD final (1).xlsx", sheet_name="\u200bStores", header=[0,1])
print(stores.columns)
#корректируем номера TT на листе Stores
stores['№ ТТ']['Unnamed: 0_level_1'].apply(lambda x: int(x.split('N')[1]) if type(x)!= int else x )
for x in range(0, len(stores['№ ТТ']['Unnamed: 0_level_1'].tolist())):
    if type(stores['№ ТТ']['Unnamed: 0_level_1'][x]) != int:
        stores.at[x, ('№ ТТ', 'Unnamed: 0_level_1')] = stores['№ ТТ']['Unnamed: 0_level_1'][x].split('N')[1]

print(stores['№ ТТ']['Unnamed: 0_level_1'])

# выбираем строки Сибирь и Урал
sorted_sales = stores.loc[((stores['МЕСТОПОЛОЖЕНИЕ']['РЕГИОН'] == 'Сибирь') | (stores['МЕСТОПОЛОЖЕНИЕ']['РЕГИОН'] == 'Урал'))]

#записываем номера тт Сибири и Урала в список
sorted_tt = sorted_sales['№ ТТ']['Unnamed: 0_level_1'].tolist()
print(sorted_tt)

#выбираем из объединенного списка stores Сибирь и Урал

new_sorted_loc = new[new['№ TT'].isin(sorted_tt)]
print(new_sorted_loc.info())

#выбираем нужные даты
new_sorted_loc_date = new_sorted_loc[new_sorted_loc['НЕДЕЛЯ'] > pd.to_datetime('2018-01-01')]

print(new_sorted_loc_date.info())
#запись в новый файл
writer = pd.ExcelWriter('./new.xlsx', engine = 'xlsxwriter')
stores.to_excel(writer, sheet_name = 'Stores')
new.to_excel(writer,sheet_name = 'All Sales' )
new_sorted_loc_date.to_excel(writer, sheet_name = 'Sales sorted')
writer.close()



