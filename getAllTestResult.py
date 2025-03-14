import os
import pandas as pd

parent_folder = 'all_test_result'
raw_folder = 'all_test_results'
prefix = '_moex'
# prefix = '_bitget'
raw_folder += prefix
raw_folder = os.path.join(parent_folder,raw_folder)
raw_files = os.listdir(raw_folder)
df1 = None
for i,rw in enumerate(raw_files):
    file_path = os.path.join(raw_folder,rw)
    df = pd.read_excel(file_path)
    try:
        df = df.drop(['mean_price','total_min_fee','total_average_fee','total_max_fee'],axis=1)
    except:
        pass
    if i == 0:
        df1 = df
    else:
        df1 = pd.concat([df1,df],axis=0)

for col in df1.columns:
    if 'Unnamed' in col:
        df1 = df1.drop(col,axis=1)

result = df1.groupby('name').sum()
result = result.sort_values(by='total_average_fee_percent',axis=0,ascending=False)
result = result.reset_index()
file_name = f'Total_All_Test_Result_{prefix}.xlsx'

with pd.ExcelWriter(file_name, engine='xlsxwriter') as writer:  
    result.to_excel(writer,sheet_name='total')
    workbook = writer.book
    worksheet = writer.sheets['total']
    for i, col in enumerate(result.columns,start=1):
        width = max(result[col].apply(lambda x: len(str(x))).max(), len(col))
        worksheet.set_column(i, i, width)
        worksheet.conditional_format(1, i, len(result), i, {
            'type': 'cell',
            'criteria': 'less than',
            'value': 0,
            'format': workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})
        })
        worksheet.conditional_format(1, i, len(result), i, {
            'type': '3_color_scale',
            'min_color': '#DA9694',
            'mid_color': '#FFFFFF',
            'max_color': '#00B0F0'
        })
    # writer._save()
