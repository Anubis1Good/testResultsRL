import os
import pandas as pd

parent_folder = 'all_test_result'
raw_folder = 'all_test_results'
prefix = '_moex'
prefix = '_bitget'
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
result = result.reset_index()
result['ticker'] = result['name'].apply(lambda x: x.split('_')[1])
result['bot'] = result['name'].apply(lambda x: "_".join(x.split('_')[2:]))
result = result.sort_values(by=['ticker','total_min_fee_percent'],axis=0,ascending=[True,False])
result = result.reset_index(drop=True)

ranks = ['count','total_per','total_min_fee_percent','total_average_fee_percent','total_max_fee_percent']
data_sum = result.groupby('bot')[ranks].mean().sort_values('total_average_fee_percent',ascending=False).round(2)
rank_names = ["rank_"+r for r in ranks]
for r in ranks:
    # print(r)
    result["rank_"+r] = result.groupby("ticker")[r].rank(ascending=False, method="min")
avg_rank = result.groupby("bot")[rank_names].mean().sort_values('rank_total_average_fee_percent').round(2)
result2 = pd.concat([avg_rank, data_sum], axis=1)
result2 = result2.sort_values('rank_total_min_fee_percent')
result2 = result2.reset_index()
# print(result2)
# print(avg_rank)
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
    result2.to_excel(writer,sheet_name='bots_info')
    workbook = writer.book
    worksheet = writer.sheets['bots_info']
    for i, col in enumerate(result2.columns,start=1):
        width = max(result2[col].apply(lambda x: len(str(x))).max(), len(col))
        worksheet.set_column(i, i, width)
        worksheet.conditional_format(1, i, len(result2), i, {
            'type': 'cell',
            'criteria': 'less than',
            'value': 0,
            'format': workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})
        })
        # print(col)
        if col in rank_names:
            worksheet.conditional_format(1, i, len(result2), i, {
                'type': '3_color_scale',
                'max_color': '#DA9694',
                'mid_color': '#FFFFFF',
                'min_color': '#00B0F0'
            })
        else:
            worksheet.conditional_format(1, i, len(result2), i, {
                'type': '3_color_scale',
                'min_color': '#DA9694',
                'mid_color': '#FFFFFF',
                'max_color': '#00B0F0'
            })

    # writer._save()
