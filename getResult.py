import os
import pandas as pd
import matplotlib.pyplot as plt

cur_name = '01.03.2024d'
raw_folder = 'testResults/onlineTests/'
raw_folder = os.path.join(raw_folder,cur_name)
result_name = raw_folder.split('/')[-1]

raw_files = os.listdir(raw_folder)

min_fee: float = 0.0004
max_fee: float = 0.0012
average_fee = (max_fee + min_fee)/2
df_main = pd.DataFrame(columns=['name','total_abs','count','mean_price'])
# df_main = pd.DataFrame(columns=['name','total_abs','total_per','total_min_fee_percent','total_max_fee_percent','total_average_fee_percent','count'])
equity_chart_folder = 'equity_chart'
if not os.path.exists(equity_chart_folder):
    os.mkdir(equity_chart_folder)
path_imgs = os.path.join(equity_chart_folder,result_name)
if not os.path.exists(path_imgs):
    os.mkdir(path_imgs)

for rw in raw_files:
    rw_path = os.path.join(raw_folder,rw)
    df = pd.read_json(rw_path)
    df = df.drop(0,axis=0)
    name_bot = rw.replace('.json','')
    if len(df.index) > 2:
        if pd.isnull(df.iloc[-1]['close_time']):
            index = df.iloc[-1].name
            df = df.drop(index,axis=0)
        df_w = pd.DataFrame({
            'name':[name_bot],
            'total_abs':[df.iloc[-1]['total']],
            'count':[df.iloc[-1]['count']],
            'mean_price':[df['open_price'].mean()]
        })
        df_main = pd.concat([df_main,df_w],axis=0)
        plt.plot(df['total'],color='blue')
        full_name_img = os.path.join(path_imgs,name_bot + '.png')
        plt.savefig(full_name_img)
        plt.close()
del df,df_w
df_main['total_min_fee'] = df_main['total_abs'] - (df_main['mean_price'] * min_fee * df_main['count'] * 2)
df_main['total_average_fee'] = df_main['total_abs'] - (df_main['mean_price'] * average_fee * df_main['count'] * 2)
df_main['total_max_fee'] = df_main['total_abs'] - (df_main['mean_price'] * max_fee * df_main['count'] * 2)
df_main['total_per'] = (df_main['total_abs']/df_main['mean_price']) * 100
df_main['total_min_fee_percent'] = (df_main['total_min_fee']/df_main['mean_price']) * 100
df_main['total_average_fee_percent'] = (df_main['total_average_fee']/df_main['mean_price']) * 100
df_main['total_max_fee_percent'] = (df_main['total_max_fee']/df_main['mean_price']) * 100
df_main = df_main.sort_values(by='total_average_fee_percent',axis=0,ascending=False)
df_main = df_main.reset_index(drop=True)
file_name = 'Total_' + result_name + '.xlsx'
path_df_main = os.path.join('total_files',file_name)
with pd.ExcelWriter(path_df_main, engine='xlsxwriter') as writer:  
    df_main.to_excel(writer,sheet_name='total')
    workbook = writer.book
    worksheet = writer.sheets['total']
    for i, col in enumerate(df_main.columns):
        width = max(df_main[col].apply(lambda x: len(str(x))).max(), len(col))
        worksheet.set_column(i, i, width)
    writer._save()