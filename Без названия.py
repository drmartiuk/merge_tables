import pandas as pd
import xlsxwriter

V_1=pd.read_excel('./start.xlsx',skiprows=0,sheet_name='Sheet1').reset_index()
display(V_1)
display(V_1.shape)
V_2=pd.read_excel('./start_2.xlsx',skiprows=0,sheet_name='Sheet1').reset_index()
display(V_2)
display(V_2.shape)

temp=V_1[['index', 'X', 'V №', 'V']].merge(V_2[['index', 'V №', 'X']],left_on='V',right_on='V №',how='outer')
temp['X_final'] =temp.apply(lambda x: x['X_y'] if pd.isnull(x['X_x'])  else x['X_x'], axis=1)
temp=temp[['index_x','index_y','X_final']].sort_values(by='X_final')
V_1=temp.merge(V_1,left_on='index_x',right_on='index',how='left')
display(V_1)
display(V_1.shape)
V_2=temp.merge(V_2,left_on='index_y',right_on='index',how='left')
display(V_2)
display(V_2.shape)


path=r"./final.xlsx"
writer = pd.ExcelWriter(path, engine = 'xlsxwriter')
V_1.to_excel(writer, sheet_name = 'V_1')
V_2.to_excel(writer, sheet_name = 'V_2')

writer.close()
