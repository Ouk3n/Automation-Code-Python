import pandas as pd
import numpy as np 


file_path_1=input('Enter orginal file absolute path without extension:')
file_path_2=input('Enter empty file absolute file path:')


df1=pd.read_excel(f'{file_path_1}.xlsx')
df2=pd.read_excel(f'{file_path_2}.xlsx')


df1 = df1[~df1.apply(lambda row: (row == df1.columns).all(), axis=1)]

df2 = df2.iloc[0:0]


df1.dropna(subset=['STONE_ID'],inplace=True) # droping NAN values from df1 which as all the parameters of data as we wont need NAN stone_id 


df2['Stone_ID']=df1['STONE_ID']

df2['No'] = range(1,len(df2)+1)


df2.reset_index(drop=True,inplace=True)


df3=df1.copy()

df3.rename(columns={'LOTNO':'Lot No','PKTNO':'Pkt No','CTS':'Pol Cts','TABLE1':'TABLE','SHAPE_ID':'Shape','PURITY':'Clarity','COLOR':'Color','CUT':'Cut','HEIGHT':'Height',
                    'DIAMETER':'Diameter','L/W':'L / W','FLUOR':'FLUORESCENCE','POLISH':'Polish','SYMMETRY':'Symmetry'},inplace=True)


cols_to_fill = ['Lot No','Pkt No','Pol Cts','TABLE', 'Shape', 'Clarity',
       'Color', 'Cut', 'Height', 'Diameter', 'L / W', 'FLUORESCENCE', 'Polish',
       'Symmetry', 'GIRDLE']

for col in cols_to_fill:
    mapping = df3.set_index('STONE_ID')[col]
    df2[col] = df2[col].fillna(df2['Stone_ID'].map(mapping))


df2['FLUORESCENCE']=df2['FLUORESCENCE'].fillna('None')


chunk_size = 1000


for i in range(0, len(df2), chunk_size):
    chunk = df2.iloc[i:i+chunk_size]
    chunk.to_excel(f"output_part_{i//chunk_size + 1}.xlsx", index=False)





    




