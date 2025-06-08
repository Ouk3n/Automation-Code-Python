import pandas as pd
import numpy as np 
import os
from datetime import datetime
import time




file_path_gia=input('Enter GIA file Path With extension:')
file_path_pred=input('Enter PRED file Path With extension:')

start = time.time()


df1=pd.read_excel(f'{file_path_gia}')
df2=pd.read_csv(f'{file_path_pred}')


df1.rename(columns={"Client Ref":"STONE_ID"},inplace=True)

df3=df1.merge(df2,on='STONE_ID',how='inner')


color_rank={'D': 1,   # Mapping Color rank
'E': 2,'F': 3,'G': 4,'H': 5,'I': 6,'J': 7,'K': 8,'L': 9,'M': 10,
'N':11,'O':12,'P':12,
'W-X': 13,
'FANCY': 14}

df3['Color'] = df3['Color'].replace({'*': 'FANCY'})

df3['Actual_color_rank'] = df3['Color'].map(color_rank)
df3['Pred_color_rank']=df3['Pred Color'].map(color_rank)   # Creating new columns of rank of color for both Actual and Pred


conditions = [
    df3['Pred_color_rank'] == df3['Actual_color_rank'],                # condition for new column GIA vs PRED color
    df3['Pred_color_rank'] > df3['Actual_color_rank'],    
    df3['Pred_color_rank'] < df3['Actual_color_rank']
]

choices = ['SAME', 'UP', 'DOWN']
df3["GIA vs PRED color"]=np.select(conditions,choices,default='Unknown')

florence_map = {'None': 1,'Faint': 2,'Medium': 3, # Mapping FL rank
'Strong': 4,'Vstrong': 5
}

df3['Fluorescence']=df3['Fluorescence'].fillna('NONE')  # filling nan values with NONE as python dosent reconize the none in excel file given to us (NONE is a FL type)

df3['Actual_fl _rank']=df3['Fluorescence'].map(florence_map)  # Creating new columns of rank of FL for both Actual and Pred
df3['Pred_fl_rank']=df3['Pred Fluorescence'].map(florence_map)

conditions1 = [                                       
    df3['Pred_fl_rank'] == df3['Actual_fl _rank'],       # condition for new column Fl GIA VS PRED
    df3['Pred_fl_rank'] > df3['Actual_fl _rank'],
    df3['Pred_fl_rank'] < df3['Actual_fl _rank']
]

choices1 = ['SAME', 'UP', 'DOWN']
df3["Fl GIA VS PRED"]=np.select(conditions1,choices1,default='Unknown')


clarity_rank={"FL":1,"IF":2,"VVS1":3,"VVS2":4,"VS1":5,"VS2":6,"SI1":7,"SI2":8,"SI3":9,"I1+":10,"I1":11,"I1-":12,"I2++":13}  # mapping Clarity rank

df3['Actual_clarity_Rank']=df3['Clarity'].map(clarity_rank)       # Creating new columns of rank of Clarity for both Actual and Pred
df3['Pred_clarity_Rank']=df3['Pred Clarity'].map(clarity_rank)  

conditions2 = [
    df3['Pred_clarity_Rank'] == df3['Actual_clarity_Rank'],
    df3['Pred_clarity_Rank'] > df3['Actual_clarity_Rank'],
    df3['Pred_clarity_Rank'] < df3['Actual_clarity_Rank']
]

choices2 = ['SAME', 'UP', 'DOWN']
df3["Clarity GIA vs PRED"]=np.select(conditions2,choices2,default='Unknown')

df3['Cut']=df3['Cut'].str.title()
df3['Pred Cut']=df3['Pred Cut'].str.title()                      # Correcting format for both Actual and Pred as some data are unmatched
df3['Pred Symmetry']=df3['Pred Symmetry'].str.title()
df3['Symmetry']=df3['Symmetry'].str.title()
df3['Polish']=df3['Polish'].str.title()
df3['Pred Polish']=df3['Pred Polish'].str.title()


Cut_Rank={"Ideal":1,"Excl":2,"Vgood":3,"Good":4,"Fair":5,"Poor":6} # mapping Cur rank 
df3["Actual_cut_rank"]=df3["Cut"].map(Cut_Rank)
df3['Pred_cut_rank']=df3['Pred Cut'].map(Cut_Rank).fillna(2)  # fiiling nan values after Mapping them as all other not mentioned in Rank map for cut are Excl 

conditions3 = [
    df3['Pred_cut_rank'] == df3['Actual_cut_rank'],  # condition for new column CUT GIA vs PRED
    df3['Pred_cut_rank'] > df3['Actual_cut_rank'],
    df3['Pred_cut_rank'] < df3['Actual_cut_rank']
]

choices3 = ['SAME', 'UP', 'DOWN']
df3["CUT GIA vs PRED"]=np.select(conditions3,choices3,default='Unknown')

polish_rank={"Ideal":1,"Excl":2,"Vgood":3,"Good":4,"Fair":5,"Poor":6} # mapping Polish rank

df3['Actual_polish_rank']=df3['Polish'].map(polish_rank)
df3['Pred_polish_rank']=df3['Pred Polish'].map(polish_rank)  # Creating new column of rank of Polish for both Actual and Prediction

conditions4 = [
    df3['Pred_polish_rank'] == df3['Actual_polish_rank'],  # condition for new column Polish GIA vs PRED
    df3['Pred_polish_rank'] > df3['Actual_polish_rank'],
    df3['Pred_polish_rank'] < df3['Actual_polish_rank']
]

choices4 = ['SAME', 'UP', 'DOWN']
df3["POLISH GIA vs PRED"]=np.select(conditions4,choices4,default='Unknown')

symmetry_rank={"Ideal":1,"Excl":2,"Vgood":3,"Good":4,"Fair":5,"Poor":6}  # mapping symmetry rank

df3['Actual_symmetry_rank']=df3['Symmetry'].map(polish_rank)
df3['Pred_symmetry_rank']=df3['Pred Symmetry'].map(polish_rank)  # creating new column of rank of symmetry for both Actual and Prediction


conditions5 = [
    df3['Pred_symmetry_rank'] == df3['Actual_symmetry_rank'],
    df3['Pred_symmetry_rank'] > df3['Actual_symmetry_rank'],
    df3['Pred_symmetry_rank'] < df3['Actual_symmetry_rank']
]

choices5 = ['SAME', 'UP', 'DOWN']
df3["Symmetry GIA vs PRED"]=np.select(conditions5,choices5,default='Unknown')


def assign_size_range(size):    # Function for mapping Weight into size 
    if  size<=0.30:
        return '0.30 Down'
    elif 0.31 <= size <= 0.499:
        return '0.31-0.499'
    elif 0.50 <= size <= 0.699:
        return '0.50-0.699'
    elif 0.70<= size <=0.999:
        return '0.70-0.999'
    elif 1.00<= size <=1.499:
        return '1.00-1.499'
    elif 1.50<= size <=1.999:
        return '1.50-1.999'
    else:
        return '2 & Up'
    

df3['Size Range']=df3['Weight'].apply(assign_size_range) # Creating new column for Size Range based on function 


df3=df3.loc[:,['STONE_ID','Color','Clarity','Cut','Polish','Symmetry','Fluorescence','Return Date','Pred Shape', 'Pred Cut', 'Pred Color', 'Pred Clarity', 'Pred Polish',
       'Pred Symmetry', 'Pred Fluorescence','GIA vs PRED color','Fl GIA VS PRED','Clarity GIA vs PRED','CUT GIA vs PRED','POLISH GIA vs PRED','Symmetry GIA vs PRED',
       'Size Range']]       # Including only needed columns for our report final data


today_str = datetime.today().strftime('%Y-%m-%d')
filename=f'GIA VS PRED_{today_str}.xlsx'

df3.to_excel(filename,index=False)
print(df3.head(1))



end = time.time()
print(f'Execution time: {end - start:.4f} seconds')
