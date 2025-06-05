import pandas as pd
import numpy as np
from datetime import datetime


file_path=input('ENTER FILE NAME:')

df=pd.read_excel(f'{file_path}.xlsx')

df=df.loc[df['PEN_INCEN']!='INCENTIVE']  # PAN_INCEN
df=df.loc[df['PEN_INCEN']!='REPAIR']


df=df.loc[~df['REVIEW_FLG'].str.startswith('R', na=False)] # removing startswith('R') rows in review_flg and keepin NaN values


values=['CROSS ENTRY','GRADER MISTAKE','LAB RECUT','UPGRADE','CROSS GRADER MISTAKE','25%','100%','75%','50%','POINTED']  # values list for DAMAGE_REMARK
df=df.loc[~df['DEMAGE_REMARK'].astype(str).isin(values)]
# when we load the data using read_excel excel automatically coverts 25% as 0.25 and python reads it as 0.25 so astype help us remove them 



df=df.loc[df['BREAKAGEDUETO']!='QC EARROR'] # BREAKAGEDUETO


values_reason=['CROSS ENTRY','PLAN MISSED','UPGRADE','QC ERROR','GRADER MISTAKE'] # values list for REASON 
df=df.loc[~df['REASON'].isin(values_reason)]


values_subreason=['UPGRADE','CROSS ENTRY','GRADER MISTAKE'] # values list for subreason
df=df.loc[~df['SUBREASON'].isin(values_subreason)]


def final_column(row):
    if pd.notna(row['REVIEWED_RATE_DIFF']):
        return row['REVIEWED_RATE_DIFF']    # defininf function for new column final
    return row['RATE_DIFF']
        
df['Final']=df.apply(final_column,axis=1)


df['New_Remark1'] = ''  # creating column new_remark1 with empty values
df['New_Remark1']=df['REMARK1'] 
def check_hash(x):
    if isinstance(x, str) and '#' in x:
        return x                           # as REMARK1 contains some int values which are not interable so we use isinstance
    return np.nan

df['New_Remark1'] = df['New_Remark1'].apply(check_hash)

df['New_Remark1']=df['New_Remark1'].fillna(df['BCNO'])

today_str = datetime.today().strftime('%Y-%m-%d')
df.to_excel(f'Damaged_Data_cleaned_{today_str}.xlsx',index=False)R MISTAKE'] # values list for REASON 
df=df.loc[~df['REASON'].isin(values_reason)]


values_subreason=['UPGRADE','CROSS ENTRY','GRADER MISTAKE