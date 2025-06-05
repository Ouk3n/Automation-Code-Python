import pandas as pd
import numpy as np 
from datetime import datetime
import os
import time 


yns_status_file_path=input('ENTER YNS FILE ABSOLUTE PATH:') 
os_data_file_path=input('ENTER OS FILE ASBOLUTE PATH:')

start = time.time()

df1=pd.read_excel(f'{yns_status_file_path}',skiprows=1)
df2=pd.read_excel(f'{os_data_file_path}')

if not os.path.exists(yns_status_file_path):
    print("❌ YNS file not found. Check the path.")
    exit()


if not os.path.exists(os_data_file_path):
    print("❌ OS file not found. Check the path.")
    exit()

df1=df1[df1['IS_VIRTUAL']!='Y']
df1.reset_index(drop=True,inplace=True)# droping index 
df2.columns=df2.columns.str.upper() # to make columns match for merge (make sure OS file columns are UPPER Case as well as YNS Status)



df3=df1.merge(df2,on='DEPT',how='left')  # merging both YNS status and OS based on DEPT 




def update_dept_XRAY(row):                                                                    # Update function for XRAY
    if pd.isna(row['SHAPE']) or str(row['SHAPE']).strip() == "":
        return 'CLV'
    else:
        return 'XRAY & 4P'


xray_depts = ['XRAY SEMI POLISH', 'XRAY AUTO POLISH', 'XRAY ANYCUT BLOCKING','X RAY SCAN']


mask_XRAY = df3['DEPT'].isin(xray_depts)
df3.loc[mask_XRAY, 'DEPT GRP'] = df3.loc[mask_XRAY].apply(update_dept_XRAY, axis=1) # Applying above function on DEPT GRP for XRAY(DEPT)





def update_dept_DNA(row):                                                                      # Update function for DNA
    if pd.isna(row['SHAPE']) or str(row['SHAPE']).strip() == "":
        return 'CLV'
    else:
        return 'LAB'


DNA_depts = ['DNA']
mask_DNA = df3['DEPT'].isin(DNA_depts)
df3.loc[mask_DNA, 'DEPT GRP'] = df3.loc[mask_DNA].apply(update_dept_DNA, axis=1)  # Applying above function on DEPT GRP for DNA






def update_dept__MFG(row):                                                      # Update function for MFG
    if pd.isna(row['SHAPE']) or str(row['SHAPE']).strip() == "":
        return 'CLV'
    else:
        return 'MFG'

MFG_depts = ['MFG BOILING']
mask_MFG = df3['DEPT'].isin(MFG_depts)
df3.loc[mask_MFG, 'DEPT GRP'] = df3.loc[mask_MFG].apply(update_dept__MFG, axis=1) # Applying above function on DEPT GRP for MFG


df3['JDATE'] = pd.to_datetime(df3['JDATE'])

# Calculate days difference (today - Jdate)
df3['days'] = (datetime.today() - df3['JDATE']).dt.days

df3['JDATE'] = df3['JDATE'].dt.strftime('%Y-%m-%d')



df3=df3[df3['DISP_LOTNO']!='SAMPLE']              # Removing SAMPLE values from DISP_LOTNO


def entity_type(row):                                                       # Update function based on ENTITYTYPE  & DEPT 
    if row['ENTITYTYPE']=='OTHER PERSON':
        if row['DEPT'] in ['DEV DIAMONDS SOLUTION','INFINITY ENTERPRISES']:
            return '4P PARTY'
        else:
            return 'PARTY'
        
mask_entity = df3['ENTITYTYPE'] == 'OTHER PERSON'
df3.loc[mask_entity, 'DEPT GRP'] = df3.loc[mask_entity].apply(entity_type, axis=1)  # Applying above function on DEPT GRP




def lot_changes(row):                                                        # Update function based on DISP_LOTNO on new column LOT
    if row['DISP_LOTNO'] in ['BOM-REP-10','BOM-REP-9']:
        return 'REPAIR'
    elif row['DISP_LOTNO'] in ['LK-N','CHARITRA-N','GAUTAM-N']:
        return 'JOB WORK'
    else:
        return 'REGULAR'
    

df3.loc[:, 'LOT'] = df3.apply(lot_changes, axis=1)   # Applying above function with also creating LOT new column (this function would create new column LOT)

today_str = datetime.today().strftime('%Y-%m-%d')
filename = f"DATA_{today_str}.xlsx"

df3.to_excel(filename,index=False)
print(df3.head())


end = time.time()
print(f'Execution time: {end - start:.4f} seconds')
NFINITY ENTERPRISES']:
            return '4P PARTY'
        else:
            return 'PARTY'
        
mask_entity = df3['ENTITYTYPE'] == 'OTHER PERSON'
df3.loc[mask_entity, 'DEPT GRP'] = df3.loc[mask_entity].apply(e