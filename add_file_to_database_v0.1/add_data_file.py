import pandas as pd
from header_fit import *
from SLPM_to_sgps import *

def add_data_file(file_name,path,d_measurements,general_date,case_name,description,additional_info):
    name=file_name.split('.txt')[0]
    full_path=path+'\\'+ file_name
    df=pd.read_csv(full_path,sep='\t',skiprows=1,header=[0,1])
    df.columns=header_fit(len(df.columns))
    column_list=pd.read_csv(full_path,sep='\t').columns.tolist()
    date,time,data_rate_Hz=column_list[0],column_list[1],int(column_list[3])

    ix_nb=[]
    for x,y in enumerate(df.columns):
        if y.find('SLPM')>=0:
            ix_nb.append(x)
    
    if len(ix_nb)==2:
        for i,j in enumerate(ix_nb):
            new_column_list='Mass flow rate - A'+ str(i+1) +' [g/s]'
            df[new_column_list]=SLPM_to_sgps(df.iloc[:,j])
            column_list.append( new_column_list)

    elif len(ix_nb)==1:
        new_column_list='Mass flow rate [g/s]'
        df[new_column_list]=SLPM_to_sgps(df.iloc[:,ix_nb[0]])
        column_list.append(new_column_list)  

    d_measurements[general_date][case_name]['data'][name]={}
    d_measurements[general_date][case_name]['data'][name]['file_info']={}
    d_measurements[general_date][case_name]['data'][name]['time_date']={}
    d_measurements[general_date][case_name]['nb_of_columns']=len(df.columns)
    d_measurements[general_date][case_name]['description'] = description
    d_measurements[general_date][case_name]['data'][name]['file_info']['file_name']=file_name
    d_measurements[general_date][case_name]['data'][name]['file_info']['path']=full_path
    d_measurements[general_date][case_name]['data'][name]['time_date']['date']=date
    d_measurements[general_date][case_name]['data'][name]['time_date']['hour']=time
    d_measurements[general_date][case_name]['data'][name]['data_rate_Hz']=data_rate_Hz
    d_measurements[general_date][case_name]['data'][name]['data_frame']={}

    d_measurements[general_date][case_name]['data'][name]['data_frame']['data'] = df.to_numpy().tolist()
    d_measurements[general_date][case_name]['data'][name]['data_frame']['header_columns'] = df.columns.tolist()
    d_measurements[general_date][case_name]['data'][name]['additional_info']=additional_info

    d_measurements[general_date][case_name]['data'][name]['data_avr']={}
    d_measurements[general_date][case_name]['data'][name]['data_avr']['data']=df.describe().to_numpy().tolist()
    d_measurements[general_date][case_name]['data'][name]['data_avr']['header_columns'] = df.describe().columns.tolist()
    d_measurements[general_date][case_name]['data'][name]['data_avr']['header_rows'] = df.describe().index.tolist()
    
    return d_measurements