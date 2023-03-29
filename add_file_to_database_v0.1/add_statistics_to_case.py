
import pandas as pd
import numpy as np

def extract_from_files(case_data_frame,to_avr,value_type):
    df=pd.DataFrame(case_data_frame['data'],
                            columns=case_data_frame['header_columns'])
    to_avr.append(df.describe().iloc[value_type].values)
    return df

def avr(m_dict,fun='mean'):
    if fun == 'mean':
        colmn = 1
        name = 'mean'
    elif fun == 'std': #standard deviation
        colmn = 2
        name = 'std'
    elif fun == 'min':
        colmn = 3
        name = 'min'
    elif fun == 'max':
        colmn = 7
        name = 'max'
    else:
        colmn = 1
        name = 'mean'
    temp_value_list=[]
    for k1 in m_dict['data'].keys():
        df=extract_from_files(m_dict['data'][k1]['data_frame'],temp_value_list,colmn)
    temp_x=np.array(temp_value_list,dtype=float)
    df2=pd.DataFrame(temp_x.reshape(temp_x.shape[0],temp_x.shape[-1]),columns=df.describe().columns)
    return df2,name

def add_statistics_to_case(db_dictionary,function='mean'):
    
    for k in db_dictionary.keys():
        for k2 in db_dictionary[k].keys():
            df3,name=avr(db_dictionary[k][k2],function)
            db_dictionary[k][k2][name]={}
            db_dictionary[k][k2][name]['data']=df3.values.tolist()
            db_dictionary[k][k2][name]['header']=df3.columns.values.tolist()
    return db_dictionary

