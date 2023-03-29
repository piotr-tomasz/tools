import os
from add_data_file import *

def add_to_data_dictionary2(folder_path,d_measurements,general_date='20000101',case_name='case_1',description='default',additional_info='default'):
    path=folder_path
    dir_list=os.listdir(path)
    file_list=[dir_list[f] for f in range(len(dir_list)) if os.path.isfile(path+'\\'+dir_list[f])] #delete folder in directory from dir list

    def select_data_to_sort(file_name):
        '''sort files in directory'''
        return int(file_name.split('.txt')[0].split('_')[-1])
    file_list.sort(key=select_data_to_sort)
    
    if general_date not in d_measurements.keys():
        d_measurements[general_date]={}
    if case_name not in d_measurements[general_date].keys():
        d_measurements[general_date][case_name]={}
    d_measurements[general_date][case_name]['data']={}
    for i in file_list:
        file_name=i
        d_measurements=add_data_file(file_name,path,d_measurements,general_date,case_name,description,additional_info)
    return d_measurements