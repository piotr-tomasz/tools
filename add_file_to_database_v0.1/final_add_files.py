from add_to_data_dictionary2 import *
from sort_dictionary import *
from add_statistics_to_case import *
import json


def final_add_files(files_to_add, database_path, data_exists=False):
    if data_exists==True:
        data=input('Write database name: ')
        if data in locals():
            pass
        else:
            data={}
    else:
        data={}
    for k in files_to_add.keys():
        folder_path,general_date,case_name,description,additional_info = \
        files_to_add[k]['folder_path'],files_to_add[k]['general_date'],files_to_add[k]['case_name'],files_to_add[k]['description'],files_to_add[k]['additional_info']
        add_to_data_dictionary2(folder_path,data,general_date,case_name,description,additional_info)

    data = sort_dictionary(data)

    add_statistics_to_case(data,'mean') # calculate selected statisticts for case

    with open(database_path, "w") as outfile:
        json.dump(data, outfile)
    return data