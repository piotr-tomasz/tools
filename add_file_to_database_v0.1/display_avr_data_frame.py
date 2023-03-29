import pandas as pd

def display_avr_data_frame(database,date,case_name,function='mean'):
    return(pd.DataFrame(database[date][case_name][function]['data'],columns=database[date][case_name][function]['header']))