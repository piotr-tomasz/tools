import pandas as pd
import matplotlib.pyplot as plt

def plot_massflow_from_avr_data(database,date,case_name,function='mean',figsize=(8,8)):
    df_temp=pd.DataFrame(database[date][case_name][function]['data'],columns=database[date][case_name][function]['header'])
    plt.figure(figsize=figsize)
    plt.scatter(df_temp.iloc[:,2],df_temp.iloc[:,-2])
    plt.scatter(df_temp.iloc[:,2],df_temp.iloc[:,-1])
    plt.plot(df_temp.iloc[:,2],df_temp.iloc[:,-2])
    plt.plot(df_temp.iloc[:,2],df_temp.iloc[:,-1])
    plt.xlabel(df_temp.columns[2])
    plt.ylabel(df_temp.columns[-2])
    plt.legend([df_temp.columns[-2],df_temp.columns[-1]])
    plt.grid()