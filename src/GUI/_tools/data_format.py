import pandas as pd
from datetime import datetime,timedelta,date

def format_data(name_file):
    print("###############  FORMATTAGE DONNEES  ##########################")

    df = pd.read_csv(name_file,sep=';',encoding='ansi')
    list_forbidden = ['TITLE','GROUP','USER','ALIAS']

    
    char_to_tej = ["."," ","°"]
    #TODO generic for all teradata keyword
    for col in list(df):
        if col.upper() in list_forbidden:
            df = df.rename(columns = {col:f'{col}_1'})
        
        elif "." in col:
           df = df.rename(columns = {col:col.split(".")[1]})
        elif "° " in col:
           df = df.rename(columns = {col:col.replace("° ","")})
        elif " " in col:
            df = df.rename(columns = {col:col.replace(" ","_")})
    new_filename = name_file.split(".")[0]+"_"+str(date.today())+".csv"
    df.to_csv(new_filename,index=False,sep=';',encoding='ansi')
    print("############### FIN  FORMATTAGE DONNEES  ##########################")
    return new_filename

