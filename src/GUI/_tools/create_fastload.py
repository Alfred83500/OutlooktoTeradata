import pandas as pd
import csv
from datetime import date


def create_fastload_execute(name_format_file,table_name,log_on_file):
    print("############### DEBUT GENERATION FASTLOAD  ##########################")
    with open(name_format_file, newline='') as f:
      reader = csv.reader(f,delimiter=';')
      row1 = next(reader)


    champ_table_a_creer = ''
    champ_define = ''
    champ_insert = ''
    champ_values = ''

    


    for name in row1:
        champ_table_a_creer += (f'{str(name).upper()} VARCHAR(250) CHARACTER SET LATIN NOT CASESPECIFIC,\n')
    champ_table_a_creer += ('DATE_HEURE_INSERT TIMESTAMP(6) NOT NULL DEFAULT CURRENT_TIMESTAMP(6)')

    search_text_table = 'champs_table_a_creer'
    search_text_define = 'champs_define'
    search_text_insert = 'champs_insert'
    search_text_values = 'champs_values'
    search_name_file = 'name_file'
    search_name_table = 'name_table'
    log_on_text = 'LOG_ON_TEXT'
    



    for i in range (len(row1)-1):
        champ_define += (f'{str(row1[i]).upper()} (VARCHAR(250)),\n')
    champ_define += (f'{str(row1[len(row1)-1]).upper()} (VARCHAR(250))')

    for i in range (len(row1) - 1):
        champ_insert += (f'{str(row1[i]).upper()},\n')
    champ_insert += (f'{str(row1[len(row1)-1]).upper()}')

    for i in range (len(row1) - 1):
        champ_values += (f':{str(row1[i]).upper()},\n')
    champ_values += (f':{str(row1[len(row1)-1]).upper()}')






    with open('fich_modif.txt', 'r') as file:
        data = file.read()
        data = data.replace(search_text_table, champ_table_a_creer)
        data = data.replace(search_text_define, champ_define)
        data = data.replace(search_text_insert, champ_insert)
        data = data.replace(search_text_values, champ_values)
        data = data.replace(search_name_file, name_format_file)
        data = data.replace(log_on_text, log_on_file)
        #TODO Adapt for different Databases
        data = data.replace(search_name_table,table_name)





    with open("fich_modif_todate.txt", 'w') as f:
        f.write(data)



    print("############### FIN GENERATION FASTLOAD  ##########################")








 
   

