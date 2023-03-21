from retrieve_csv import retreive_mail
from data_format import format_data
from create_fastload import create_fastload_execute
import os

def main_manual():
    subject_name = input("Objet du mail: ")
    sender_name = input("Auteur du mail: ")

    name_file = retreive_mail(subject_name,sender_name)
    name_format_file = format_data(r"data/"+name_file)
    create_fastload_execute(os.path.abspath(name_format_file),subject_name,True)


    cmd_fastload =  f"fastload < {os.path.abspath('fich_modif_todate.txt')} >> log/log_fastload.txt "
    print("############### DEBUT FASTLOAD ##########################")
    #os.system(cmd_fastload)
    print("############### FIN FASTLOAD ##########################")
    with open('log/log_fastload.txt', 'r') as f:
        last_line = f.readlines()[-2]
    print(last_line)

#### Cas spécial Armonde pour les 4 fichiers SYS_USER #####
def main_auto(subject_name):
    sender_name = 'DAVIDSON CONSULTING - SHIVBARAN, Armonde'

    name_file = retreive_mail(subject_name,sender_name)
    name_format_file = format_data(r"data/"+name_file)
    create_fastload_execute(os.path.abspath(name_format_file),subject_name,False)

    cmd_fastload =  f"fastload < {os.path.abspath('fich_modif_todate.txt')} >> log/log_fastload.txt "
    print("############### DEBUT FASTLOAD ##########################")
    os.system(cmd_fastload)
    print("############### FIN FASTLOAD ##########################")
    with open('log/log_fastload.txt', 'r') as f:
        last_line = f.readlines()[-2]
    print(last_line)
#############################################################

if __name__ == '__main__':

    manual_or_auto = input("voulez vous lancer en mode manuel (1) ou automatique (0)")
    if manual_or_auto == "1":
        main_manual()
    else:
        confirmation = input("confirmation version automatique oui(1)/non(0)?")
        if confirmation == "1":
            list_subject_names = ['04_SYS_USER_Export_vrai_avant_2022','03_SYS_USER_Export_vrai_apres_2022',
                        '02_SYS_USER_Export_Faux_avant_2022','01_SYS_USER_Export_Faux_apres_2022']
            for subject_name in list_subject_names:
                main_auto(subject_name)
        else:
            print("attention la prochaine fois")
