from __future__ import print_function
from mailmerge import MailMerge
from time import gmtime, strftime
from datetime import datetime, timedelta
from configparser import SafeConfigParser
import os
import shutil
import glob
import datetime
import sys
import openpyxl
import configparser


#----- FUNC BLOCK BEGIN


def logo():
    print(
        '\n'
        '███╗   ███╗ █████╗ ██████╗ ██╗  ██╗██╗   ██╗██████╗ ██╗ ██████╗ ███╗   ██╗ \n'   
        '████╗ ████║██╔══██╗██╔══██╗██║ ██╔╝██║   ██║██╔══██╗██║██╔═══██╗████╗  ██║  \n'  
        '██╔████╔██║███████║██████╔╝█████╔╝ ██║   ██║██████╔╝██║██║   ██║██╔██╗ ██║  \n'   
        '██║╚██╔╝██║██╔══██║██╔══██╗██╔═██╗ ██║   ██║██╔══██╗██║██║   ██║██║╚██╗██║  \n'   
        '██║ ╚═╝ ██║██║  ██║██║  ██║██║  ██╗╚██████╔╝██║  ██║██║╚██████╔╝██║ ╚████║  \n'   
        '╚═╝     ╚═╝╚═╝  ╚═╝╚═╝  ╚═╝╚═╝  ╚═╝ ╚═════╝ ╚═╝  ╚═╝╚═╝ ╚═════╝ ╚═╝  ╚═══╝  \n'   
        ' ---------------RT Report update machine---------------------   \n'                                                                                                                                                
    )

def make_space():
    os.system('cls')
    logo()

#req_folder_location = input('Location of the REQ File: ')
make_space()
print('Please enter folder location where test_data.ini is: ')
folder_location = input('Enter it here: ')

os.system('cls')
make_space()

#Change when finall app is ok
#folder_location = r'E:\temp\07Feb19'
os.chdir(folder_location)

#Read config file
config = configparser.ConfigParser()
config.read('test_data.ini')

client_name = config['TEST']['client']
client_adr = config['TEST']['client_adr']
date = config['TEST']['date']
job_no = config['TEST']['job_no']
test_type = config['TEST']['test_type']
report_no = config['TEST']['report_no']
date_formated = config['TEST']['date_formated']
main_report_file_name = "Report " + report_no + " RT " + client_name + " Job " + job_no + ".docx"

xls_file_name = "REQ " + date.upper() + " " + client_name + ".xlsx"

print('Using file name: ' + xls_file_name)
print(' \n --------- ------- ----- ------ ------ -----')

wb = openpyxl.load_workbook(xls_file_name)
sheet = wb['Sheet1']

#Display all welders from weld log
loop = 4
welders_id = []
welders_list = []
welders_dob = []
welders_nationality = []
welders_pbpf = []
welders_slml = []

welders_process = []
welders_material = []
welders_thickness = []
welders_fwbw = []
welders_pipe_plate = []

#Data Load to Cashe
while True:
    cell_C = "C" + str(loop) #Welders credentials
    cell_E = "E" + str(loop) #Welders dob
    cell_D = "D" + str(loop) #Welders Nationality
    cell_G = "G" + str(loop) #Welders Process G
    cell_H = "H" + str(loop) #Welders Material H
    cell_I = "I" + str(loop) #Welders THICKNESS I
    cell_J = "J" + str(loop) #Welders PB/PF
    cell_K = "K" + str(loop) #Welders SL/ML
    cell_L = "L" + str(loop) #Welders FW/BW L
    cell_A = "A" + str(loop) #Welders ID-no
    cell_M = "M" + str(loop) #Welders Pipe / Pleate
    

    if sheet[cell_C].value == 'END':
        break

    welders_list.append(sheet[cell_C].value)
    welders_dob.append(sheet[cell_E].value)
    welders_nationality.append(sheet[cell_D].value)
    welders_pbpf.append(sheet[cell_J].value)
    welders_slml.append(sheet[cell_K].value)
    welders_process.append(sheet[cell_G].value)
    welders_material.append(sheet[cell_H].value)
    welders_thickness.append(sheet[cell_I].value)
    welders_fwbw.append(sheet[cell_L].value)
    welders_id.append(sheet[cell_A].value)
    welders_pipe_plate.append(sheet[cell_M].value)
    
    loop += 1


print('Data source read: \n')
print(welders_list)
print(welders_dob)
print(welders_pbpf)   
print(welders_slml)
print(welders_process)
print(welders_material)
print(welders_thickness)
print(welders_fwbw)
print(' \n --------- ------- ----- ------ ------ -----')


    #  ----------------------------------------------------------------------------

resurces_folder_name = sys.path[0] + "/Resources"
word_reports_folder = folder_location + "/" + "/Reports"

os.chdir(resurces_folder_name)

doc = MailMerge('RT.docx')


doc.merge(
    Job_no = job_no,
    Report_no = report_no,
    client = client_name,
    date = date_formated,
    word_client_adr = client_adr
)

welders_master_table = []

loop = 0
while True:
    if loop == len(welders_list):
        break

    welders_master_table.append(
    {
        'wqt_no': str(welders_id[loop]),
        'welder_name': str(welders_list[loop]),
        'sl_ml': str(welders_slml[loop]),
        'postion': str(welders_pbpf[loop]),
        'size': str(welders_thickness[loop]),
        'w_data' : str(welders_process[loop]) + "-" + str(welders_material[loop]),
        'comment': 'No defects noted',
        'result': 'Accepted'
    })
    loop += 1

doc.merge_rows('wqt_no',welders_master_table)

empty_picture_table = []


loop = 0
while True:
    if loop == len(welders_list):
        break

    weld_technique_data = " " + str(welders_process[loop]) + "-" + str(welders_material[loop]) + " " + welders_thickness[loop]
    insert_text = str(welders_id[loop]) + ". " + str(welders_list[loop]) + weld_technique_data + "\n Acceptable to Specification."
    
    empty_picture_table.append(
     {
        'picture_table': ''
    })
    
    empty_picture_table.append(
     {
        'picture_table': insert_text
    })

    loop += 1

doc.merge_rows('picture_table', empty_picture_table)

os.chdir(word_reports_folder)
doc.write(main_report_file_name) 
    
print(' \n --------- ------- ----- ------ ------ -----')
print('DONE')
print(' \n --------- ------- ----- ------ ------ -----')
os.system("pause") 