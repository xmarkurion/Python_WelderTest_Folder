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
        ' ---------------Reqest Data Engine---------------------   \n'                                                                                                                                                
    )

def make_space():
    os.system('cls')
    logo()


#req_folder_location = input('Location of the REQ File: ')
make_space()

print('Please enter folder location where test_data.ini is: ')
folder_location = input('Enter it here: ')



#folder_location = r'E:\temp\22Jan19'
os.chdir(folder_location)

#Read config file
config = configparser.ConfigParser()
config.read('test_data.ini')

client_name = config['TEST']['client']
date = config['TEST']['date']
job_no = config['TEST']['job_no']
test_type = config['TEST']['test_type']
report_no = config['TEST']['report_no']
date_formated = config['TEST']['date_formated']
main_report_file_name = config['TEST']['report_name']

xls_file_name = "REQ " + date.upper() + " " + client_name + ".xlsx"
print(xls_file_name)

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

logo()

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


#Paste data module 
amout_of_records = loop - 4

certs_folder = folder_location + "/" + "Certs"
os.chdir(certs_folder)

wb = openpyxl.load_workbook('cert_gen.xlsx')
sheet = wb['Welders']

#print( sheet['F2'].value )

loop = 2
cert_loop = 1
while True:
    cell_A = "A" + str(loop) #Welder ID
    cell_F = "F" + str(loop) #Welder Name
    cell_G = "G" + str(loop) #Welder DOB
    cell_H = "H" + str(loop) #Welder Nationality
    cell_I = "I" + str(loop) #Welder postion PB/PF/PA etc.
    cell_J = "J" + str(loop) #Welder ragnge for PB/PA
    cell_K = "K" + str(loop) #Welder sl/ml
    cell_L = "L" + str(loop) #Welder sl/ml range
    cell_N = "N" + str(loop) #Welder mm size
    cell_V = "V" + str(loop) #Welders Welding Type
    cell_W = "W" + str(loop) #Welders Welding Material Type
    cell_X = "X" + str(loop) #Welders pipe / plate

    if loop == amout_of_records + 1:
        break

    sheet[cell_A].value = welders_id[cert_loop]
    sheet[cell_F].value = welders_list[cert_loop]
    sheet[cell_G].value = welders_dob[cert_loop]
    sheet[cell_H].value = welders_nationality[cert_loop]

    sheet[cell_I].value = welders_pbpf[cert_loop]
    if welders_pbpf[cert_loop] == "PA":
        sheet[cell_J].value = "PA"
    if welders_pbpf[cert_loop] == "PB":
        sheet[cell_J].value = "PA, PB"
    if welders_pbpf[cert_loop] == "PF":
        sheet[cell_J].value = "PF, PB, PA"

    sheet[cell_K].value = welders_slml[cert_loop]
    if welders_slml[cert_loop] == "SL":
        sheet[cell_L].value = "SL"
    if welders_slml[cert_loop] == "ML":
        sheet[cell_L].value = "ML, SL"
    if welders_slml[cert_loop] == "SS, NB":
        sheet[cell_L].value = "SS NB, SS MB, BS, SS GB"

    sheet[cell_N].value = welders_thickness[cert_loop]
    sheet[cell_V].value = welders_process[cert_loop]
    sheet[cell_W].value = welders_material[cert_loop]
    sheet[cell_X].value = welders_pipe_plate[cert_loop]
    

    cert_loop += 1
    loop += 1

wb.save('cert_gen.xlsx')

resurces_folder_name = sys.path[0] + "/Resources"
word_reports_folder = folder_location + "/" + "/Reports"

os.chdir(resurces_folder_name)

if int(test_type) == 1:
    doc = MailMerge('Fracture.docx')

if int(test_type) == 2:
    doc = MailMerge('Macro.docx')

doc.merge(
    Job_no = job_no,
    Report_no = report_no,
    client = client_name,
    date = date_formated
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
        'comment': 'No defects noted',
        'result': 'Accepted'
    })
    loop += 1

print(welders_master_table)

doc.merge_rows('wqt_no',welders_master_table)

empty_picture_table = []

loop = 0
while True:
    if loop == len(welders_list):
        break

    
    #If Macro is 
    if int(test_type) == 2:
        insert_text_macro = str(welders_id[loop]) + ". " + str(welders_list[loop]) + "\n Acceptable to Specification."
        empty_picture_table.append(
        {
            'picture_table_1': insert_text_macro,
            'picture_table_2': insert_text_macro
        })
    
    #If Fracture is
    if int(test_type) == 1:
        insert_text_fracture_1 = str(welders_id[loop]) + ". " + str(welders_list[loop]) + "\n Acceptable to Specification."
        
        if (loop < len(welders_list)+1) :
            if (len(welders_list)-1 == loop):
                break
            insert_text_fracture_2 = str(welders_id[loop+1]) + ". " + str(welders_list[loop+1]) + "\n Acceptable to Specification."

            empty_picture_table.append(
            {
            'picture_table_1': insert_text_fracture_1,
            'picture_table_2': insert_text_fracture_2
            })

            empty_picture_table.append(
            {
                'picture_table_1': '\n \n \n \n \n \n \n \n \n \n \n \n \n \n \n \n \n \n \n \n',
                'picture_table_2': '\n \n \n \n \n \n \n \n \n \n \n \n \n \n \n \n \n \n \n \n'
            })

            loop += 1

    loop += 1

doc.merge_rows('picture_table_1', empty_picture_table)

os.chdir(word_reports_folder)
doc.write(main_report_file_name) 
    

os.system("pause") 