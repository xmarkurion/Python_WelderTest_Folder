import os
import shutil
import glob
import datetime
import sys
from time import gmtime, strftime
import openpyxl

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
        ' ----------------Welder Test Folder Structure Editor---------------------   \n'                                                                                                                                                
    )

def make_space():
    os.system('cls')
    logo()


#Przerobic all xls na xlsx i poprawic obrazki oraz sprawdzić gdzie zapisują się pewne dane.
def excel_cert_data_editor(excel_file, folder, date, client_name, test_type):
    work_folder = folder + "/" + date + "/Certs"

    os.chdir(work_folder)
    wb = openpyxl.load_workbook(excel_file)
    print(wb)

def excel_req_data_editor(excel_file, folder, date, client_name, job_no, report_no, test_type):
    
    time_string_gen = date
    work_folder = folder + "/" + time_string_gen

    os.chdir(work_folder)   #Select work Folder
    wb = openpyxl.load_workbook(excel_file)  #open excel file
    
    sheet = wb['Sheet1'] # open sheet for reading or editing

    if(test_type == 1):
        sheet['D1'].value = "FRACTURE WELDING SURVEYOR'S REPORT"
    if(test_type == 2):
        sheet['D1'].value = "MACRO WELDING SURVEYOR'S REPORT"

    sheet['D2'].value = client_name #making changes
    sheet['J24'].value = job_no
    sheet['J26'].value = report_no
    wb.save(excel_file)
    

def create_folder_structure(type,folder,date):
    
    resurces_folder_name = sys.path[0] + "/Resources"
    excel_req_file_name = resurces_folder_name + "/REQ.xlsx"
    excel_certgen_file_name = resurces_folder_name + "/cert_gen.xlsx"
    word_fracture_file_name = resurces_folder_name + "/Fracture.docx"

    #Genereate the folder name as exammple 13Jun19
    time_string_gen = date

    #Open the path created by User. 
    os.chdir(folder)

    if not (os.path.isdir(time_string_gen)):
        os.mkdir(time_string_gen)
    
    #Folder Name redirect's
    base_foler_name = folder + "/" + time_string_gen
    inside_base_folder_name = base_foler_name
    inside_cers_folder_name = inside_base_folder_name + "/Certs"
    inside_reports_folder_name = inside_base_folder_name + "/Reports"

    os.chdir(inside_base_folder_name)
    if not(os.path.isdir("Certs")):
        os.mkdir("Certs")

    if not(os.path.isdir("Reports")):
        os.mkdir("Reports")

    if not(os.path.isdir("Welders")):
        os.mkdir("Welders")

    if(type == 1):
        if not(os.path.isdir("Fracture")):
            os.mkdir("Fracture")

    if(type == 2):
        if not(os.path.isdir("Macro")):
            os.mkdir("Macro")
        
    #Here Script copy the REQ.xls from resources file.
    if not(os.path.isfile("REQ.xlsx")):
        shutil.copyfile(excel_req_file_name,"REQ.xlsx")

    #Here Script copy the Cert_gen.xls from resources file.
    os.chdir(inside_cers_folder_name)
    if not(os.path.isfile("cert_gen.xlsx")):
        shutil.copyfile(excel_certgen_file_name,"cert_gen.xlsx")

    #Here Script copy Reports into reperts acording to choice.
    #os.chdir(inside_reports_folder_name)
    #if not(os.path.isfile("Fracture.docx")
    #    shutil.copyfile(word_fracture_file_name,"Fracture.docx")
   
    

#----- FUNC BLOCK ENDS

# - 1 For Fracture  - 2 For Macro
logo()

#folder_name = input('Please enter folder path: ')
folder_name = "C:/Users/M/Downloads/test"

print('Select test type: \n 1.Fracture \n 2.Macro \n')
test_type = input('Enter choice: ')

make_space()

print('Do you want to use today date for folder & reports: ' + strftime("%d%b%y", gmtime()) + '\n 1. For Yes \n 2. For NO.' )
folder_date = input('\n Enter choice: ')

if(int(folder_date) == 2):
    make_space()
    print('Your own date in format DayMonthYear ')
    folder_custom_date = input('Please enter the date: ')

make_space()

client_name = input('Enter client name: ')

make_space()

print('Creating folder structure in: ' + folder_name)

if(int(folder_date) == 1):
    now_date = strftime("%d%b%y", gmtime())
    create_folder_structure(int(test_type),folder_name,now_date)
    excel_req_data_editor('REQ.xlsx',folder_name, now_date, client_name,'1234','23004',int(test_type))
    excel_cert_data_editor('cert_gen.xlsx',folder_name,now_date,client_name,test_type)
    

if(int(folder_date) == 2):
    create_folder_structure(int(test_type),folder_name,folder_custom_date)
    excel_req_data_editor('REQ.xlsx',folder_name, folder_custom_date, client_name,'1234','23004',int(test_type))
    

print(" ---- ^^_^^ -----")

print('Editing excel client data name & job & report')


print(" ---- ^^_^^ -----")
os.system("pause")
