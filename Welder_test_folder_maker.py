import os
import shutil
import glob
import datetime
import sys
from time import gmtime, strftime

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
        ' ------------------------App for sorting images--------------------------   \n'                                                                                                                                                
    )

def line():
    print('\n  -------------------------------------------------------------------   \n')

def create_folder_structure(type):
    base_foler_name = strftime("%d%b%y", gmtime())
    
    resurces_folder_name = sys.path[0] + "/Resources"
    excel_req_file_name = resurces_folder_name + "/REQ.xls"
    excel_certgen_file_name = resurces_folder_name + "/cert_gen.xls"
    word_fracture_file_name = resurces_folder_name + "/Fracture.docx"

    if not (os.path.isdir(base_foler_name)):
        os.mkdir(base_foler_name)
    
    #Folder Name redirect's
    inside_base_folder_name = sys.path[0] + "/" + base_foler_name
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
    if not(os.path.isfile("REQ.xls")):
        shutil.copyfile(excel_req_file_name,"REQ.xls")

    #Here Script copy the Cert_gen.xls from resources file.
    os.chdir(inside_cers_folder_name)
    if not(os.path.isfile("cert_gen.xls")):
        shutil.copyfile(excel_certgen_file_name,"cert_gen.xls")

    #Here Script copy Reports into reperts acording to choice.
    #os.chdir(inside_reports_folder_name)
    #if not(os.path.isfile("Fracture.docx")
    #    shutil.copyfile(word_fracture_file_name,"Fracture.docx")
   
    

#----- FUNC BLOCK ENDS

# - 1 For Fracture  - 2 For Macro
logo()

print('Select test type: \n 1.Fracture \n 2.Macro \n')
test_type = input('Enter choice: ')

create_folder_structure(int(test_type))
os.system("pause")