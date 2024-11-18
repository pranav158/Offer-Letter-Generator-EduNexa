#              $$\      $$\                       $$\     $$\           
#              $$$\    $$$ |                      $$ |    \__|          
#              $$$$\  $$$$ |$$\   $$\  $$$$$$$\ $$$$$$\   $$\  $$$$$$$\ 
#              $$\$$\$$ $$ |$$ |  $$ |$$  _____|\_$$  _|  $$ |$$  _____|
#              $$ \$$$  $$ |$$ |  $$ |\$$$$$$\    $$ |    $$ |$$ /      
#              $$ |\$  /$$ |$$ |  $$ | \____$$\   $$ |$$\ $$ |$$ |      
#              $$ | \_/ $$ |\$$$$$$$ |$$$$$$$  |  \$$$$  |$$ |\$$$$$$$\ 
#              \__|     \__| \____$$ |\_______/    \____/ \__| \_______|
#                           $$\   $$ |                                  
#                           \$$$$$$  |                                  
#                            \______/               
# 

import pyautogui
import pandas
import datetime
import time
from docx import Document
import os

# Read data from excel
excel_data = pandas.read_excel('data.xlsx', sheet_name='Recipient Details')
count = 0
directory = 'generated letters'

def replaceWord(oldString, newString, paragraph):
    if oldString in paragraph:
        inline = p.runs
        # Loop added to work with runs (strings with same style)
        for i in range(len(inline)):
            if oldString in inline[i].text:
                text = inline[i].text.replace(oldString, newString)
                inline[i].text = text

# Iterate excel rows till to finish
for column in excel_data['CERT_ID'].tolist():
    document = Document('19jan.docx')
    doc = document
    empName = excel_data['CERT_ID'][count]
    empName1 = excel_data['First Name'][count]
    for p in doc.paragraphs:
        replaceWord('NAME', excel_data['NAME'][count], p.text)
        replaceWord('CERT_ID', excel_data['CERT_ID'][count], p.text)
        replaceWord('DOMAIN', excel_data['Domain'][count], p.text)

    try:
        path = os.getcwd()+"/"+directory
        os.mkdir(path)
    except OSError:
        a = 10
    doc.save(os.path.join(os.getcwd(), directory, f'{empName}_{empName1}.docx'))
    #doc.save(os.getcwd()+"/"+directory+"/"+empName+"/"+{empName}_{empName1}+ '.docx')
    print("Letter generated for " + empName)
    count = count + 1
    
print("""
              $$\      $$\                       $$\     $$\            
              $$$\    $$$ |                      $$ |    \__|           
              $$$$\  $$$$ |$$\   $$\  $$$$$$$\ $$$$$$\   $$\  $$$$$$$\  
              $$\$$\$$ $$ |$$ |  $$ |$$  _____|\_$$  _|  $$ |$$  _____| 
              $$ \$$$  $$ |$$ |  $$ |\$$$$$$\    $$ |    $$ |$$ /       
              $$ |\$  /$$ |$$ |  $$ | \____$$\   $$ |$$\ $$ |$$ |       
              $$ | \_/ $$ |\$$$$$$$ |$$$$$$$  |  \$$$$  |$$ |\$$$$$$$\  
              \__|     \__| \____$$ |\_______/    \____/ \__| \_______| 
                           $$\   $$ |                                   
                           \$$$$$$  |                                   
                            \______/                
""")
print("Total letters are created " + str(count))
input("Press Enter to continue...")
