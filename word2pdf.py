# Coded by Gabriel O. González Rodríguez
# Github: gabrielgr01
# 26/06/2023

import sys
import os
import comtypes.client
import re

in_folder = input("Directory: ")
in_files = os.listdir(in_folder)

print ("")
print ("....... DOCX to PDF Converter .......")
for file_name_ext in in_files:
    file_name = re.findall(".*\." ,file_name_ext)
    file_name = file_name[0]
    file_name = file_name.strip(".")
    file_path = in_folder + file_name
    file_path_docx = file_path + ".docx"
    file_path_pdf = file_path + ".pdf"

    wdFormatPDF = 17

    in_file = os.path.abspath(file_path_docx)
    out_file = os.path.abspath(file_path_pdf)

    print ("")
    print ("Converting: ", in_file)
    print ("To:         ", out_file)

    word = comtypes.client.CreateObject('Word.Application')
    doc = word.Documents.Open(in_file)
    doc.SaveAs(out_file, FileFormat=wdFormatPDF)
    doc.Close()
    word.Quit()

print ("")
print (".....................................")
