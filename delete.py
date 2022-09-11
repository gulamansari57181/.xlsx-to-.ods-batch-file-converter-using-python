import ezodf
import os ,glob,re

source_directory =  "C:/Users/Ansari/Desktop/excel-ods/output/"
# output_directory = "C:/Users/Ansari/Desktop/delete_sheets/after-delete-worksheet/"

files = os.listdir(source_directory)
files_ods = [i for i in files if i.endswith('.ods')]
for i in files_ods:
    print(i)
    doc = ezodf.opendoc(source_directory + i)
    print(list(doc.sheets.names()))
    del doc.sheets['Evaluation Warning']
    print(list(doc.sheets.names()))
    doc.save()
   
exit()


