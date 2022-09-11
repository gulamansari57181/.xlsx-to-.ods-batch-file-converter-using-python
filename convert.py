import  jpype     
import  asposecells  
import glob,os,re
import ezodf


def convert():
  jpype.startJVM() 
  from asposecells.api import Workbook ,SaveFormat


  source_directory = "C:/Users/Ansari/Desktop/excel-ods/input/"
  output_directory = "C:/Users/Ansari/Desktop/excel-ods/output/"

  files = os.listdir(source_directory)
  files_xlsx = [i for i in files if i.endswith('.xlsx')]

  subString=".xlsx"
  for i in files_xlsx:
    workbook = Workbook(source_directory + i)
    if i.endswith(subString):
      i = re.sub(subString, '',i)

    workbook.save( output_directory + i +'.ods', SaveFormat.ODS)
  jpype.shutdownJVM()

def delete():
  source_directory =  "C:/Users/Ansari/Desktop/excel-ods/output/"
  files = os.listdir(source_directory)
  files_ods = [i for i in files if i.endswith('.ods')]
  for i in files_ods:
    print(i)
    doc = ezodf.opendoc(source_directory + i)
    print(list(doc.sheets.names()))
    del doc.sheets['Evaluation Warning']
    print(list(doc.sheets.names()))
    doc.save()  

convert()
delete()
print("Task Completed")




  

