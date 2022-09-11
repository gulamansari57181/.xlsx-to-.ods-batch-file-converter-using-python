import  jpype     
import  asposecells  
import glob,os,re



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




  

