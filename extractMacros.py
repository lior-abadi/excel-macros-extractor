import win32com.client as win32
import pandas as pd
import os

directory = r"INPUT_COMPLETE_PATH" ## E.G "C:\Users\User\Desktop"
convertedPath = r"OUTPUT_COMPLETE_PAT" ## E.G "C:\Users\User\Desktop\converted/"


listDirectory = os.listdir(directory)
print(listDirectory)

print("Collecting and extracting macros from the following files...")
for filename in listDirectory:
    if (filename[-2:] == "py" or filename[:2] == "~$"):
        continue
    
    if (filename == "converted"):
        continue
    
    f = os.path.join(directory, filename)
    if os.path.isfile(f):
        print(f)
     
    excel = win32.Dispatch("Excel.Application")
    workbook = excel.Workbooks.Open(f)
    dict_modules = {}
    for i in workbook.VBProject.VBComponents:
        name = i.name
        lines = workbook.VBProject.VBComponents(name).CodeModule.CountOfLines

        # To jump empty modules
        if lines == 0:
            pass
        else:
            text = workbook.VBProject.VBComponents(name).CodeModule.Lines(1,lines)
            dict_modules[name] = [text]

    df = pd.DataFrame(dict_modules)
    
    finalPath = convertedPath  + filename[:-5] + ".bas"
    print(finalPath)
    df.to_csv(finalPath, header=None, index=None, sep=' ', mode='a')


