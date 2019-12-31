# To add a new cell, type '# %%'
# To add a new markdown cell, type '# %% [markdown]'
# %%
import pandas as pd
import docx
import csv
import os
import shutil


# %%
import os

path = 'c:\\projects\\hc2\\'
log_path ='c:\\projects\\log\\'

files = []
# r=root, d=directories, f = files
for r, d, f in os.walk(path):
    for file in f:
        if '.xlsx' in file:
            files.append(os.path.join(r, file))

for f in files:
    print(f)


# %%
ExcelFile= f
csvFile= ExcelFile[:-4]+'csv'
docFile= ExcelFile[:-4]+'docx'
Token ='TokenFile.docx'


# %%
def CopyFile(source,destination):
    try: 
        shutil.copyfile(source, destination) 
        print("File copied successfully.") 

    # If source and destination are same 
    except shutil.SameFileError: 
        print("Source and destination represents the same file.") 

    # If destination is a directory. 
    except IsADirectoryError: 
        print("Destination is a directory.") 

    # If there is any permission issue 
    except PermissionError: 
        print("Permission denied.") 

    # For other errors 
    except: 
        print("Error occurred while copying file.") 


# %%
df = pd.read_excel(ExcelFile, sheet_name='Page1_1',header= 1)


# %%
df.to_csv(csvFile,sep= '\t')


# %%
f = open(csvFile,"r")
s=f.read()
f.close()


# %%
CopyFile(Token,docFile)


# %%
doc=docx.Document(docFile)
doc.add_paragraph(s)
doc.save(docFile)


# %%
shutil.move(ExcelFile,log_path)


# %%
#os.remove(docFile)
#os.remove(csvFile)

# %% [markdown]
# import sys
# !{sys.executable} -m pip install "D:\TA_Lib-0.4.17-cp36-cp36m-win_amd64.whl"

# %%



# %%



# %%



# %%


