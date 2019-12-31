# To add a new cell, type '# %%'
# To add a new markdown cell, type '# %% [markdown]'
# %%
import pandas as pd
import docx
import csv
import os


# %%
import os

path = 'c:\\projects\\AdmanGo\\'

files = []

# r=root, d=directories, f = files
for r, d, f in os.walk(path):
    for file in f:
        if '.docx' in file:
            files.append(os.path.join(r, file))

for f in files:
    print(f)


# %%
WordFile=f
csvFile=f[0:-4]+"csv"
csvFile


# %%
doc=docx.Document(WordFile)


# %%
file = open(csvFile,"w") 
for p in doc.paragraphs:
    file.write(p.text)
file.close()


# %%


