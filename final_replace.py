import pandas as pd
import zipfile
import os

#pass in excel file path, the key and the value for columns
dic = pd.read_excel('./new_replacement.xlsx', engine='openpyxl').set_index('Key')['Value'].to_dict()
color_change_needed = set()
def replace(old_path,new_path, dic):

    zin = zipfile.ZipFile (old_path, 'r')
    zout = zipfile.ZipFile (new_path, 'w')

    for i in zin.infolist():
        buffer = zin.read(i.filename)
        if (i.filename == 'word/document.xml'):
            result = buffer.decode("utf-8")
            for key in dic:
                result = result.replace(key,dic[key])
                if '**' in result:
                    color_change_needed.add(old_path)
            buffer = result.encode("utf-8")
        zout.writestr(i, buffer)
    zout.close()
    zin.close()

#directory of the input and output folder containing the files that need to be replaced.
directory_input = r'C:\Users\AnushPoudel\Documents\Input'
directory_output = r'C:\Users\AnushPoudel\Documents\Output'

for entry in os.scandir(directory_input):
    if (entry.path.endswith(".docx") or entry.path.endswith(".doc")) and entry.is_file():
       replace(entry.path, directory_output + '\\' + entry.name, dic)

print(color_change_needed)