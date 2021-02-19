import openpyxl
import os
from pathlib import Path

excel_path = r'/Users/ryanco/Desktop/资源元表/服饰元表Excel.xlsx'
project_path = r'/Users/ryanco/Projects/AndoidProject/wonder_party/avatarProject/'
artsrc_path = r'/Users/ryanco/Projects/avatar_art_resources/'

fbx_path = r'/Users/ryanco/Projects/avatar_art_resources/animation/bangding/'
png_path = r'/Users/ryanco/Projects/avatar_art_resources/model/change clothes_new/'

#get excel
wb = openpyxl.load_workbook(excel_path)
ws = wb['AssetsInfo - Assets_Art_model_c']

namecode = "AA0269G"
dress_type = "shirt"
sub_type = "普通上衣"

def AddShirt(namecode, dress_type, sub_type):
    asset_list = []

###fbx
    for r, d, f in os.walk(fbx_path):
        for file in f:
            # print(file)
            if file.lower().__contains__(namecode.lower()) and os.path.join(r, file).lower().__contains__("update") and file.lower().endswith(".fbx"):
                asset_list.append(os.path.join(r, file).replace(project_path, ""))

### png
    for r, d, f in os.walk(png_path):
        for file in f:
            # print(file)
            if file.lower().__contains__(namecode.lower()) and os.path.join(r, file).lower().__contains__("update") and file.lower().endswith(".png"):
                asset_list.append(os.path.join(r, file).replace(project_path, ""))

    for i in asset_list:
        print(i)

def GetLastRow():
    # last_row = 0
    for row in range(2, ws.max_row):
        # print(ws.cell(row, 1).value)
        if(ws.cell(row, 1).value is None):
            last_row = row;
            # print(last_row)
            break
    return last_row





###function start###

last_row = GetLastRow()

#loop from last to upper
for row in last_row:
    cur_row = last_row - row
    if ws.cell(cur_row, 7).value is None:
        ###for every new asset

        #get source files
        AddShirt(namecode, dress_type, sub_type)

        #

    #value is not null, end operation
    else:
        break