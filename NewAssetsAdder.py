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

def GetAssetList(namecode, dress_type, sub_type):
    fbx_list = []
    png_list = []

###fbx
    for r, d, f in os.walk(fbx_path):
        for file in f:
            # print(file)
            if file.lower().__contains__(namecode.lower()) and os.path.join(r, file).lower().__contains__("update") and file.lower().endswith(".fbx"):
                fbx_list.append(os.path.join(r, file).replace(project_path, ""))

### png
    for r, d, f in os.walk(png_path):
        for file in f:
            # print(file)
            if file.lower().__contains__(namecode.lower()) and os.path.join(r, file).lower().__contains__("update") and file.lower().endswith(".png"):
                png_list.append(os.path.join(r, file).replace(project_path, ""))

    return fbx_list, png_list

def GetLastRow():
    # last_row = 0
    for row in range(2, ws.max_row):
        # print(ws.cell(row, 1).value)
        if(ws.cell(row, 4).value is None):
            last_row = row;
            print(last_row)
            break

    return last_row


# def GetDressType(asset_list):
#     for asset in asset_list:
#         if asset.__contains__("/A/"):


    # if asset_list.__contains__("_mask"):






###function start###
last_row = GetLastRow()

#get all asset in art path

#get new added asset list

#sort by name code
    ###for each name code###
    #copy to unity folder

    #write excel

#loop from last to upper
for row in range(1, last_row):
    cur_row = last_row - row
    if ws.cell(cur_row, 7).value is None:
        ###for every new asset

        #get source files [0]:fbx [1]:png
        asset_list = GetAssetList(namecode, dress_type, sub_type)
        fbx_list = asset_list[0]
        png_list = asset_list[1]

        #copy asset
        for i in fbx_list:
            print("fbx:" + i)
        for i in png_list:
            print("png:" + i)
