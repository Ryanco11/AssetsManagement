import openpyxl
import os
from pathlib import Path

excel_path = r'/Users/ryanco/Desktop/资源元表/服饰元表Excel.xlsx'
project_path = r'/Users/ryanco/Projects/AndoidProject/wonder_party/avatarProject/'
artsrc_path = r'/Users/ryanco/Projects/avatar_art_resources/'

fbx_path = r'/Users/ryanco/Projects/avatar_art_resources/Dress/'
png_path = r'/Users/ryanco/Projects/avatar_art_resources/Dress/'

#get excel
wb = openpyxl.load_workbook(excel_path)
ws = wb['NewAssets - Assets_Art_model_co']

# namecode = "AA0269G"
# dress_type = "shirt"
# sub_type = "普通上衣"

def GetAssetList():
    fbx_list = []
    png_list = []

###fbx
    for r, d, f in os.walk(fbx_path):
        for file in f:
            # print(file)
            if file.lower().endswith(".fbx"):
                fbx_list.append(os.path.join(r, file).replace(project_path, ""))

### png
    for r, d, f in os.walk(png_path):
        for file in f:
            # print(file)
            if file.lower().endswith(".png"):
                png_list.append(os.path.join(r, file).replace(project_path, ""))

    return fbx_list, png_list

def GetLastRow():
    # last_row = 0
    for row in range(2, ws.max_row):
        # print(ws.cell(row, 1).value)
        if(ws.cell(row, 4).value is None):
            last_row = row;
            print("last_row: " + str(last_row))
            break

    return last_row


# def GetDressType(asset_list):
#     for asset in asset_list:
#         if asset.__contains__("/A/"):


    # if asset_list.__contains__("_mask"):





###function start###


#get all asset in art path
# get source files [0]:fbx [1]:png
sperate_asset_list = GetAssetList()
fbx_list = sperate_asset_list[0]
png_list = sperate_asset_list[1]

asset_list = []

for i in fbx_list:
    # print("fbx:" + i)
    asset_list.append(i)
for i in png_list:
    # print("png:" + i)
    asset_list.append(i)

for i in asset_list:
    print("assets: " + i)

#get new added asset list
    #get last row from new excel
last_row = GetLastRow()

namecode_list = []
mt_namecode_list = []
    #get namecode list from excel
for row in range(2, last_row):
    namecode_list.append(ws.cell(row, 4).value)
    print("ws: " + ws.cell(row, 4).value)

    #remove existing asset
for asset in asset_list:
    for namecode in namecode_list:
        if asset.__contains__(namecode):
            asset_list.remove(asset)
            break

#sort by name code
    #get new name code list
for asset in asset_list:
    mt_namecode_list.append(asset.split(r'/')[-2])

for i in mt_namecode_list:
    if i not in namecode_list:
        namecode_list.append(i)

print(namecode_list)


    #create

#detect dress type

#detect sub type
    ###for each name code###
    #copy to unity folder

    #write excel

