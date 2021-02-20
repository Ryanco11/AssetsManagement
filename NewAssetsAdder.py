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


def GetNewAssetInfo():

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
        print("all_assets: " + i)

    #get new added asset list
        #get last row from new excel
    last_row = GetLastRow()

    excel_namecode_list = []
    new_name_list = []
    mt_namecode_list = []
        #get namecode list from excel
    for row in range(2, last_row):
        excel_namecode_list.append(ws.cell(row, 4).value)
        print("ws: " + ws.cell(row, 4).value)

        #remove existing asset
    for asset in asset_list:
        for namecode in excel_namecode_list:
            if asset.__contains__(namecode):
                asset_list.remove(asset)
                break

    for i in asset_list:
        print("new_assets: " + i)
    new_asset_list = asset_list

    #sort by name code
        #get new name code list
    for asset in asset_list:
        mt_namecode_list.append(asset.split(r'/')[-2])
    print(mt_namecode_list)

        #remove mt name code
    for i in mt_namecode_list:
        if i not in new_name_list:
            new_name_list.append(i)

    print(new_name_list)

    return new_asset_list, new_name_list


def GetDressType(new_asset_list ,new_name_list):
    last_row = GetLastRow()
    # print(last_row)
    # for each name code
    for namecode in new_name_list:
        print("t-nc:" + namecode)
        for asset in new_asset_list:
            print("t-asset:" + asset)
            if asset.__contains__(namecode):
                #for specfic file path with its namecode
                #write type
                if asset.__contains__(r'/X/'):
                    print("write in")
                    ws.cell(last_row, 4).value = namecode
                    ws.cell(last_row, 5).value = "shoes"
                    last_row += 1
                    break
                #move file
                #write rest info

    wb.save("/Users/ryanco/Desktop/资源元表/服饰元表Excel.xlsx")



    # if new_asset_list.__contains__("_mask"):

# create

#detect dress type

#detect sub type
    ###for each name code###
    #copy to unity folder

    #write excel



# Start
new_asset_list, new_name_list = GetNewAssetInfo()
GetDressType(new_asset_list, new_name_list)