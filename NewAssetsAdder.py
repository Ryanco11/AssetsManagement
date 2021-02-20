import openpyxl
import os
import shutil
from pathlib import Path

excel_path = r'/Users/ryanco/Desktop/资源元表/服饰元表Excel.xlsx'
project_path = r'/Users/ryanco/Projects/AndoidProject/wonder_party/avatarProject/'
artsrc_path = r'/Users/ryanco/Projects/avatar_art_resources/'
asset_path = r'/Users/ryanco/Projects/AndoidProject/wonder_party/avatarProject/Assets/Art/model/coat/'
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


def ProcessAssetInfo(new_asset_list ,new_name_list):
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
                if asset.__contains__(r'/A/') or asset.__contains__(r'/FQ/'):
                    print("write in suit")
                    ws.cell(last_row, 4).value = namecode
                    ws.cell(last_row, 5).value = "suit"
                    ws.cell(last_row, 6).value = "普通套装"
                    MoveFiles(namecode, new_asset_list, "suit")
                    last_row += 1
                    break
                elif asset.__contains__(r'/K/') or asset.__contains__(r'/HQ/'):
                    print("write in pants")
                    ws.cell(last_row, 4).value = namecode
                    ws.cell(last_row, 5).value = "pants"
                    ws.cell(last_row, 6).value = "普通裤子"
                    MoveFiles(namecode, new_asset_list, "pants")

                    last_row += 1
                    break
                elif asset.__contains__(r'/T/'):
                    print("write in headwear")
                    ws.cell(last_row, 4).value = namecode
                    ws.cell(last_row, 5).value = "headwear"
                    ws.cell(last_row, 6).value = "普通头饰"
                    MoveFiles(namecode, new_asset_list, "headwear")
                    last_row += 1
                    break
                elif asset.__contains__(r'/F/'):
                    print("write in hair")
                    ws.cell(last_row, 4).value = namecode
                    ws.cell(last_row, 5).value = "hair"
                    ws.cell(last_row, 6).value = "普通头发"
                    MoveFiles(namecode, new_asset_list, "hair")
                    last_row += 1
                    break
                elif asset.__contains__(r'/M/'):
                    print("write in hair")
                    ws.cell(last_row, 4).value = namecode
                    ws.cell(last_row, 5).value = "hair"
                    ws.cell(last_row, 6).value = "帽子头发"
                    MoveFiles(namecode, new_asset_list, "hair")
                    last_row += 1
                    break
                elif asset.__contains__(r'/QT/'):
                    print("write in bladic")
                    ws.cell(last_row, 4).value = namecode
                    ws.cell(last_row, 5).value = "baldric"
                    ws.cell(last_row, 6).value = "普通背包"
                    MoveFiles(namecode, new_asset_list, "baldric")
                    last_row += 1
                    break
                # elif asset.__contains__(r'/W/'):
                #     print("write in sock")
                #     ws.cell(last_row, 4).value = namecode
                #     ws.cell(last_row, 5).value = "sock"
                #     ws.cell(last_row, 6).value = "普通袜子"
                #     last_row += 1
                #     break
                elif asset.__contains__(r'/X/'):
                    print("write in shoes")
                    ws.cell(last_row, 4).value = namecode
                    ws.cell(last_row, 5).value = "shoes"
                    ws.cell(last_row, 6).value = "普通鞋子"
                    MoveFiles(namecode, new_asset_list, "shoes")
                    last_row += 1
                    break
                elif asset.__contains__(r'/Y/'):
                    print("write in glasses")
                    ws.cell(last_row, 4).value = namecode
                    ws.cell(last_row, 5).value = "glasses"
                    ws.cell(last_row, 6).value = "普通眼镜"
                    MoveFiles(namecode, new_asset_list, "glasses")
                    last_row += 1
                    break
                elif asset.__contains__(r'/S/'):
                    print("write in shirt")
                    ws.cell(last_row, 4).value = namecode
                    ws.cell(last_row, 5).value = "shirt"
                    ws.cell(last_row, 6).value = "普通上衣"
                    MoveFiles(namecode, new_asset_list, "shirt")
                    last_row += 1
                    break


    wb.save("/Users/ryanco/Desktop/资源元表/服饰元表Excel.xlsx")

# def WriteSubType():



def MoveFiles(namecode, new_asset_list, dress_type):
    for asset in new_asset_list:
        if asset.__contains__(namecode):
            #create folder in unity model folder
            path_to_create = asset_path + dress_type + "/" + namecode
            if os.path.isdir(path_to_create):
                print("Exists")
            else:
                try:
                    os.mkdir(path_to_create)
                except OSError:
                    print("Creation of the directory %s failed" % path_to_create)
                else:
                    print("Successfully created the directory %s " % path_to_create)

            #copy file to new folder
            shutil.copy2(asset, path_to_create)  # target filename is /dst/dir/file.ext

            # shutil.copy2('/src/dir/file.ext', '/dst/dir/newname.ext')  # complete target filename given


#
# def WritePathInfo():

    # if new_asset_list.__contains__("_mask"):

# create

#detect dress type

#detect sub type
    ###for each name code###
    #copy to unity folder

    #write excel



# Start
new_asset_list, new_name_list = GetNewAssetInfo()
ProcessAssetInfo(new_asset_list, new_name_list)