import openpyxl

excel_path = r'/Users/ryanco/Desktop/资源元表/服饰元表Excel.xlsx'
project_path = r'/Users/ryanco/Projects/AndoidProject/wonder_party/avatarProject/'

#get excel
wb = openpyxl.load_workbook(excel_path)
lws = wb['Lagecy_Assets_List']

def AccessAssetPath(last_row):
    for row in range(2, last_row):   # start at 2 , cus first row is not the actual info
        # namecode
        print("this is asset: " +  lws.cell(row, 2).value)

        ### Assets
        file_col_list = [7, 10, 12, 14, 16]

        for col in file_col_list:
            print("asset " + str(col) + ": " + lws.cell(row, col).value)

        # #Prefab
        # print("pfb: " + ws.cell(row, 7).value)
        # # Tex
        # print("tex: " + ws.cell(row, 10).value)
        # # Mesh
        # print("mash: " + ws.cell(row, 12).value)
        # # Mat
        # print("mat: " + ws.cell(row, 14).value)
        # # Sprite
        # print("sprs: " + ws.cell(row, 16).value)

def AccessAssetSpecificSetting(last_row, ws):
    for row in range(2, last_row):
        # namecode
        print("this is asset: " + ws.cell(row, 2).value)

def GetLastRow(ws):
    for row in range(1, ws.max_row):
        if (ws.cell(row, 4).value is None):
            last_row = row;  # this last_row value is actual plus one by actual last row in excel, cus py access col by minis one
            print("last row is : " + str(last_row))
            return last_row

# def CheckAssetLosing():







###Start
#1. Check Lagecy Assets
last_row = GetLastRow(lws)
# AccessAssetPath(last_row, lws)


#2. Cehck New Added Asset Since 2021-02-21









###########          Tips              ###########

#(row, col)

# loop row  # start at 2 , cus first row is not the actual info
# for row in range(2, 7):
#     for col in range(2, 3):  # (2, 3) only access to number 2
#         print(ws.cell(row, col).value)
