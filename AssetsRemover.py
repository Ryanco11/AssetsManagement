import openpyxl
import os
from pathlib import Path
from openpyxl.styles import Color, PatternFill, Font, Border


excel_path = r'/Users/ryanco/Desktop/资源元表/服饰元表Excel.xlsx'
project_path = r'/Users/ryanco/Projects/AndoidProject/wonder_party/avatarProject/'
prefab_path = r'/Users/ryanco/Projects/AndoidProject/wonder_party/avatarProject/Assets/Art/BundleResources/Dress'
assets_path = r'/Users/ryanco/Projects/AndoidProject/wonder_party/avatarProject/Assets/Art/model/coat'
sprite_path = r'/Users/ryanco/Projects/AndoidProject/wonder_party/avatarProject/Assets/Art/BundleResources/Sprites'


wb = openpyxl.load_workbook(excel_path)
nws = wb['New_Added_Assets_List']

def GetLastRow(ws):
    for row in range(1, ws.max_row + 100000):
        if (ws.cell(row, 3).value is None): #check time stamp for last row
            last_row = row;  # this last_row value is actual plus one by actual last row in excel, cus py access col by minis one
            print("last row is : " + str(last_row))
            return last_row

def CheckAssetToRemove(nws):
    #loop whole excel
    for row in range(2, last_row):  # start at 2 , cus first row is not the actual info
        file_col_list = [7, 10, 12, 14, 16]
        if nws.cell(row, 4).value == "":
            for col in file_col_list:
                if not nws.cell(row, col).value == "0" and not nws.cell(row, col).value.__contains__("旧资源"):
                    #detele file in unity
                    DeleteCell(nws.cell(row, col).value)

            #detele whole row

def DeleteCell(cell_value):
    cell_path_list = cell_value.split('|-|')
    for path in cell_path_list:
        if len(path) < 5:
            # 绕过序号
            continue
        #delete every file
        




last_row = GetLastRow(nws)
CheckAssetToRemove()