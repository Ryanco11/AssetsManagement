import openpyxl
import os
from pathlib import Path

excel_path = r'/Users/ryanco/Desktop/资源元表/服饰元表Excel.xlsx'
project_path = r'/Users/ryanco/Projects/AndoidProject/wonder_party/avatarProject/'
artsrc_path = r'/Users/ryanco/Projects/avatar_art_resources/'


namecode = "AA0269G"
dress_type = "shirt"
sub_type = "普通上衣"

def AddShirt(namecode, sub_type):
    asset_list = []

    for r, d, f in os.walk(artsrc_path):
        for file in f:
            # print(file)
            if file.lower().__contains__(namecode.lower()) and file.lower().endswith(".png"):
                asset_list.append(os.path.join(r, file).replace(project_path, ""))
            if file.lower().__contains__(namecode.lower()) and file.lower().endswith(".fbx"):
                asset_list.append(os.path.join(r, file).replace(project_path, ""))


    for i in asset_list:
        print(i)

AddShirt(namecode, sub_type)