import openpyxl
import os

excel_path = r'/Users/ryanco/Desktop/资源元表/服饰元表Excel.xlsx'
project_art_path = r'/Users/ryanco/Projects/AndoidProject/wonder_party/avatarProject/Assets/Art'
project_path = r'/Users/ryanco/Projects/AndoidProject/wonder_party/avatarProject'

#get excel
wb = openpyxl.load_workbook(excel_path)
lws = wb['Lagecy_Assets_List']
nws = wb['New_Added_Assets_List']

