import openpyxl
import os

excel_path = r'/Users/ryanco/Desktop/资源元表/服饰元表Excel.xlsx'
json_path = r'/Users/ryanco/PycharmProjects/AssetsManagement/Lagecy_Assets_List'

#get excel
wb = openpyxl.load_workbook(excel_path)
lws = wb['Lagecy_Assets_List']
nws = wb['New_Added_Assets_List']

