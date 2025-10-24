import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Border

inv_wb = openpyxl.load_workbook("/mnt/c/Users/Bart/Desktop/Harmonized Chapters/INVDAY_master.xlsx")
inv_ws = inv_wb[ 'Sheet1']

metal_master_wb = openpyxl.load_workbook("/mnt/c/Users/Bart/Desktop/Harmonized Chapters/metal_content_master.xlsx")
metal_master_ws = metal_master_wb['Sheet1']

final_wb = Workbook()
final_ws = final_wb.active

steel_codes = "/mnt/c/Users/Bart/Desktop/Harmonized Chapters/steelHTSlist_justnumbers.txt"
alum_codes = "/mnt/c/Users/Bart/Desktop/Harmonized Chapters/aluminumHTSlist_justnumbers.txt"

has_data = True     # while loop used to find the number of rows in the inv sheet that contain data.
inv_row_count = 0

while has_data:
    inv_row_count += 1
    data = inv_ws.cell(row=inv_row_count, column=1).value
    if data == None:
        has_data = False

has_data = True     # while loop used to find the number of rows in the metal_master sheet that contain data.
metal_master_row_count = 0

while has_data:
    metal_master_row_count += 1
    data = metal_master_ws.cell(row=metal_master_row_count, column=9).value
    if data == None:
        has_data = False

def harm_codes(fp):        # returns the contents of a text file as a list of strings from a file path
    codes = []
    with open(fp) as f:
        file_contents = f.read()    
    for content in file_contents.split():
        codes.append(content)
    return codes

def find_declaration_req(codes):       # finds the codes in workbook that require steel declaration.
    declaration_req = []
    for i in range(1, inv_row_count):    
        for hc in codes:
            value = inv_ws.cell(row=i, column=6).value
            if value[:len(hc)] == hc:
                declaration_req.append(inv_ws.cell(row=i, column=6))
                break
    return declaration_req

def declared_sku(harm_cell):        # finds sku associated with harm requiring declaration.
    skus = []
    for cell in harm_cell:
        skus.append(inv_ws.cell(row=cell.row, column=3))
    return skus

steel_sku = declared_sku(find_declaration_req(harm_codes(steel_codes)))
alum_sku = declared_sku(find_declaration_req(harm_codes(alum_codes)))

final_ws_tracker = 0

for sku in steel_sku:
    for i in range(1, metal_master_row_count):
        value = metal_master_ws.cell(row=i, column=3).value
        if sku.value == value:
            metal_master_range = metal_master_ws[f"A{i-1}" : f"H{i+2}"]
            final_range = final_ws[f"A{1+final_ws_tracker}" : f"H{4+final_ws_tracker}"]
            metal_master_cells = [val for row in metal_master_range for val in row]
            final_cells = [val for row in final_range for val in row]
            for i in range(32):     # 32 is the 4x8 length/area of the ranges.
                final_ws[final_cells[i].coordinate] = metal_master_ws[metal_master_cells[i].coordinate].value
            final_ws_tracker += 5



final_wb.save("/mnt/c/Users/Bart/Desktop/Harmonized Chapters/final_test.xlsx")