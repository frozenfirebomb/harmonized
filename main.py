import openpyxl
from openpyxl import workbook

from functions import harm_codes

inv_wb = openpyxl.load_workbook("/mnt/c/Users/Bart/Desktop/Harmonized Chapters/INVDAY_master.xlsx")
inv_ws = inv_wb[ 'Sheet1']

metal_master_wb = openpyxl.load_workbook("/mnt/c/Users/Bart/Desktop/Harmonized Chapters/metal_content_master.xlsx")
# metal_master_ws = metal_master_wb['Sheet1']

steel_codes = "/mnt/c/Users/Bart/Desktop/Harmonized Chapters/steelHTSlist_justnumbers.txt"
alum_codes = "/mnt/c/Users/Bart/Desktop/Harmonized Chapters/aluminumHTSlist_justnumbers.txt"

# rows = ws.iter_rows(min_row=1, max_row=10, min_col=1, max_col=2)

# for a,b in rows:
#     print(a.value, b.value)

has_data = True
inv_row_count = 0

while has_data:
    inv_row_count += 1
    data = inv_ws.cell(row=inv_row_count, column=1).value
    if data == None:
        has_data = False

def find_declaration_req(codes):       # finds the codes in workbook that require steel declaration.
    declaration_req = []
    for i in range(1, inv_row_count):    
        for hc in codes:
            value = inv_ws.cell(row=i, column=6).value
            if value[:len(hc)] == hc:
                declaration_req.append(inv_ws.cell(row=i, column=6).coordinate)
                break
    return declaration_req

print(find_declaration_req(harm_codes(steel_codes)))
