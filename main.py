import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font

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

def final_ws_editing(skus):         # adds shipment info to the final sheet for each sku that needs to be declared.
    global final_ws_tracker
    sku_row_in_metal_master = [] 

    for i in list(metal_master_ws.iter_rows(min_row=1, max_row=metal_master_row_count, min_col=3, max_col=3)):
        for x in i:
            sku_row_in_metal_master.append(x.value)

    for sku in skus:
        if sku.value not in sku_row_in_metal_master:
            print(f"Sku {sku.value} needs to be declared.")
            continue

        for i in range(1, metal_master_row_count):
            value = metal_master_ws.cell(row=i, column=3).value
            
            if sku.value == value:
                metal_master_range = metal_master_ws[f"A{i-1}" : f"H{i+2}"] # cell range of information relevant to the sku
                metal_master_cells = [val for row in metal_master_range for val in row] # nested tuple unpacking
                
                final_range = final_ws[f"A{1+final_ws_tracker}" : f"H{4+final_ws_tracker}"] # cell range to have information added
                final_cells = [val for row in final_range for val in row]
                
                for i in range(len(final_cells)):    # 
                    final_ws[final_cells[i].coordinate] = metal_master_ws[metal_master_cells[i].coordinate].value
                final_ws[f"A{2 + final_ws_tracker}"] = inv_ws[f"A{sku.row}"].value  # adds in the shipment ID
                final_ws[f"B{2 + final_ws_tracker}"] = inv_ws[f"B{sku.row}"].value  # adds in the Invoice Number
                final_ws[f"D{2 + final_ws_tracker}"] = inv_ws[f"D{sku.row}"].value  # changes the quanity of items sent
                final_ws[f"E{2 + final_ws_tracker}"] = round(inv_ws[f"E{sku.row}"].value    # calculates the value of metal declared
                                                             * final_ws[f"E{2 + final_ws_tracker}"].value, 2)
                final_ws[f"H{2 + final_ws_tracker}"] = (final_ws[f"D{2 + final_ws_tracker}"].value  # calculates total value of metal
                                                        * final_ws[f"E{2 + final_ws_tracker}"].value)
                
                font_bold = Font(bold= True)

                first_row_range = final_ws[f"A{1 + final_ws_tracker}" : f"H{1 + final_ws_tracker}"]
                first_row = [val for row in first_row_range for val in row]
                
                third_row_range = final_ws[f"A{3 + final_ws_tracker}" : f"H{3 + final_ws_tracker}"]
                third_row = [val for row in third_row_range for val in row]

                fourth_row_range = final_ws[f"A{4 + final_ws_tracker}" : f"H{4 + final_ws_tracker}"]
                fourth_row = [val for row in fourth_row_range for val in row]

                for cell in first_row:
                    final_ws[cell.coordinate].font = font_bold
                
                for cell in third_row:
                    final_ws[cell.coordinate].font = font_bold

                final_ws_tracker += 5

final_ws_editing(steel_sku)
final_ws_editing(alum_sku)

final_wb.save("/mnt/c/Users/Bart/Desktop/Harmonized Chapters/final_test.xlsx")