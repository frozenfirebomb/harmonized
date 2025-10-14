import openpyxl

inv_wb = openpyxl.load_workbook("/mnt/c/Users/Bart/Desktop/Harmonized Chapters/INVDAY_master.xlsx")
inv_ws = inv_wb[ 'Sheet1']

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

# rows = inv_ws.iter_rows(min_row=1, max_row=inv_row_count-1, min_col=1, max_col=2)

# for a,b in rows:
#     print(a.value, b.value)

# for i in range(2, inv_row_count):
#     print(inv_ws.cell(row=i, column=6).value)

def harm_codes(fp):        #returns the contents of a text file as a list of strings from a file path provided as a string
    codes = []
    with open(fp) as f:
        file_contents = f.read()    
    for content in file_contents.split():
        codes.append(content)
    return codes

# for i in range(2, inv_row_count):
#     if inv_ws.cell(row=i, column=6).value in harm_codes("/mnt/c/Users/Bart/Desktop/Harmonized Chapters/steelHTSlist_justnumbers.txt"):
#         print("steel decleration required", inv_ws.cell(row=i, column=6).value)
for i in range(2, inv_row_count):
    for hc in harm_codes("/mnt/c/Users/Bart/Desktop/Harmonized Chapters/steelHTSlist_justnumbers.txt"):
        hc_len = len(hc)
        value = inv_ws.cell(row=i, column=6).value
        if value[:hc_len] == hc:
            print("steel decleration required", inv_ws.cell(row=i, column=6).value)
            break