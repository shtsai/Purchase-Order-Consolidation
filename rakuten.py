import openpyxl

# path name
# Windows
# filename = 'C:/Users/user/Desktop/樂天 版本.xlsx'
# Linux
filename = '/home/sean/Desktop/rakuten/樂天 版本.xlsx'

# open workbook
wb = openpyxl.load_workbook(filename)

# open work sheet, by default the first sheet
sheet = wb.get_sheet_by_name('Sheet1')

# open write file
nwb = openpyxl.Workbook()
newfilename = 'rakuten.xlsx'
newsheet = nwb.active
newsheet.title = "rakuten"

r = 3
c = 1
# print all non-empty cells
for row in sheet.iter_rows():
    if r == 3:  # skip the first line
        r += 1
        continue

    if ("常溫" not in row[48].value):   # only handle 常溫 for now
        continue

    # 1. date ???
    newsheet.cell(row=r, column=2).value = (row[0].value[:11])
    # 2. ID number
    newsheet.cell(row=r, column=3).value = (row[55].value)
    # 3. tracking ID ???
    
    # 4. order ID ???
    newsheet.cell(row=r, column=5).value = (row[1].value[15:])
    # 5. order date ???
    newsheet.cell(row=r, column=6).value = (row[0].value[:11])
    # 6. customer name
    newsheet.cell(row=r, column=7).value = (row[53].value)
    # 7. receiptant name
    newsheet.cell(row=r, column=8).value = (row[61].value)
 
    # Quantity
    # 8. Kala shrimp original
    # 9. Kala shrimp spicy
    # 10. Kala you original
    # 11. Kala you spicy
    # 12. Kala you wasabi
    # 13. Kala crab original
    # 14. Kala crab spicy
    # 15. Kala long zhu original
    # 16. Kala long zhu spicy
    # 17. Kala long zhu wasabi
    # 18. Kala xiao juan original
    # 19. Kala xiao juan wasabi
    # 20. fish chips sea weed
    # 21. fish chips black pepper
    # 22. fish chips garlic
    
    # 23. gift 1
    # 24. gift 2
    # 25. preparation
    # 26. ship date
    # 27. Note

    # 28. phone number 
    newsheet.cell(row=r, column=29).value = (row[55].value)
    # 29. contact phone number 
    newsheet.cell(row=r, column=30).value = (row[55].value)
    # 30. shipping address
    newsheet.cell(row=r, column=31).value = (row[64].value)

    # 31. amount due
    # 32. online discount
    # 33. total
    # 34. payment method
    # 35. payment status
    # 36. code


    # move to next row
    r += 1


nwb.save(filename = newfilename)

# prevent output window from closing
# input()
