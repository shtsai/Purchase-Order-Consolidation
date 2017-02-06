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

# add header to new file
newsheet.cell(row=3, column=1).value = "轉入頂新"
newsheet.cell(row=3, column=2).value = "單據日期"
newsheet.cell(row=3, column=3).value = "會員編號"
newsheet.cell(row=3, column=4).value = "託運單號"
newsheet.cell(row=3, column=5).value = "訂單號碼"
newsheet.cell(row=3, column=6).value = "訂單日期"
newsheet.cell(row=3, column=7).value = "訂購人"
newsheet.cell(row=3, column=8).value = "收貨人"

'''
newsheet.cell(row=3, column=9).value = ""
newsheet.cell(row=3, column=10).value = ""
newsheet.cell(row=3, column=11).value = ""
newsheet.cell(row=3, column=12).value = ""
newsheet.cell(row=3, column=13).value = ""
newsheet.cell(row=3, column=14).value = ""
newsheet.cell(row=3, column=15).value = ""
newsheet.cell(row=3, column=16).value = ""
newsheet.cell(row=3, column=17).value = ""
newsheet.cell(row=3, column=18).value = ""
newsheet.cell(row=3, column=19).value = ""
newsheet.cell(row=3, column=20).value = ""
newsheet.cell(row=3, column=21).value = ""
newsheet.cell(row=3, column=22).value = ""
newsheet.cell(row=3, column=23).value = ""
'''

newsheet.cell(row=3, column=24).value = "贈送"
newsheet.cell(row=3, column=25).value = "備貨"
newsheet.cell(row=3, column=26).value = "出貨"
newsheet.cell(row=3, column=27).value = "備注"
newsheet.cell(row=3, column=28).value = "聯絡電話"
newsheet.cell(row=3, column=29).value = "收件人聯絡電話"
newsheet.cell(row=3, column=30).value = "送貨地址"
newsheet.cell(row=3, column=31).value = "訂單金額"
newsheet.cell(row=3, column=32).value = "網路金額抵扣"
newsheet.cell(row=3, column=33).value = "合計"
newsheet.cell(row=3, column=34).value = "買家付款方式"
newsheet.cell(row=3, column=35).value = "網購收款狀態"
newsheet.cell(row=3, column=36).value = "品號"



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
