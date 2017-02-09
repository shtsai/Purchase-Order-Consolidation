import openpyxl

def set_auto_column_widths():
    column_widths = {}
    for row in newsheet.iter_rows():
        for cell in row:
            if cell.value:
                # add 5 at the end because Chinese chars take larger spaces
                column_widths[cell.column] = max((column_widths.get(cell.column, 0), len(cell.value)+5)) 
    for i, column_width in column_widths.items():
        newsheet.column_dimensions[i].width = column_width
    return


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
newsheet.title = "信用卡"

# add header to new file
newsheet.cell(row=3, column=1).value = "轉入頂新"
newsheet.cell(row=3, column=2).value = "單據日期"
newsheet.cell(row=3, column=3).value = "會員編號"
newsheet.cell(row=3, column=4).value = "託運單號"
newsheet.cell(row=3, column=5).value = "訂單號碼"
newsheet.cell(row=3, column=6).value = "訂單日期"
newsheet.cell(row=3, column=7).value = "訂購人"
newsheet.cell(row=3, column=8).value = "收貨人"
newsheet.cell(row=3, column=9).value = "卡拉蝦"
newsheet.cell(row=4, column=9).value = "原味"
newsheet.cell(row=4, column=10).value = "辣味"
newsheet.cell(row=3, column=11).value = "卡拉魷"
newsheet.cell(row=4, column=11).value = "原味"
newsheet.cell(row=4, column=12).value = "辣味"
newsheet.cell(row=3, column=14).value = "卡拉魷"
newsheet.cell(row=4, column=13).value = "芥末"
newsheet.cell(row=4, column=14).value = "原味"
newsheet.cell(row=4, column=15).value = "辣味"
newsheet.cell(row=3, column=16).value = "預留"
newsheet.cell(row=3, column=17).value = "預留"
newsheet.cell(row=3, column=18).value = "預留"
newsheet.cell(row=3, column=19).value = "預留"
newsheet.cell(row=3, column=20).value = "預留"
newsheet.cell(row=3, column=22).value = "卡拉龍珠"
newsheet.cell(row=4, column=21).value = "原味"
newsheet.cell(row=4, column=22).value = "辣味"
newsheet.cell(row=4, column=23).value = "芥末"
newsheet.cell(row=3, column=24).value = "卡拉小卷"
newsheet.cell(row=4, column=24).value = "原味"
newsheet.cell(row=4, column=25).value = "芥末"
newsheet.cell(row=3, column=27).value = "虱目魚薄燒脆片"
newsheet.cell(row=4, column=26).value = "海苔"
newsheet.cell(row=4, column=27).value = "黑胡椒"
newsheet.cell(row=4, column=28).value = "蒜香"
newsheet.cell(row=3, column=29).value = "贈送"
newsheet.cell(row=3, column=30).value = ""
newsheet.cell(row=3, column=31).value = "備貨"
newsheet.cell(row=3, column=32).value = "出貨"
newsheet.cell(row=3, column=33).value = "備注"
newsheet.cell(row=3, column=34).value = "聯絡電話"
newsheet.cell(row=3, column=35).value = "收件人聯絡電話"
newsheet.cell(row=3, column=36).value = "送貨地址"
newsheet.cell(row=3, column=37).value = "訂單金額"
newsheet.cell(row=3, column=38).value = "網路金額抵扣"
newsheet.cell(row=3, column=39).value = "Coupon"
newsheet.cell(row=3, column=40).value = "合計"
newsheet.cell(row=3, column=41).value = "買家付款方式"
newsheet.cell(row=3, column=42).value = "網購收款狀態"
newsheet.cell(row=3, column=43).value = "品號"



r = 4
c = 1
# print all non-empty cells
for row in sheet.iter_rows():
    if r == 4:  # skip the first line
        r += 1
        continue

    if ("常溫" not in row[48].value):   # only handle 常溫 for now
        continue
    
    # 1. date
    # 2. order date 
    newsheet.cell(row=r, column=2).value = (row[0].value[:11])
    # 3. ID number
    newsheet.cell(row=r, column=3).value = (row[55].value)
    # 4. tracking ID 
    
    # 5. order ID 
    newsheet.cell(row=r, column=5).value = (row[1].value[15:])
    # 6. order date 
    newsheet.cell(row=r, column=6).value = "20" + (row[1].value[8:14])
    # 7. customer name
    newsheet.cell(row=r, column=7).value = (row[53].value)
    # 8. receiptant name
    newsheet.cell(row=r, column=8).value = (row[61].value)
 
    # Quantity
    # 9. Kala shrimp original
    # 10. Kala shrimp spicy
    # 11. Kala you original
    # 12. Kala you spicy
    # 13. Kala you wasabi
    # 14. Kala crab original
    # 15. Kala crab spicy
    # 16. reserved
    # 17. reserved
    # 18. reserved
    # 19. reserved
    # 20. reserved
    # 21. Kala long zhu original
    # 22. Kala long zhu spicy
    # 23. Kala long zhu wasabi
    # 24. Kala xiao juan original
    # 25. Kala xiao juan wasabi
    # 26. fish chips sea weed
    # 27. fish chips black pepper
    # 28. fish chips garlic
    
    # 29. gift 1
    # 30. gift 2
    # 31. preparation
    # 32. ship date
    newsheet.cell(row=r, column=32).value = (row[40].value)
    # 33. Note

    # 34. phone number 
    newsheet.cell(row=r, column=34).value = (row[55].value)
    # 35. contact phone number 
    newsheet.cell(row=r, column=35).value = (row[62].value)
    # 36. shipping address
    newsheet.cell(row=r, column=36).value = (row[64].value)

    # 37. order total
    newsheet.cell(row=r, column=37).value = (row[14].value)
    # 38. point spent
    newsheet.cell(row=r, column=38).value = (row[20].value)
    # 39. coupon
    newsheet.cell(row=r, column=39).value = (row[15].value)
    # 40. amount paid 
    newsheet.cell(row=r, column=40).value = (row[27].value)
    # 41. payment method
    newsheet.cell(row=r, column=41).value = "13" 
    # 42. payment status
    newsheet.cell(row=r, column=42).value = (row[29].value)
    # 43. code

    # move to next row
    r += 1


# set column width
set_auto_column_widths()

# save new workbook
nwb.save(filename = newfilename)


# prevent output window from closing
# input()


