from openpyxl import load_workbook
from openpyxl.drawing.image import Image
 
# 加载Excel工作簿
wb = load_workbook('JD金蒂.xlsx')
ws = wb.active
 
# 复制图片
for image in ws._images:
    print(image.anchor._from)
    print(image.anchor._from.col)
    print(image.anchor._from.row)
