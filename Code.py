import openpyxl
from openpyxl import *
from openpyxl.drawing.image import Image


wb = openpyxl.Workbook()
ws = wb.active


image_path = 'Grue.jpg'
img = Image(image_path)
ws.add_image(img, "A2")
ws['A1'] = 'HelloWorld'
wb.save ('helloWorld.xls')

