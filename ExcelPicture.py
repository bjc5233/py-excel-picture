#!/usr/bin/python
#! -*- coding: utf8 -*-
# 说明
#   照片像素化(excel)
# 参考
#   https://blog.csdn.net/JohnJim0/article/details/104329144
# 注意点
#   1. excel 全屏显示         2010需要从选项中添加"全屏显示"到功能区
#   2. excel 单元格正方形     全选，拖动分割线，将宽高的数值拖到同一值
#   3. excel VBA编辑器        2010需要从选项中添加开发者工具  快捷键Alt+F11
#   4. excel VBA代码位置      对象树ThisWorkbook下
# external
#   date       2020-05-03 17:19:01
#   face       ●ω●
#   weather    Shanghai Sunny 32℃
from PIL import Image


# ================ 基本配置 ================
picPath = 'C:\\path\\python\\ExcelPicture\\resources\\1.jpg'
picShowStyle = 1        #1:从左到右   2::从上到下
picExcelFitLine = 125   #Excel中显示较为完整的行数
# ================ 基本配置 ================


imgLoad = Image.open(picPath)
widthSrc, heightSrc = imgLoad.size
if (heightSrc > picExcelFitLine):   
    height = picExcelFitLine
else:
    height = heightSrc
width = int(widthSrc * height / heightSrc)
imgLoad = imgLoad.resize((width, height), Image.ANTIALIAS)
img = imgLoad.convert("RGB")

RGBs = open('resources\\config.txt', 'w')
RGBs.write(str(picShowStyle) + "\t" + str(width) + "\t" + str(height))
RGBs.close()
RGBs2 = open('resources\\pictureRGB.txt', 'w')
for y in range(height):
    for x in range(width):
        rgb = str(img.getpixel((x,y)))
        RGBs2.write(rgb[1:-1] + "\t")
    RGBs2.write("\n")
RGBs2.close()