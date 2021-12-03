# coding=utf-8
from tkinter import messagebox
from openpyxl import load_workbook
import tkinter.filedialog

from StudyPython.src import tkinter_utils, xlsx_utils

window = tkinter_utils.createWindow('全委子表处理工具')
# 省会和城市所属的事业部
provinceCityDict = {
    '浙江-嘉兴': '事业一部',
    '浙江-湖州': '事业一部',
    '浙江-金华': '事业一部',
    '浙江-丽水': '事业一部',
    '浙江-衢州': '事业一部',
    '安徽': '事业一部',
    '湖北': '事业一部',
    '广东': '事业一部',
    '广西': '事业一部',
    '海南': '事业一部',
    '福建': '事业一部',
    '浙江-温州': '事业二部',
    '浙江-台州': '事业二部',
    '山东': '事业二部',
    '重庆': '事业二部',
    '四川': '事业二部',
    '贵州': '事业二部',
    '云南': '事业二部',
    '浙江-绍兴': '事业三部',
    '浙江-宁波': '事业三部',
    '浙江-舟山': '事业三部',
    '江苏': '事业三部',
    '河北': '事业三部',
    '北京': '事业三部',
    '黑龙江': '事业三部',
    '吉林': '事业三部',
    '辽宁': '事业三部',
    '天津': '事业三部',
    '陕西': '事业三部',
    '山西': '事业三部',
    '新疆': '事业三部',
    '青海': '事业三部',
    '甘肃': '事业三部',
    '宁夏': '事业三部',
    '河南': '事业三部',
    '内蒙古': '事业三部',
    '浙江-杭州': '事业四部',
    '上海': '事业四部',
    '湖南': '事业四部',
    '江西': '事业四部',
}


def handelFiles():
    fileName = tkinter.filedialog.askopenfilename(filetypes=[('xlsx files', '*.xlsx')])
    if fileName:
        # 选择了文件
        tkinter_utils.clearText(text=text, window=window)
        # 展示已选文件
        tkinter_utils.updateText(content='已选文件：' + fileName, text=text, window=window)
        # 打开源xlsx表
        wb = load_workbook(fileName)
        # 已存在的全部工作簿
        sheetNames = wb.sheetnames
        # 处理指定的工作簿
        sheet_name = sheetNames[0]
        sheet = wb[sheet_name]
        # 循环遍历，读取，判断
        height = sheet.row_dimensions[sheet.max_column].height
        targetColumn = sheet.max_column + 1
        xlsx_utils.insertContent(sheet=sheet, row=2, col=targetColumn, content='所属事业部', rowHeight=height)
        # 合并单元格
        sheet.merge_cells(start_column=targetColumn, end_column=targetColumn, start_row=2, end_row=4)
        startRow = 5
        for row in range(startRow, sheet.max_row + 1):
            # 省会
            province = xlsx_utils.parserMergedCell(sheet=sheet, row=row, col=sheet['B' + str(row)].column).value
            # 城市
            city = xlsx_utils.parserMergedCell(sheet=sheet, row=row, col=sheet['C' + str(row)].column).value
            # 项目名称
            project = xlsx_utils.parserMergedCell(sheet=sheet, row=row, col=sheet['G' + str(row)].column).value
            if province == '浙江':
                pcKey = str(province) + '-' + str(city)
                if pcKey in provinceCityDict:
                    bmKey = provinceCityDict[pcKey]
                    height = sheet.row_dimensions[targetColumn - 1].height
                    xlsx_utils.insertContent(sheet=sheet, row=row, col=targetColumn, content=bmKey, rowHeight=height)
                else:
                    tkinter_utils.updateText(content='找不到地区对应的部门数据，请检查配置：' + str(pcKey) + ' ' + str(project), text=text,
                                             window=window)
            else:
                if province in provinceCityDict:
                    bmKey = provinceCityDict[province]
                    height = sheet.row_dimensions[targetColumn - 1].height
                    xlsx_utils.insertContent(sheet=sheet, row=row, col=targetColumn, content=bmKey, rowHeight=height)
                else:
                    tkinter_utils.updateText(content='找不到地区对应的部门数据，请检查配置：' + str(province) + ' ' + str(project),
                                             text=text, window=window)
        # 保存输出文件
        wb.save(fileName)
        # 关闭文件
        wb.close()
        messagebox.showinfo('提示', '文件处理完成！！！')
    else:
        # 取消了选择
        messagebox.showinfo('提示', '未选择文件！！！')


btn = tkinter.Button(window, text='选择文件', command=handelFiles)
btn.pack()
text = tkinter.Text(window)
text.pack()
window.mainloop()
