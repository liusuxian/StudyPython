# coding:utf-8
from tkinter import *
from tkinter import messagebox
from PIL import Image
import tkinter.filedialog
import warnings
import shutil
import fitz
import glob
import os

warnings.filterwarnings('ignore')
window = Tk()
# 设置窗口大小
winWidth = 800
winHeight = 600
# 获取屏幕分辨率
screenWidth = window.winfo_screenwidth()
screenHeight = window.winfo_screenheight()
x = int((screenWidth - winWidth) / 2)
y = int((screenHeight - winHeight) / 2)
# 设置主窗口标题
window.title("绿城咨询知识库系统资源处理工具")
# 设置窗口初始位置在屏幕居中
window.geometry("%sx%s+%s+%s" % (winWidth, winHeight, x, y))
# 设置窗口图标
window.iconbitmap("./image/favicon.ico")
# 设置窗口宽高固定
window.resizable(0, 0)


# 检查整个字符串是否包含中文
def is_chinese(string):
    """
    检查整个字符串是否包含中文
    :param string: 需要检查的字符串
    :return: bool
    """
    for ch in string:
        if u'\u4e00' <= ch <= u'\u9fff':
            return True

    return False


# PDF第一页转PNG图片
def pdf_page1_to_png(image_path, pdf_file):
    """
    PDF第一页转PNG图片
    :param image_path: 图片文件保存的目录
    :param pdf_file: pdf文件全路径
    :return: 图片文件全路径
    """
    # 读取pdf文件
    pdf = fitz.open(pdf_file)
    # 获取pdf文件第一页内容
    page = pdf[0]
    # 每个尺寸的缩放系数为2，这将为我们生成分辨率提高四倍的图像
    zoom_x = 2.0
    zoom_y = 2.0
    # 旋转角度
    rotate = int(0)
    trans = fitz.Matrix(zoom_x, zoom_y).prerotate(rotate)
    pm = page.get_pixmap(matrix=trans, alpha=False)
    # 保存图片文件
    filePrefix = pdf_file.split('/')[-1].split('.')[0]
    imageName = image_path + '/' + filePrefix + '.png'
    pm.save(imageName)
    pdf.close()
    return imageName


# PDF所有页转JPG图片
def pdf_allpage_to_jpg(pdf_file, image_path=''):
    """
    PDF所有页转JPG图片
    :param pdf_file: pdf文件全路径
    :param image_path: 图片文件保存的目录
    :return: 图片文件保存的全路径
    """
    # 读取pdf文件
    pdf = fitz.open(pdf_file)
    # 获取pdf文件总页数
    total = pdf.pageCount
    # 保存图片文件全路径
    filePrefix = pdf_file.split('/')[-1].split('.')[0]
    newImagePath = os.path.join(image_path, filePrefix)
    # 判断路径是否存在
    isExists = os.path.exists(newImagePath)
    if not isExists:
        os.makedirs(newImagePath)

    for pg in range(total):
        page = pdf[pg]
        zoom = int(100)
        rotate = int(0)
        trans = fitz.Matrix(zoom / 100.0, zoom / 100.0).prerotate(rotate)
        pm = page.get_pixmap(matrix=trans, alpha=False)
        imageName = image_path + '/%s.jpg' % str(pg + 1)
        pm.save(imageName)
    pdf.close()
    return image_path


# 合并图片生成PDF
def pic_to_pdf(image_path, out_file=''):
    """
    合并图片生成PDF
    :param image_path: 图片文件全路径
    :param out_file: PDF文件保存的目录
    :return:
    """
    doc = fitz.open()
    # 读取图片，确保按文件名排序
    for img in sorted(glob.glob(image_path + '/*')):
        # 打开图片
        img_doc = fitz.open(img)
        # 使用图片创建单页的PDF
        pdf_bytes = img_doc.convert_to_pdf()
        img_pdf = fitz.open("pdf", pdf_bytes)
        # 将当前页插入文档
        doc.insert_pdf(img_pdf)

    pdfName = image_path.split('/')[-1]
    outPdfFile = out_file + '/' + pdfName + '.pdf'
    # 保存pdf文件
    doc.save(outPdfFile)
    shutil.rmtree(image_path, ignore_errors=True)
    doc.close()


# 压缩PDF
def compress_pdf(infile, outfile='', mb=1000):
    """
    压缩PDF
    :param infile: 输入文件
    :param outfile: 输出文件
    :param mb: 文件大小多少mb以内不进行压缩
    :return:
    """
    size = get_size(infile)
    if size <= mb:
        return
    outPath = pdf_allpage_to_jpg(infile, outfile)
    pic_to_pdf(outPath, outfile)


# 获取文件的大小
def get_size(file):
    """
    获取文件的大小
    :return: 文件大小（单位：KB）
    """
    # 获取文件大小:KB
    size = os.path.getsize(file)
    return size / 1024


# 拼接输出文件地址
def get_outfile(infile, outfile):
    """
    拼接输出文件地址
    :return: 输出文件地址
    """
    if outfile:
        return outfile
    path, suffix = os.path.splitext(infile)
    outfile = '{}-out{}'.format(path, suffix)
    return outfile


# 不改变图片尺寸压缩到指定大小
def compress_image(infile, outfile='', mb=150, step=10, quality=75):
    """
    不改变图片尺寸压缩到指定大小
    :param infile: 压缩源文件
    :param outfile: 压缩文件保存地址
    :param mb: 压缩目标，KB
    :param step: 每次调整的压缩比率
    :param quality: 初始压缩比率
    :return: 压缩文件地址，压缩文件大小
    """
    o_size = get_size(infile)
    if o_size <= mb:
        return infile
    outfile = get_outfile(infile, outfile)
    while o_size > mb:
        im = Image.open(infile)
        im.save(outfile, quality=quality)
        if quality - step < 0:
            break
        quality -= step
        o_size = get_size(outfile)
    return outfile, get_size(outfile)


# 修改图片尺寸
def resize_image(infile, outfile='', w=368, h=207):
    """
    修改图片尺寸
    :param w: 设置的宽度
    :param h: 设置的高度
    :param infile: 图片源文件
    :param outfile: 重设尺寸文件保存地址
    :return:
    """
    im = Image.open(infile)
    out = im.resize((w, h), Image.ANTIALIAS)
    outfile = get_outfile(infile, outfile)
    out.save(outfile)


# 更新窗口上的文本显示内容
def update_text(content):
    """
    更新窗口上的文本显示内容
    :param content: 文本字符串
    :return:
    """
    text.insert(INSERT, content + '\n')
    window.update()


# 清除窗口上的文本显示内容
def clear_text():
    """
    清除窗口上的文本显示内容
    :return:
    """
    text.delete('1.0', 'end')
    window.update()


def select_dirs():
    clear_text()
    fileDir = tkinter.filedialog.askdirectory()
    for root, dirs, files in os.walk(fileDir):
        # root 表示当前正在访问的文件夹路径
        # dirs 表示该文件夹下的子目录名list
        # files 表示该文件夹下的文件list
        # 遍历所有的文件夹
        for subDir in dirs:
            for sRoot, sDirs, sFiles in os.walk(os.path.join(root, subDir)):
                for sFile in sFiles:
                    filename = os.path.join(sRoot, sFile)
                    filePrefix = sFile.split(bytes('.'))[0]
                    projectName = sRoot.split(bytes('/'))[-1]
                    if sFile.endswith(bytes('.pdf')) and filePrefix != projectName:
                        # pdf 转 image
                        imageName = pdf_page1_to_png(sRoot, filename)
                        # image 修改尺寸
                        resize_image(imageName, imageName)
                        # image 压缩
                        compress_image(imageName, imageName)
                        text.insert(INSERT, "已处理：" + str(filename) + "\n")
                        window.update()
                    elif sFile.endswith(bytes('.xlsx')) or sFile.endswith(bytes('.DS_Store')) or (
                            bytes('.~') in sFile) or (bytes('._') in sFile) or (bytes('._') in sFile):
                        # 删除不需要的文件
                        os.remove(filename)
    messagebox.showinfo("提示", "文件处理完成！！！")


tkinter.Button(window, text="选择文件夹", command=select_dirs).pack()
text = tkinter.Text(window)
text.insert(INSERT, "")
text.pack()
window.mainloop()
