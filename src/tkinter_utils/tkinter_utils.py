from tkinter import Tk, INSERT, Text, NORMAL, DISABLED


# 创建窗口
def createWindow(title: str, width: int = 800, height: int = 600, resizable: bool = False):
    """
    创建窗口
    :param title: 窗口标题
    :param width: 窗口宽度
    :param height: 窗口高度
    :param resizable: 是否可调整窗口尺寸
    :return: 返回窗口标识
    """
    window = Tk()
    # 设置窗口大小
    winWidth = width
    winHeight = height
    # 获取屏幕分辨率
    screenWidth = window.winfo_screenwidth()
    screenHeight = window.winfo_screenheight()
    x = int((screenWidth - winWidth) / 2)
    y = int((screenHeight - winHeight) / 2)
    # 设置主窗口标题
    window.title(title)
    # 设置窗口初始位置在屏幕居中
    window.geometry('%sx%s+%s+%s' % (winWidth, winHeight, x, y))
    # 设置是否可调整窗口尺寸
    window.resizable(resizable, resizable)
    return window


# 更新窗口上的文本显示内容
def updateText(content: str, text: Text, window: Tk):
    """
    更新窗口上的文本显示内容
    :param content: 文本字符串
    :param text: tkinter.Text
    :param window: tkinter.Tk
    :return:
    """
    text.config(state=NORMAL)
    text.insert(INSERT, content + '\n')
    text.config(state=DISABLED)
    window.update()


# 清除窗口上的文本显示内容
def clearText(text: Text, window: Tk):
    """
    清除窗口上的文本显示内容
    :return:
    """
    text.config(state=NORMAL)
    text.delete('1.0', 'end')
    text.config(state=DISABLED)
    window.update()
