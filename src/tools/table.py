import pytesseract
from PIL import Image
import os

# 指定 tesseract 的安装路径
pytesseract.pytesseract.tesseract_cmd = "/opt/homebrew/bin/tesseract"
# 加载图片
image = Image.open("image.png")
# 使用pytesseract进行OCR处理
text = pytesseract.image_to_string(image, lang="chi_sim")  # 使用中文简体模式
# 打印识别出的文本
print(text)
# 解析文本中的数字并求和
import re

# 假设我们知道数据是在特定列，我们可以通过正则表达式来找到这些数字
numbers = re.findall(r"\b-?\d+\.\d+\b", text)  # 匹配包括负数在内的浮点数
# 将字符串数字转换为浮点数并计算总和
total_sum = sum(map(float, numbers))
print(f"The sum of the numbers: {total_sum}")
