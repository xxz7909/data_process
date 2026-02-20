import os
import re
import pytesseract
from PIL import Image
import pandas as pd

# 文件根目录
img_dir = 'pics'
# 组数（可根据实际图片数量调整）
group_count = 12  # 你可以根据实际图片数量调整

# 结果列表
results = []

for i in range(1, group_count + 1):
    # 文件名
    diff_file = f"{i}_差额.png.bmp"
    code_file = f"{i}_代码.png.bmp"
    value_file = f"{i}_市值.png.bmp"
    # 路径
    diff_path = os.path.join(img_dir, diff_file)
    code_path = os.path.join(img_dir, code_file)
    value_path = os.path.join(img_dir, value_file)

    # 差额
    if os.path.exists(diff_path):
        diff_text = pytesseract.image_to_string(Image.open(diff_path), lang='chi_sim+eng')
        diff_match = re.search(r'([0-9.]+)\s*亿', diff_text)
        diff_value = diff_match.group(1) if diff_match else ''
    else:
        diff_value = ''

    # 代码和名称
    if os.path.exists(code_path):
        code_text = pytesseract.image_to_string(Image.open(code_path), lang='chi_sim+eng')
        code_match = re.search(r'([\u4e00-\u9fa5]+)\s*([0-9]{6})', code_text)
        code_name = code_match.group(1) if code_match else ''
        code_num = code_match.group(2) if code_match else ''
    else:
        code_name = ''
        code_num = ''

    # 市值
    if os.path.exists(value_path):
        value_text = pytesseract.image_to_string(Image.open(value_path), lang='chi_sim+eng')
        value_match = re.search(r'([0-9,.]+[万亿元]*)', value_text)
        value = value_match.group(1) if value_match else ''
    else:
        value = ''

    results.append({
        '序号': i,
        '差额(亿)': diff_value,
        '名称': code_name,
        '代码': code_num,
        '市值': value
    })

# 写入Excel
df = pd.DataFrame(results)
df.to_excel('result.xlsx', index=False)

print('识别完成，结果已写入 result.xlsx')
