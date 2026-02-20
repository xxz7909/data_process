"""
股票图片OCR批量识别工具
功能：批量识别 差额/代码/市值 图片，提取关键信息写入Excel
依赖：rapidocr-onnxruntime, pillow, pandas, openpyxl
"""
import os
import re
import numpy as np
from PIL import Image
from rapidocr_onnxruntime import RapidOCR
import pandas as pd


def create_ocr():
    """创建OCR实例"""
    return RapidOCR()


def load_image(path, max_width=1200):
    """加载图片并限制宽度以加速OCR"""
    img = Image.open(path)
    if img.width > max_width:
        ratio = max_width / img.width
        img = img.resize((max_width, int(img.height * ratio)), Image.LANCZOS)
    if img.mode != 'RGB':
        img = img.convert('RGB')
    return np.array(img)


def ocr_image(ocr_engine, path):
    """OCR识别图片，返回文本行列表"""
    img_arr = load_image(path)
    result, _ = ocr_engine(img_arr)
    if result:
        return [line[1] for line in result]
    return []


def extract_diff(lines):
    """从差额图片中提取差额(亿)数值"""
    text = ''.join(lines)
    # 匹配 "差额亿：XX.XX" 或 "差额亿:XX.XX"
    match = re.search(r'差额亿[：:]\s*([\d.]+)', text)
    if match:
        return match.group(1)
    # 备用：匹配任何 X.XX 后跟 著名
    match = re.search(r'([\d.]+)\s*著名', text)
    if match:
        return match.group(1)
    return ''


def _is_valid_stock_code(code_str):
    """检查是否为合法A股代码前缀"""
    valid_prefixes = [
        '000', '001', '002', '003',          # 深市主板/中小板
        '300', '301',                         # 创业板
        '600', '601', '603', '605',           # 沪市主板
        '688',                                # 科创板
        '830', '831', '832', '833', '834',    # 北交所
        '835', '836', '837', '838', '839',
        '920',                                # 北交所
    ]
    return any(code_str.startswith(p) for p in valid_prefixes)


def extract_code_and_name(lines):
    """从代码图片中提取6位代码和汉字名称"""
    full_text = ' '.join(lines)

    # 提取6位数字代码 —— 优先匹配合法股票代码前缀
    code = ''
    # 找所有连续数字段
    digit_segments = re.findall(r'\d+', full_text)
    for seg in digit_segments:
        if len(seg) == 6 and _is_valid_stock_code(seg):
            code = seg
            break
        elif len(seg) > 6:
            # 从长数字段中滑窗找合法6位代码
            for j in range(len(seg) - 5):
                candidate = seg[j:j + 6]
                if _is_valid_stock_code(candidate):
                    code = candidate
                    break
            if code:
                break
    # 回退：匹配任意6位数字
    if not code:
        code_match = re.search(r'(\d{6})', full_text)
        if code_match:
            code = code_match.group(1)
    # 再回退：匹配5位数字（OCR可能截断）
    if not code:
        code_match = re.search(r'(\d{5})', full_text)
        if code_match:
            code = code_match.group(1)

    # 提取股票名称（中文字符，可能带 -U, -UW 等后缀）
    # 排除常见干扰词
    noise_words = {'决策', '传统', '交易', '亮点', '特色', '涨停', '跌停'}
    # OCR末尾常见噪声字符（仅中文单字噪声）
    trailing_noise = set('高托拉上下')
    name = ''
    for line in lines:
        line_clean = line.strip()
        if line_clean in noise_words:
            continue
        # 匹配以中文开头，可能有 -U/-UW 后缀的股票名
        name_match = re.match(r'([\u4e00-\u9fa5]{2,}(?:-[A-Za-z]{1,3})?)', line_clean)
        if name_match:
            candidate = name_match.group(1)
            if candidate not in noise_words and len(candidate) >= 2:
                # 仅去除末尾中文噪声字符（不影响英文后缀）
                while candidate and candidate[-1] in trailing_noise:
                    candidate = candidate[:-1]
                if len(candidate) >= 2:
                    name = candidate
                    break

    return name, code


def extract_market_cap(lines):
    """从市值图片中提取总市值（带单位）"""
    full_text = '\n'.join(lines)

    # 找到所有带"亿"的数值
    values = re.findall(r'([\d,.]+亿)', full_text)
    if not values:
        return ''

    # 总市值是最大的那个亿值
    max_val = ''
    max_num = 0
    for v in values:
        num_str = v.replace('亿', '').replace(',', '')
        try:
            num = float(num_str)
            if num > max_num:
                max_num = num
                max_val = v
        except ValueError:
            continue

    return max_val


def detect_groups(img_dir):
    """自动检测图片组数"""
    groups = set()
    if not os.path.isdir(img_dir):
        return 0
    for fname in os.listdir(img_dir):
        match = re.match(r'^(\d+)_', fname)
        if match:
            groups.add(int(match.group(1)))
    return max(groups) if groups else 0


def process_images(img_dir, output_path, progress_callback=None):
    """
    主处理函数：批量识别图片并写入Excel
    
    Args:
        img_dir: 图片目录路径
        output_path: 输出Excel路径
        progress_callback: 进度回调函数 callback(current, total, message)
    
    Returns:
        (success, message) 元组
    """
    group_count = detect_groups(img_dir)
    if group_count == 0:
        return False, "未在指定目录中找到图片文件（格式：N_差额.png.bmp）"

    ocr_engine = create_ocr()
    results = []
    total_steps = group_count * 3  # 每组3张图

    for i in range(1, group_count + 1):
        step_base = (i - 1) * 3

        # --- 差额 ---
        diff_path = os.path.join(img_dir, f'{i}_差额.png.bmp')
        diff_value = ''
        if os.path.exists(diff_path):
            if progress_callback:
                progress_callback(step_base + 1, total_steps, f'正在识别 第{i}组 差额...')
            try:
                lines = ocr_image(ocr_engine, diff_path)
                diff_value = extract_diff(lines)
            except Exception as e:
                diff_value = f'识别失败: {e}'

        # --- 代码 ---
        code_path = os.path.join(img_dir, f'{i}_代码.png.bmp')
        name, code = '', ''
        if os.path.exists(code_path):
            if progress_callback:
                progress_callback(step_base + 2, total_steps, f'正在识别 第{i}组 代码...')
            try:
                lines = ocr_image(ocr_engine, code_path)
                name, code = extract_code_and_name(lines)
            except Exception as e:
                name = f'识别失败: {e}'

        # --- 市值 ---
        value_path = os.path.join(img_dir, f'{i}_市值.png.bmp')
        market_cap = ''
        if os.path.exists(value_path):
            if progress_callback:
                progress_callback(step_base + 3, total_steps, f'正在识别 第{i}组 市值...')
            try:
                lines = ocr_image(ocr_engine, value_path)
                market_cap = extract_market_cap(lines)
            except Exception as e:
                market_cap = f'识别失败: {e}'

        results.append({
            '序号': i,
            '差额(亿)': diff_value,
            '名称': name,
            '代码': code,
            '总市值': market_cap,
        })

    # 写入Excel
    df = pd.DataFrame(results)
    df.to_excel(output_path, index=False, engine='openpyxl')

    # 调整列宽
    from openpyxl import load_workbook
    wb = load_workbook(output_path)
    ws = wb.active
    col_widths = {'A': 8, 'B': 14, 'C': 18, 'D': 12, 'E': 16}
    for col_letter, width in col_widths.items():
        ws.column_dimensions[col_letter].width = width
    # 表头加粗
    from openpyxl.styles import Font
    for cell in ws[1]:
        cell.font = Font(bold=True)
    wb.save(output_path)

    return True, f"处理完成！共 {group_count} 组数据已写入 {output_path}"


if __name__ == '__main__':
    # 命令行模式直接运行
    img_dir = 'pics'
    output_path = 'result.xlsx'

    def cli_progress(current, total, msg):
        print(f'[{current}/{total}] {msg}')

    success, message = process_images(img_dir, output_path, cli_progress)
    print(message)
