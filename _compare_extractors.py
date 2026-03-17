import os
import re
import numpy as np
from PIL import Image
from paddleocr import PaddleOCR


def _normalize_minus(text):
    for ch in '一—–−﹣':
        text = text.replace(ch, '-')
    return text.replace('\uff0d', '-')


def _detect_text_color(img_arr, box):
    pts = np.array(box, dtype=np.int32)
    x_min, y_min = pts.min(axis=0)
    x_max, y_max = pts.max(axis=0)
    h, w = img_arr.shape[:2]
    x_min, y_min = max(0, x_min), max(0, y_min)
    x_max, y_max = min(w, x_max), min(h, y_max)
    crop = img_arr[y_min:y_max, x_min:x_max]
    if crop.size == 0:
        return 'unknown'

    pixels = crop.reshape(-1, 3).astype(np.float32)
    r, g, b = pixels[:, 0], pixels[:, 1], pixels[:, 2]
    brightness = (r + g + b) / 3
    mask = (brightness > 30) & (brightness < 230)
    if mask.sum() < 10:
        return 'unknown'

    r, g, b = r[mask], g[mask], b[mask]
    red_mask = (r > 120) & (r > g * 1.4) & (r > b * 1.4)
    green_mask = (g > 120) & (g > r * 1.4) & (g > b * 1.4)
    red_ratio = red_mask.sum() / len(r)
    green_ratio = green_mask.sum() / len(r)
    if green_ratio > 0.05 and green_ratio > red_ratio:
        return 'green'
    if red_ratio > 0.05 and red_ratio > green_ratio:
        return 'red'
    return 'unknown'


def extract_diff(lines, ocr_items=None, img_arr=None):
    text = _normalize_minus(''.join(lines))
    match = re.search(r'(-?\d+\.?\d*)', text)
    if not match:
        return ''

    value_str = match.group(1)
    has_text_minus = value_str.startswith('-')
    if not has_text_minus and ocr_items and img_arr is not None:
        color = 'unknown'
        for item_text, box, _ in ocr_items:
            normalized = _normalize_minus(item_text)
            if value_str in normalized:
                color = _detect_text_color(img_arr, box)
                if color != 'unknown':
                    break
        if color == 'green':
            value_str = '-' + value_str
    return value_str


def _is_valid_stock_code(code_str):
    valid_prefixes = [
        '000', '001', '002', '003', '300', '301', '600', '601', '603', '605',
        '688', '830', '831', '832', '833', '834', '835', '836', '837', '838',
        '839', '920',
    ]
    return any(code_str.startswith(p) for p in valid_prefixes)


def extract_code_and_name(lines):
    full_text = ' '.join(lines)
    code = ''
    digit_segments = re.findall(r'\d+', full_text)
    for seg in digit_segments:
        if len(seg) == 6 and _is_valid_stock_code(seg):
            code = seg
            break
        if len(seg) > 6:
            for j in range(len(seg) - 5):
                candidate = seg[j:j + 6]
                if _is_valid_stock_code(candidate):
                    code = candidate
                    break
            if code:
                break

    if not code:
        code_match = re.search(r'(\d{6})', full_text)
        if code_match:
            code = code_match.group(1)
    if not code:
        code_match = re.search(r'(\d{5})', full_text)
        if code_match:
            code = code_match.group(1)

    noise_words = {
        '决策', '传统', '交易', '亮点', '特色', '涨停', '跌停', '盘面', '资金',
        '主力', '散户', '买入', '卖出', '成交', '板块', '行情', '分时', '日线',
        '周线', '月线', '指标',
    }
    trailing_noise = set('高托拉上下')
    name = ''
    for line in lines:
        line_clean = line.strip()
        if line_clean in noise_words:
            continue
        name_match = re.match(r'([\u4e00-\u9fa5]{2,}(?:-[A-Za-z]{1,3})?)', line_clean)
        if name_match:
            candidate = name_match.group(1)
            if candidate not in noise_words and len(candidate) >= 2:
                while candidate and candidate[-1] in trailing_noise:
                    candidate = candidate[:-1]
                if len(candidate) >= 2:
                    name = candidate
                    break
    return name, code


def extract_market_cap(lines):
    full_text = '\n'.join(lines)
    values = re.findall(r'([\d,.]+)亿', full_text)
    if not values:
        return ''

    max_val = ''
    max_num = 0
    for v in values:
        num_str = v.replace(',', '')
        try:
            num = float(num_str)
            if num > max_num:
                max_num = num
                max_val = num_str
        except ValueError:
            continue
    return max_val


pics = r'c:\Users\xzw65\Desktop\jiedan\data_process\pics'
ocr = PaddleOCR(use_angle_cls=True, lang='ch', show_log=False)

for i in range(1, 11):
    row = {'序号': i}
    diff_path = os.path.join(pics, f'{i}_差额.png.bmp')
    code_path = os.path.join(pics, f'{i}_代码.png.bmp')
    value_path = os.path.join(pics, f'{i}_市值.png.bmp')

    diff_result = ocr.ocr(diff_path, cls=True)
    diff_items = []
    if diff_result and diff_result[0]:
        for item in diff_result[0]:
            diff_items.append((item[1][0], item[0], item[1][1]))
    diff_img_arr = np.array(Image.open(diff_path).convert('RGB'))
    row['差额(亿)'] = extract_diff([x[0] for x in diff_items], diff_items, diff_img_arr)

    code_result = ocr.ocr(code_path, cls=True)
    code_lines = [item[1][0] for item in code_result[0]] if code_result and code_result[0] else []
    row['名称'], row['代码'] = extract_code_and_name(code_lines)

    value_result = ocr.ocr(value_path, cls=True)
    value_lines = [item[1][0] for item in value_result[0]] if value_result and value_result[0] else []
    row['总市值'] = extract_market_cap(value_lines)

    print(row)
