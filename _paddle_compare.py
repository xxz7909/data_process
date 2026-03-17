import os
from paddleocr import PaddleOCR

pics = r'c:\Users\xzw65\Desktop\jiedan\data_process\pics'
ocr = PaddleOCR(use_angle_cls=True, lang='ch')

for i in [1, 4, 9]:
    print(f'=== Group {i} ===')
    for suffix in ['差额', '代码', '市值']:
        path = os.path.join(pics, f'{i}_{suffix}.png.bmp')
        result = ocr.ocr(path, cls=True)
        print(f'[{suffix}]')
        if result and result[0]:
            for item in result[0]:
                print(item[1][0])
        else:
            print('NO_RESULT')
    print()
