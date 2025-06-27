import os
import pdfplumber
from pdfminer.high_level import extract_text

def detect_encoding_by_fonts(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        # 收集文档中所有使用的字体
        all_fonts = set()
        for page in pdf.pages:
            for char in page.chars:
                if "fontname" in char:
                    fontname = char["fontname"].lower()
                    all_fonts.add(fontname)
        
        # 分析字体名称判断编码
        encoding_map = {
            "gb": "gbk",          # 简体中文
            "gbk": "gbk",
            "gb2312": "gb2312",
            "big5": "big5",        # 繁体中文
            "msung": "big5",       # 繁体中文常见字体
            "gothic": "shift-jis", # 日文
            "mincho": "shift-jis",
            "batang": "euc-kr",    # 韩文
            "dotum": "euc-kr",
            "cyrillic": "cp1251",  # 俄文
            "1251": "cp1251"
        }
        
        for font in all_fonts:
            for key, encoding in encoding_map.items():
                if key in font:
                    return encoding
        
        return None  # 无法确定

file_name = {
    2009: "600844_丹化科技_2009",
}
# 遍历file_name
for year, name in file_name.items():
    pdf_path = f"/Users/bl/git/pdftToText/reports/{year}/txt/{name}.pdf"
    txt_file_path = f"{pdf_path.replace('.pdf','.txt')}"
    # 检测pdf_path下文件是否存在
    if not os.path.exists(pdf_path):
        # 输出log
        print(f"{pdf_path} 不存在")
        continue
    # 识别编码
    encoding = detect_encoding_by_fonts(pdf_path)
    # 指定提取
    text = extract_text(pdf_path, codec=encoding)
    with open(txt_file_path, 'w', encoding='utf-8') as f:
        f.write(text)
    # 输出log
    print(f"{pdf_path} 已转换为 {txt_file_path}")