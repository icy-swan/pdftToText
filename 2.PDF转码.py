'''
@Project ：PycharmProjects
@File    ：年报批量下载.py
@IDE     ：PyCharm
@Date    ：2023/5/30 11:39
'''

import pandas as pd
import requests
import os
import multiprocessing
import pdfplumber
import logging
import re
from pdfminer.high_level import extract_text

#日志配置文件
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

work_path = "/Users/bl/git/pdftToText/"

# 识别编码
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
        
        return "utf-8"  # 无法确定

#下载模块
def download_pdf(pdf_url, pdf_file_path):
    try:
        with requests.get(pdf_url, stream=True, timeout=10) as r:
            r.raise_for_status()
            with open(pdf_file_path, 'wb') as f:
                for chunk in r.iter_content(chunk_size=8192):
                    f.write(chunk)
    except requests.exceptions.RequestException as e:
        logging.error(f"下载PDF文件失败：{e}")
        return False
    else:
        return True

#文件转换
def convert(code, name, year, pdf_url, pdf_dir, txt_dir, flag_pdf):
    pdf_file_path = os.path.join(pdf_dir, f"{code:06}_{name}_{year}.pdf")
    txt_file_path = os.path.join(txt_dir, f"{code:06}_{name}_{year}.txt")

    try:
        # 下载PDF文件
        if not os.path.exists(pdf_file_path):
            retry_count = 3
            while retry_count > 0:
                if download_pdf(pdf_url, pdf_file_path):
                    break
                else:
                    retry_count -= 1
            if retry_count == 0:
                logging.error(f"下载失败：{pdf_url}")
                return

        # 转换PDF文件为TXT文件
        try:
            encoding = detect_encoding_by_fonts(pdf_file_path)
            text = extract_text(pdf_file_path, codec=encoding)
            with open(txt_file_path, 'w', encoding='utf-8') as f:
                f.write(text)
        except Exception as e:
            logging.error(f"写入文件 {code:06}_{name}_{year}时出错： {e}")

        logging.info(f"{txt_file_path} 已保存.")

    except Exception as e:
        logging.error(f"处理 {code:06}_{name}_{year}时出错： {e}")
    else:
        # 删除已转换的PDF文件，以节省空间
        if flag_pdf:
            os.remove(pdf_file_path)
            logging.info(f"{pdf_file_path} 已被删除.")



def main(file_name,pdf_dir,txt_dir,flag_pdf):
    print("程序开始运行，请耐心等待……")
    # 读取Excel文件
    try:
        df = pd.read_excel(file_name)
    except Exception as e:
        logging.error(f"读取失败，请检查路径是否设置正确，建议输入绝对路径 {e}")
        return
    try:
        os.makedirs(pdf_dir, exist_ok=True)
        os.makedirs(txt_dir, exist_ok=True)
    except Exception as e:
        logging.error(f"创建文件夹失败！请检查文件夹是否为只读！ {e}")
        return

    # 读取文件内容并存储为字典
    content_dict = ((row['公司代码'], row['公司简称'], row['年份'], row['年报链接']) for _, row in df.iterrows())

    # 多进程下载PDF并转为TXT文件
    with multiprocessing.Pool() as pool:
        for code, name, year, pdf_url in content_dict:    
            txt_file_name = f"{code:06}_{name}_{year}.txt"
            txt_file_path = os.path.join(txt_dir, txt_file_name)
            if os.path.exists(txt_file_path):
                # logging.info(f"{txt_file_name} 已存在，跳过.")
                continue
            else:
                pool.apply_async(convert, args=(code, name, year, pdf_url, pdf_dir, txt_dir, flag_pdf))

        pool.close()
        pool.join()


if __name__ == '__main__':
    # 是否删除pdf文件，True为是，False为否
    flag_pdf = False
    # 是否批量处理多个年份，True为是，False为否
    Flag = False
    if Flag:
        #批量下载并转换年份区间
        for year in range(2009,2013):
            file_name = f"{work_path}年报链接_{year}Alice.xlsx"
            # 创建存储文件的文件夹路径，如有需要请修改
            pdf_dir = f'reports/{year}/pdf'
            txt_dir = f'reports/{year}/txt'
            main(file_name,pdf_dir,txt_dir,flag_pdf)
            print(f"{year}年年报处理完毕，若报错，请检查后重新运行")
    else:
        #处理单独年份：
        #特定年份的excel表格，请务必修改。
        year = 2015
        file_name = f"{work_path}年报链接_{year}Alice.xlsx"
        pdf_dir = f'reports/{year}/pdf'
        txt_dir = f'reports/{year}/txt'
        main(file_name, pdf_dir, txt_dir, flag_pdf)
        print(f"{year}年年报处理完毕，若报错，请检查后重新运行")
