'''
@Project ：PycharmProjects
@File    ：词频分析.py
@IDE     ：PyCharm
@Date    ：2023/5/30 14:34

文本和相关性分析的逻辑
1. 待抽取的词分为N个数组(比如2个 - 供应链、风险）
2. 对年报进行分词，成为一个大数组
3. 按照N个数组内的每个词，寻找每个词出现的位置，构建N个位置数组
4. 对比遍历N个数组，寻找在K个词是否存在其他关联词

'''
import os
import re
import xlwt
import jieba

# 输入年份区间
start_year = "2013"
end_year = "2023"
work_path="/Users/bl/git/pdftToText/reports"

def extract_keywords(filename, keywords):

    keyword_counts = [0] * len(keywords)
    total_words = 0  # 统计总字数

    try:
        with open(filename, 'r', encoding='utf-8') as f:
            content = f.read()

        # 将关键词添加到自定义词典
        for word in keywords:
            jieba.add_word(word)

        # 使用jieba库进行分词
        words = jieba.cut(content)
        words = [word for word in words if word.strip()]

        # 统计关键词出现次数
        for i, keyword in enumerate(keywords):
            keyword_counts[i] = words.count(keyword)

        total_words = len(words)  # 统计总字数

    except FileNotFoundError:
        print(f"文件不存在: {filename}")
    except PermissionError:
        print(f"没有访问权限: {filename}")
    except Exception as e:
        print(f"从文件中获取关键词失败: {filename}")
        print(str(e))

    return keyword_counts, total_words


def count_txt_files(folder_path, start_year=None, end_year=None):
    """
    统计指定文件夹及其子文件夹中符合年份要求的所有txt文件数量。
    """
    total_files = 0
    processed_files = 0

    try:
        # 遍历文件夹中的所有文件和子文件夹
        for root, dirs, files in os.walk(folder_path):
            # 根据文件夹路径获取年份信息
            match = re.match(r'.*([12]\d{3}).*', os.path.basename(root))
            if match:
                year = match.group(1)
                if (start_year is not None and int(year) < int(start_year)) or (end_year is not None and int(year) > int(end_year)):
                    # 如果年份不符合要求，则移除该文件夹，防止遍历其中的文件
                    dirs[:] = []
                    continue

            # 遍历当前文件夹中的所有txt文件
            for filename in files:
                if filename.endswith('.txt'):
                    total_files += 1
    except FileNotFoundError:
        print(f"文件夹不存在: {folder_path}")
    except PermissionError:
        print(f"没有访问权限: {folder_path}")
    except Exception as e:
        print(f"统计文件数量失败: {folder_path}")
        print(str(e))

    return total_files

def process_files(folder_path, keywordsGroup, start_year=None, end_year=None):
    """
    处理指定文件夹及其子文件夹中的所有txt文件，提取关键词并统计词频和总字数，将结果存储到Excel表格中。
    """
    try:
        # 创建Excel工作簿
        workbook = xlwt.Workbook(encoding="utf-8")
        worksheet = workbook.add_sheet("alice")
        row = 0
        # 添加Excel表头
        worksheet.write(row, 0, '股票代码')
        worksheet.write(row, 1, '公司简称')
        worksheet.write(row, 2, '年份')
        worksheet.write(row, 3, '总字数')  # 添加总字数列

        j = 0
        for i, keyword in enumerate(keywordsGroup):
            worksheet.write(row, i + 4, f'type{i}')  # 按类型进行计数统计
            j = i+5
        worksheet.write(row, j, '关联词计数')
        row += 1

        total_files = count_txt_files(folder_path, start_year, end_year)
        processed_files = 0

        try:
            # 遍历文件夹中的所有文件和子文件夹
            for root, dirs, files in os.walk(folder_path):
                # 根据文件夹路径获取年份信息
                match = re.match(r'.*([12]\d{3}).*', os.path.basename(root))
                if match:
                    year = match.group(1)
                    if (start_year is not None and int(year) < int(start_year)) or (end_year is not None and int(year) > int(end_year)):
                        # 如果年份不符合要求，则移除该文件夹，防止遍历其中的文件
                        dirs[:] = []
                        continue

                # 遍历当前文件夹中的所有txt文件
                for filename in files:
                    if filename.endswith('.txt'):
                        # 解析文件名，提取股票代码、公司简称和年份
                        match = re.match(r'^(\d{6})_(.*?)_(\d{4})\.txt$', filename)
                        if match:
                            stock_code = match.group(1)
                            company_name = match.group(2)

                            # 提取关键词并统计词频和总字数
                            related_words_counts, keyword_counts, total_words = extract_keywords(os.path.join(root, filename), keywordsGroup)

                            # 将结果写入Excel表格
                            worksheet.write(row, 0, stock_code)
                            worksheet.write(row, 1, company_name)
                            worksheet.write(row, 2, year)
                            worksheet.write(row, 3, total_words)  # 写入总字数
                            for i, count in enumerate(keyword_counts):
                                worksheet.write(row, i + 4, count)  # 调整关键词列的索引
                            worksheet.write(row, j, related_words_counts) #写入关联词计数
                            row += 1

                            # 更新进度
                            processed_files += 1
                            progress = (processed_files / total_files) * 100
                            print(f"\r当前进度: {progress:.2f}%", end='', flush=True)

                            # 每处理指定数目个数据就保存一次Excel文件
                            if processed_files % size == 0:
                                workbook.save(name)
        except FileNotFoundError:
            print(f"文件夹不存在: {folder_path}")
        except PermissionError:
            print(f"没有访问权限: {folder_path}")
        except Exception as e:
            print(f"处理文件失败: {folder_path}")
            print(str(e))

        # 保存Excel文件
        try:
            workbook.save(name)
            print("\nExcel文件保存成功！")
        except FileNotFoundError:
            print(f"保存Excel文件失败: 文件夹不存在")
        except PermissionError:
            print(f"保存Excel文件失败: 没有访问权限")
        except Exception as e:
            print("\n保存Excel文件失败。")
            print(str(e))
    except FileNotFoundError:
        print(f"文件夹不存在: {folder_path}")
    except PermissionError:
        print(f"没有访问权限: {folder_path}")
    except Exception as e:
        print("处理文件失败！")
        print(str(e))


if __name__ == '__main__':
    # 设置要提取的关键词列表
    keywordsGroup = [
        ['负债率'],
        ['效率']    
    ]
    
    root_folder = f"{work_path}"

    # 输入处理结果的文件名
    name = "词频分析结果.xlsx"
    # 暂存数目大小，默认为100，尽量别改太小，否则IO压力很大。
    size = 100
    # 处理文件夹中的所有txt文件，并将结果存储到Excel表格中
    try:
        if start_year > end_year:
            print("起始年份不能大于中止年份！！！！！")
        else:
            process_files(root_folder, keywordsGroup, start_year, end_year)
    except Exception as e:
        print("文件处理失败！！")
        print(str(e))

    #！！！注意：如果程序运行无反应，多半是路径和txt文件命名问题！
    # 推荐文件名命名格式：“600519_贵州茅台_2019.txt”
