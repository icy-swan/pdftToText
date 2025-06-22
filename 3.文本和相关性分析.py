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
# import difflib
import openpyxl
import re

def clean_excel_text(text):
    """
    移除Excel不允许的控制字符（ASCII 0-31中除\t\n\r外的所有字符）
    """
    if not isinstance(text, str):
        return text
    
    # 正则表达式：匹配所有非法控制字符
    cleaned = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f]', '', text)
    return cleaned


# 输入年份区间
start_year = "2013"
end_year = "2013"
# 相关性计算，多少个词内
steps = [5, 10, 15]
# 文件存储路径
work_path="/Users/bl/git/pdftToText/reports"

# 在指定步长里找是否匹配
def check_key_in_else_target(idx, step, keywordsgroup_index):
    result = 0

    for i, target in enumerate(keywordsgroup_index):
        if i != 0:
            found = False
            # 在n步之内找正向index
            for x in range(step):
                next_idx = idx + x + 1
                # 如果找到，设置为true，不继续循环
                try:
                    if target.index(next_idx):
                        found = True
                        break
                except Exception as e:
                    found = False
            # 正向没找到，去负向继续找
            if found == False:
                for x in range(step):
                    next_idx = idx - x - 1
                    # 如果找到，设置为true，不继续循环
                    try:
                        if target.index(next_idx):
                            found = True
                            break
                    except Exception as e:
                        found = False
            # 如果一个idx在第一个目标查找词内就没找到，就不用继续遍历了
            if found == False:
                break
            # 如果找到了就继续找下一个
            # 如果找到最后一个匹配也是能找到，就设置数
            if (i == len(keywordsgroup_index) - 1 & found == True):
                result = 1

    return result

# 计算匹配值
def count_relative(keywordsgroup_index, steps):
    # 统计总相关字出现字数
    related_words_counts = [0] * len(steps)
    # 目标遍历数组，其余数组对其进行比较
    target = keywordsgroup_index[0]
    # cur_i = 1
    # length = len(keywordsgroup_index)
    # 遍历
    for idx in target:
        # 根据不同step进行查找
        for i, step in enumerate(steps):
            check_step_count = check_key_in_else_target(idx, step, keywordsgroup_index)
            related_words_counts[i] = related_words_counts[i] + check_step_count

    return  related_words_counts

# 抽取关键字
def extract_keywords(keywordsGroup, root, file_origin_name):
    filename = os.path.join(root, file_origin_name)

    # 每个分类单独计数
    keywordsgroup_counts = [0] * len(keywordsGroup)
    keywordsgroup_index = [[] for i in range(len(keywordsGroup))]

    # 统计总字数
    total_words = 0
    # 抽取后的文件的存放路径
    tempPath= os.path.join(root, f"temp")
    try:
        with open(filename, 'r', encoding='utf-8') as f:
            content = f.read()

        # 将多个分组的关键词，全部添加到自定义词典
        for keywords in keywordsGroup:
            for keyword in keywords:
                jieba.add_word(keyword)

        # 使用jieba库进行分词
        words = jieba.cut(content)
        words = [word for word in words if word.strip()]
        # 统计总字数
        total_words = len(words)


        #保存下分词后的数据，主要用于过程检测
        workbook = openpyxl.Workbook()
        worksheet = workbook.active
        worksheet.title = "Alice"
        worksheet.append(["分词结果"])
        try:
            for w in words:
                worksheet.append([clean_excel_text(w)])
        except Exception as e:
            print(str(e))

        try:
            os.makedirs(tempPath, exist_ok=True)
            temp_file_path = f"{tempPath}/{file_origin_name}分词结果.xlsx"
            workbook.save(temp_file_path)
        except Exception as e:
            print(f"创建文件路径错误: {temp_file_path}")

        # 1. 生产一个key与index的字典
        words_dict = {}
        for idx, item in enumerate(words):
            if item not in words_dict:
                words_dict[item]=[]
            words_dict[item].append(idx)
        # 2. 进行遍历匹配
        for j, keywords in enumerate(keywordsGroup):
            for keyword in keywords:
                # 关键词创建正则表达式，与分词的每个词进行匹配
                # 拥有解决“负债率”和“企业负债率”不一致的问题
                # 如果不考虑多词组的关联计算，其实可以在分词时使用search分词，可以直接拆细词
                # 但是考虑关联，需要计算词的距离，所以无法search分词
                patten = rf"{keyword}"
                for word in words_dict:
                    if(re.search(patten, word)):
                        match_item = words_dict[word]
                        # 找到，该分类的统计数据+1
                        keywordsgroup_counts[j] = keywordsgroup_counts[j] + len(match_item)
                        # 记录该分类的index
                        keywordsgroup_index[j].extend(match_item)
        # 进行关联计算
        # 将第一个关键词组依次找出顺序，并于剩余关键词组内找寻相关顺序step步长内是否都有关联
        related_words_counts = count_relative(keywordsgroup_index, steps)

    except FileNotFoundError:
        print(f"文件不存在: {filename}")
    except PermissionError:
        print(f"没有访问权限: {filename}")
    except Exception as e:
        print(f"从文件中获取关键词失败: {filename}", e)
        print(str(e))

    return related_words_counts, keywordsgroup_counts, total_words


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

        for i, keyword in enumerate(keywordsGroup):
            worksheet.write(row, i + 4, f'关键词类别-{i}的次数')  # 按类型进行计数统计
        for i, step in enumerate(steps):
            worksheet.write(row, i + 4 + len(keywordsGroup), f'关联词step为{step}的次数') #写入关联词计数
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
                            related_words_counts, keywordsgroup_counts, total_words = extract_keywords(keywordsGroup, root, filename)

                            # 将结果写入Excel表格
                            worksheet.write(row, 0, stock_code)
                            worksheet.write(row, 1, company_name)
                            worksheet.write(row, 2, year)
                            worksheet.write(row, 3, total_words)  # 写入总字数
                            for i, count in enumerate(keywordsgroup_counts):
                                worksheet.write(row, i + 4, count)  # 调整关键词列的索引
                            for i, count in enumerate(related_words_counts):
                                worksheet.write(row, i + 4 + len(keywordsgroup_counts), count) #写入关联词计数

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
        [
            "1973年石油危机", "海上钻井", "西德克萨斯中质原油", "濒危物种", "天然气价格",
            "能源", "气候变化", "石油平台", "海上钻进平台", "西德克萨斯中质原油",
            "WTI", "能源安全", "石油出口", "石油", "油", "原油", "石油输出国组织",
            "OPEC", "美国石油协会", "API", "水力压裂", "石油供应", "能源效率",
            "污染", "碳税", "绿色能源", "太阳能", "光伏发电", "光伏技术", "环境",
            "太阳能", "乙醇燃料混合物", "内燃机", "页岩油", "汽油价格", "测井",
            "电动汽车", "天然气", "可持续能源", "温室气体", "替代能源", "能源行业",
            "风力", "能源冲击", "自然资源", "碳排放", "化石燃料", "油页岩", "能源危机",
            "石油出口禁令", "COGIS", "走向绿色", "环保化", "绿色发展", "石油储集层",
            "能源价格冲击", "油井", "钻井泥浆", "碳氢化合物", "烃", "残油", "渣油",
            "能源价格冲击", "管道", "管线", "能源市场", "液化石油气", "可持续性",
            "能源不安全", "太阳能电池", "乙醇价格", "石油和天然气", "风能", "碳足迹",
            "全球变暖", "石油储量", "COGCC", "混合动力汽车", "石油工业", "石油行业",
            "清洁能源法案", "能源价格波动", "定向钻井", "液化天然气", "可再生能源",
            "玉米乙醇", "石油危机", "能源独立", "压裂开采", "节能", "乙醇燃料",
            "石油价格", "油价", "勘探井", "能源税", "地热能", "石油", "布伦特原油",
            "汽油", "水平钻井", "探明储量", "清洁能源", "温室效应", "京都议定书",
            "太阳能发电", "太阳能", "堆肥", "煤油"
        ],
        [
            "风险", "经营风险", "市场风险", "信用风险", "不确定性", "不确定",
            "波动", "变化", "改变", "徘徊", "不稳", "不稳定性", "不稳定",
            "不寻常", "错综复杂", "非常复杂", "纷繁复杂", "纷纭复杂", "十分复杂",
            "结构复杂", "变得复杂", "风云变幻", "风云突变", "矛盾突出", "突变",
            "复杂", "复杂多变", "诡谲多变", "阵痛", "过渡", "问责", "整顿",
            "危险", "动荡", "动荡不安", "动荡不定", "多变性", "振荡下行", "震荡",
            "震荡不安", "政治波动", "难以确定", "难以预测", "难以预料", "难以捉摸",
            "接受考验", "混乱", "混乱状态", "有时", "时而", "随机", "冲击", "危机"
        ]  
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

