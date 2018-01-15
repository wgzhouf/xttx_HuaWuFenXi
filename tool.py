import csv
import xlrd



def read_csv(file_path, row_readstart=0):
    """
    读取 csv，如果存在表头可以设置开始读取行数来略过
    :param file_path: 文件路径
    :param row_readstart: 开始行数
    :return:数组
    """
    with open(file_path) as csvfile:
        csvData = csv.reader(csvfile)
        dataGrid = []
        count = 0  # 记录行数用

        for row in csvData:
            # 计数器先加1，以符合人的数数习惯
            count += 1
            if count <= row_readstart:
                continue
            dataGrid.append(row)
    return dataGrid


def read_xls(file_path, sheet_name, row_readstart=0):
    """

    :param file_path:
    :param row_readstart:
    :return:
    """
    workbook = xlrd.open_workbook(file_path)
    worksheet1 = workbook.sheet_by_name(sheet_name)
    num_rows = worksheet1.nrows

    dataGrid = []
    count = 0  # 记录行数用

    for curr_row in range(num_rows):
        count += 1
        if count <= row_readstart:
            continue
        row = worksheet1.row_values(curr_row)
        dataGrid.append(row)

    return dataGrid


def toNumber(str, need_xiaoshudian=False):
    """
    将字符类型数字转换成数字类型数字
    :param str: 字符类型数字
    :param need_xiaoshudian 是否要保留小数点 默认 不要
    :return: 数字类型数字
    """
    number = 0
    if (not need_xiaoshudian):
        if '.' in str:
            number = str[0:str.index('.')]
            if(number != ''):
                number = int(number)
            else:
                number = 0
        else:
            number = int(str)
    else:
        if '.' in str:
            number = float(str)
        else:
            number = int(str)
    return number


def formatDate(str):
    """
    将 20180108 转换成 2018年1月8日
    :param str:
    :return:
    """
    # 截取年
    year = str[0:4]
    # 截取月
    month = str[4:6]
    # 截取日
    day = str[6:]

    return year + '年' + month + '月' + day + '日'
