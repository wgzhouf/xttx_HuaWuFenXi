import csv
import xlrd

code_area_map = {
    '10': '北京',
    '20': '广东',
    '21': '上海',
    '22': '天津',
    '23': '重庆',
    '24': '辽宁',
    '25': '江苏',
    '27': '湖北',
    '28': '四川',
    '29': '陕西',
    '311': '河北',
    '351': '山西',
    '371': '河南',
    '431': '吉林',
    '451': '黑龙江',
    '471': '内蒙古',
    '531': '山东',
    '551': '安徽',
    '571': '浙江',
    '591': '福建',
    '731': '湖南',
    '771': '广西',
    '791': '江西',
    '851': '贵州',
    '871': '云南',
    '891': '西藏',
    '898': '海南',
    '931': '甘肃',
    '951': '宁夏',
    '971': '青海',
    '991': '新疆'

}

area_code_map = {
    '北京': '10',
    '广东': '20',
    '上海': '21',
    '天津': '22',
    '重庆': '23',
    '辽宁': '24',
    '江苏': '25',
    '湖北': '27',
    '四川': '28',
    '陕西': '29',
    '河北': '311',
    '山西': '351',
    '河南': '371',
    '吉林': '431',
    '黑龙江': '451',
    '内蒙古': '471',
    '山东': '531',
    '安徽': '551',
    '浙江': '571',
    '福建': '591',
    '湖南': '731',
    '广西': '771',
    '江西': '791',
    '贵州': '851',
    '云南': '871',
    '西藏': '891',
    '海南': '898',
    '甘肃': '931',
    '宁夏': '951',
    '青海': '971',
    '新疆': '991'

}

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

    return month + '月' + day + '日'
