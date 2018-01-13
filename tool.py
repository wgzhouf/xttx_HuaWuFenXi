import csv


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
