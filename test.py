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
        else:
            number = int(str)
    else:
        if '.' in str:
            number = float(str)
        else:
            number = int(str)
    return number

if __name__ == '__main__':
    area_code_map = {
        '北京': '10',
        '广东': '20'}
    area_code_map['上海'] = '21'
    print(area_code_map['北京'])
    print(area_code_map['上海'])