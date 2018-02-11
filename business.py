import xlrd
import xlutils.copy

from tool import toNumber, formatDate
from tool import area_code_map,code_area_map


def meiri_yonghu_huawuhuoyue(filepath, data_xttx, data_yx_zhuanxian, data_yx_duoren):
    """
    运行周报 4、本期间每日用户话务、活跃统计
    :param filepath:xttx_huawufenxi.xls 路径
    :param data_xttx: 20180105_20180111协同通信每日活跃数据统计.csv 数据
    :param data_yx_zhuanxian:通话20180105_0111.csv
    :param data_yx_duoren:会议20180105_0111.csv
    :return:
    """
    rb = xlrd.open_workbook(filepath)
    wb = xlutils.copy.copy(rb)

    """话务总计数据"""
    total_xttx_tonghua = 0
    total_xttx_duanxin = 0
    total_xttx_IM = 0
    total_yx_zhuanxian = 0
    total_yx_duoren = 0
    total_yx_zonghuawu = 0


    # 读取 sheet “meiri_yonghu_huawuhuoyue”
    ws = wb.get_sheet(0)
    row_pointer = 2
    for xttx_crr_row in data_xttx:
        """易信专线电话部分"""
        count_yx_zhuanxian_fenzhong = 0
        for yx_crr_row in data_yx_zhuanxian:
            if yx_crr_row[0] == xttx_crr_row[0]:
                count_yx_zhuanxian_fenzhong += toNumber(yx_crr_row[2])
            else:
                continue
        # 易信专线电话时长
        ws.write(row_pointer, 5, count_yx_zhuanxian_fenzhong)
        # 累加 专线总时长
        total_yx_zhuanxian += count_yx_zhuanxian_fenzhong

        count_yx_duoren_fenzhong = 0
        for yx_crr_row in data_yx_duoren:
            if yx_crr_row[0] == xttx_crr_row[0]:
                count_yx_duoren_fenzhong += toNumber(yx_crr_row[2])
            else:
                continue
        # 易信多人电话时长
        ws.write(row_pointer, 6, count_yx_duoren_fenzhong)
        # 累加 多人总时长
        total_yx_duoren += count_yx_duoren_fenzhong

        # 易信总话务
        ws.write(row_pointer, 7, count_yx_duoren_fenzhong+count_yx_zhuanxian_fenzhong)
        # 累加 总话务时长
        total_yx_zonghuawu += count_yx_duoren_fenzhong+count_yx_zhuanxian_fenzhong

        """协同通信部分"""
        # 日期
        ws.write(row_pointer, 0, formatDate(xttx_crr_row[0]))

        # 协同通信 通话时长 要减去易信部分的数据
        xttx_temp = toNumber(xttx_crr_row[1])+toNumber(xttx_crr_row[2])-count_yx_duoren_fenzhong-count_yx_zhuanxian_fenzhong
        ws.write(row_pointer, 1, xttx_temp)
        total_xttx_tonghua += xttx_temp

        # 短信发送量
        ws.write(row_pointer, 2, xttx_crr_row[5])
        total_xttx_duanxin += toNumber(xttx_crr_row[5])

        # IM发送量
        ws.write(row_pointer, 3, xttx_crr_row[6])
        total_xttx_IM += toNumber(xttx_crr_row[6])

        # 活跃用户
        ws.write(row_pointer, 4, xttx_crr_row[7])

        row_pointer += 1

    ws.write(row_pointer, 1, total_xttx_tonghua)
    ws.write(row_pointer, 2, total_xttx_duanxin)
    ws.write(row_pointer, 3, total_xttx_IM)
    ws.write(row_pointer, 5, total_yx_zhuanxian)
    ws.write(row_pointer, 6, total_yx_duoren)
    ws.write(row_pointer, 7, total_yx_zonghuawu)

    wb.save(filepath)


def dianhuahuiyi_shiyong_tongji(filepath, data_xttx, data_yx_duoren):
    rb = xlrd.open_workbook(filepath)
    wb = xlutils.copy.copy(rb)

    """统计"""
    total_xttx_dhhy_shichang = 0
    total_xttx_dhhy_renshu = 0
    avg_xttx_dhhy_shichang = 0
    total_yx_duoren_shichang = 0
    total_yx_duoren_renshu = 0
    avg_yx_duoren_shichang = 0

    # 读取 sheet “dianhuahuiyi_shiyong_tongji”
    ws = wb.get_sheet(1)
    row_pointer = 2
    for xttx_crr_row in data_xttx:

        count_yx_duoren_fenzhong = 0
        count_yx_duoren_renshu = 0
        for yx_crr_row in data_yx_duoren:
            if yx_crr_row[0] == xttx_crr_row[0]:
                count_yx_duoren_fenzhong += toNumber(yx_crr_row[2])
                count_yx_duoren_renshu += toNumber(yx_crr_row[3])
            else:
                continue
        # 易信多人电话时长
        ws.write(row_pointer, 4, count_yx_duoren_fenzhong)
        # 易信多人电话使用人数
        ws.write(row_pointer, 5, count_yx_duoren_renshu)
        # 易信多人电话人均时长
        temp_avg_yx_renjun = round(count_yx_duoren_fenzhong/count_yx_duoren_renshu, 1)
        ws.write(row_pointer, 6, temp_avg_yx_renjun)

        # 累加 多人总时长，人数
        total_yx_duoren_shichang += count_yx_duoren_fenzhong
        total_yx_duoren_renshu += count_yx_duoren_renshu
        avg_yx_duoren_shichang += temp_avg_yx_renjun


        """协同通信部分"""
        # 日期
        ws.write(row_pointer, 0, formatDate(xttx_crr_row[0]))

        # 协同通信 通话时长 要减去易信部分的数据
        xttx_fenzhong = toNumber(xttx_crr_row[2])-count_yx_duoren_fenzhong
        xttx_renshu = toNumber(xttx_crr_row[3])-count_yx_duoren_renshu
        temp_avg_xttx_renjun = round(xttx_fenzhong/xttx_renshu, 1)
        ws.write(row_pointer, 1, xttx_fenzhong)
        ws.write(row_pointer, 2, xttx_renshu)
        ws.write(row_pointer, 3, temp_avg_xttx_renjun)

        total_xttx_dhhy_shichang += xttx_fenzhong
        total_xttx_dhhy_renshu += xttx_renshu
        avg_xttx_dhhy_shichang += temp_avg_xttx_renjun

        row_pointer += 1

    ws.write(row_pointer, 1, total_xttx_dhhy_shichang)
    ws.write(row_pointer, 2, total_xttx_dhhy_renshu)
    ws.write(row_pointer, 3, str(round(avg_xttx_dhhy_shichang/7, 1))+'（周平均）')
    ws.write(row_pointer, 4, total_yx_duoren_shichang)
    ws.write(row_pointer, 5, total_yx_duoren_renshu)
    ws.write(row_pointer, 6, str(round(avg_yx_duoren_shichang/7, 1))+'（周平均）')

    wb.save(filepath)


def dianhuahuiyi_shiyong_fenbu(filepath, data_xttx, data_yx_duoren):
    rb = xlrd.open_workbook(filepath)
    wb = xlutils.copy.copy(rb)

    # 读取 sheet “dianhuahuiyi_shiyong_fenbu”
    ws = wb.get_sheet(2)
    row_pointer = 2
    for xttx_crr_row in data_xttx:

        """易信部分"""
        count_yx_duoren_cishu = 0
        for yx_crr_row in data_yx_duoren:
            if yx_crr_row[0] == xttx_crr_row[0]:
                count_yx_duoren_cishu += toNumber(yx_crr_row[3])
            else:
                continue
        ws.write(row_pointer, 5, count_yx_duoren_cishu)

        """协同通信部分"""
        # 日期
        ws.write(row_pointer, 0, formatDate(xttx_crr_row[0]))
        ws.write(row_pointer, 1, xttx_crr_row[8])
        ws.write(row_pointer, 2, xttx_crr_row[9])
        ws.write(row_pointer, 3, xttx_crr_row[10])
        ws.write(row_pointer, 4, xttx_crr_row[11])
        row_pointer += 1
    wb.save(filepath)


def gesheng_yonghu_xinzeng_huoyue_fenbu(filepath, data_xttx, data_it):
    rb = xlrd.open_workbook(filepath)
    wb = xlutils.copy.copy(rb)

    ws = wb.get_sheet(3)

    data_it_grid = {}
    # 全国净增统计
    for row_it in data_it:
        temp_jingzeng = toNumber(str(row_it[9])) - toNumber(str(row_it[11]))
        data_it_grid[str(row_it[8])] = temp_jingzeng
    # 浙江净增统计
    i = 0
    zj_kaihu = 0
    zj_xiaohu = 0
    while True:
        if data_it[i][0] == '合计':
            zj_kaihu = toNumber(str(data_it[i][1]))
            break
        else:
            i += 1
    j = 0
    while True:
        if data_it[j][4] == '合计':
            zj_xiaohu = toNumber(str(data_it[j][5]))
            break
        else:
            j += 1

    data_it_grid['浙江'] = zj_kaihu - zj_xiaohu
    # data_it_grid['浙江'] = toNumber(str(data_it[2][1])) - toNumber(str(data_it[6][5]))

    xttx_zong_yonghu = 0
    xttx_zong_huoyue = 0
    xttx_zong_duanxin = 0
    row_pointer = 1
    for xttx_crr_row in data_xttx:
        if len(xttx_crr_row) == 0:
            continue
        # 省份
        ws.write(row_pointer, 0, code_area_map[xttx_crr_row[0]])
        # 净增用户
        ws.write(row_pointer, 1, data_it_grid[code_area_map[xttx_crr_row[0]]])
        # 总用户
        xttx_zong_yonghu += toNumber(xttx_crr_row[8])
        ws.write(row_pointer, 2, xttx_crr_row[8])
        # 活跃用户
        xttx_zong_huoyue += toNumber(xttx_crr_row[6])
        ws.write(row_pointer, 3, xttx_crr_row[6])
        # 短信业务
        ws.write(row_pointer, 4, xttx_crr_row[4])
        # 主叫话务量
        ws.write(row_pointer, 5, xttx_crr_row[1])
        # 电话会议
        ws.write(row_pointer, 6, xttx_crr_row[2])

        row_pointer += 1
    ws.write(row_pointer, 0, '小计')
    ws.write(row_pointer, 2, xttx_zong_yonghu)
    ws.write(row_pointer, 3, xttx_zong_huoyue)

    wb.save(filepath)


def zte_caps_caculate(filepath, data_zte):
    rb = xlrd.open_workbook(filepath)
    wb = xlutils.copy.copy(rb)

    ws = wb.get_sheet(4)

    data_grid = {}
    for row_zte in data_zte:
        if(row_zte[9] not in data_grid.keys()):
            data_grid[row_zte[9]] = {}
        else:
            if(row_zte[12] not in data_grid[row_zte[9]].keys()):
                data_grid[row_zte[9]][row_zte[12]] = toNumber(row_zte[14], True)
            else:
                data_grid[row_zte[9]][row_zte[12]] += toNumber(row_zte[14], True)

    row_pointer = 0
    result_grid = {}
    for day, day_grid in data_grid.items():
        max = 0
        for time, time_data in day_grid.items():
            if time_data >= max:
                max = time_data
        result_grid[day] = max

    for key, value in result_grid.items():
        ws.write(row_pointer, 0, key)
        ws.write(row_pointer, 1, value)
        row_pointer += 1

    wb.save(filepath)
