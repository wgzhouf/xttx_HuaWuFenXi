from tool import read_csv
from tool import read_xls
from business import meiri_yonghu_huawuhuoyue
from business import dianhuahuiyi_shiyong_tongji
from business import dianhuahuiyi_shiyong_fenbu


def main():
    meiri_yonghu_huawuhuoyue_origin_data_xttx = read_csv(r'C:\Users\ecp\Desktop\运行周报\20180105_20180111协同通信每日活跃数据统计.csv', 1)
    meiri_yonghu_huawuhuoyue_origin_data_yx_zhuanxian = read_csv(r'C:\Users\ecp\Desktop\运行周报\通话20180105_0111.csv', 0)
    meiri_yonghu_huawuhuoyue_origin_data_yx_duoren = read_csv(r'C:\Users\ecp\Desktop\运行周报\会议20180105_0111.csv', 0)

    meiri_yonghu_huawuhuoyue(r'C:\Users\ecp\Desktop\运行周报\program\xttx_huawufenxi.xls',
                             meiri_yonghu_huawuhuoyue_origin_data_xttx,
                             meiri_yonghu_huawuhuoyue_origin_data_yx_zhuanxian,
                             meiri_yonghu_huawuhuoyue_origin_data_yx_duoren)

    dianhuahuiyi_shiyong_tongji(r'C:\Users\ecp\Desktop\运行周报\program\xttx_huawufenxi.xls',
                                meiri_yonghu_huawuhuoyue_origin_data_xttx,
                                meiri_yonghu_huawuhuoyue_origin_data_yx_duoren)

    dianhuahuiyi_shiyong_fenbu(r'C:\Users\ecp\Desktop\运行周报\program\xttx_huawufenxi.xls',
                                meiri_yonghu_huawuhuoyue_origin_data_xttx,
                                meiri_yonghu_huawuhuoyue_origin_data_yx_duoren)


if __name__ == '__main__':
    main()
