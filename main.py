from tool import read_csv
from tool import read_xls
from business import meiri_yonghu_huawuhuoyue
from business import dianhuahuiyi_shiyong_tongji
from business import dianhuahuiyi_shiyong_fenbu
from business import gesheng_yonghu_xinzeng_huoyue_fenbu


def main():
    # 黄德伟提供的文件
    meiri_yonghu_huawuhuoyue_origin_data_xttx = read_csv(r'C:\Users\ecp\Desktop\运行周报\20180105_20180111协同通信每日活跃数据统计.csv', 1)
    meiri_yonghu_huawuhuoyue_origin_data_yx_zhuanxian = read_csv(r'C:\Users\ecp\Desktop\运行周报\通话20180105_0111.csv', 0)
    meiri_yonghu_huawuhuoyue_origin_data_yx_duoren = read_csv(r'C:\Users\ecp\Desktop\运行周报\会议20180105_0111.csv', 0)
    meiri_yonghu_huawuhuoyue_origin_data_xttx_huoyue = read_csv(r'C:\Users\ecp\Desktop\运行周报\20180105-20180111日协同通信活跃情况.csv', 1)

    # 明师傅提供的文件
    meiri_yonghu_huawuhuoyue_origin_data_xttx_it_huoyue = read_xls(r'C:\Users\ecp\Desktop\运行周报\全国&浙江开销户工单统计(20180105~0111).xlsx', '开销户-按产品', 4)

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

    gesheng_yonghu_xinzeng_huoyue_fenbu(r'C:\Users\ecp\Desktop\运行周报\program\xttx_huawufenxi.xls',
                                        meiri_yonghu_huawuhuoyue_origin_data_xttx_huoyue,
                                        meiri_yonghu_huawuhuoyue_origin_data_xttx_it_huoyue)


if __name__ == '__main__':
    main()
