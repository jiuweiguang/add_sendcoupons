# -*- coding:utf-8 -*-
# @Time : 2020/3/24 11:07
# @Author: lsj
# @File : add_sendcoupons.py
import time,datetime
import xlrd

#服务费抵扣券sql插入语句
def insert_service_coupons_sql(couponsid,consid,money,number):
    time_now = time.strftime('%Y-%m-%d %H:%M:%S',time.localtime(time.time())) #获取当前时间
    now = datetime.datetime.now()
    delta = datetime.timedelta(days=365)
    n_days = now + delta
    add_time_year=n_days.strftime('%Y-%m-%d %H:%M:%S').split('.')[0] #新增一年时间
    txt_file = open('./coupons_sql_out.txt','a+')
    for i in range(0,number):
        try:
            sql_str="insert `c_coupons`(`COUPON_ID`,`CONS_ID`,`SERIAL_NO`,`GRANT_TIME`,`START_TIME`,`END_TIME`,`STATUS`,`money`,`min_bill_amount`,`coupon_type`,`recharge_id`,`deduct_range`,`use_type`) " \
            "values("+ couponsid + ","+ consid + ",0,'" + time_now + "'," + "'" + time_now + "'" + ",'" + add_time_year + "',0,"+ money + "," + money + ",1,-1,'2','1');"
            txt_file.write(sql_str)
            txt_file.write('\n')
        except Exception as e:
            print(e)
    txt_file.write('\n')
    txt_file.close()


def coupons_data():
    file = xlrd.open_workbook('./优惠券sql导出数据.xlsx')
    sheet_d = file.sheets()[0]
    nrows =sheet_d.nrows
    data = []
    try:
        for i in range(1,nrows):
            data_dist = {"cons_id": "",  # cons_id
                              "coupons_id": "",  # coupons_id
                              "server_5": "",  # 5元服务费抵扣券数量
                              "server_10": "",  # 10元服务费抵扣券数量
                              "fee_5": "",  # 5元礼品券数量
                              "fee_10": "",  # 10元礼品券数量
                              "fee_15": "",  # 15元礼品券数量
                              "fee_20": "",  # 20元礼品券数量
                              "fee_25": ""  # 25元礼品券数量
                               }
            data_dist['cons_id'] = int(sheet_d.cell(i, 0).value)
            data_dist['coupons_id'] = int(sheet_d.cell(i, 1).value)
            data_dist['server_5'] = int(sheet_d.cell(i, 2).value)
            data_dist['server_10'] = int(sheet_d.cell(i, 3).value)
            data_dist['fee_5'] = int(sheet_d.cell(i, 4).value)
            data_dist['fee_10'] = int(sheet_d.cell(i, 5).value)
            data_dist['fee_15'] = int(sheet_d.cell(i, 6).value)
            data_dist['fee_20'] = int(sheet_d.cell(i, 7).value)
            data_dist['fee_25'] = int(sheet_d.cell(i, 8).value)
            data.append(data_dist)
        for i in data:
            if i['server_5']!=0:
                insert_service_coupons_sql(couponsid=str(i['coupons_id']),consid=str(i['cons_id']),money='5',number=i['server_5'])
            if i['server_10']!=0:
                insert_service_coupons_sql(couponsid=str(i['coupons_id']), consid=str(i['cons_id']), money='10',
                                           number=i['server_10'])
            if i['fee_5']!=0:
                insert_service_coupons_sql(couponsid=str(i['coupons_id']), consid=str(i['cons_id']), money='5',
                                           number=i['fee_5'])
            if i['fee_10']!=0:
                insert_service_coupons_sql(couponsid=str(i['coupons_id']), consid=str(i['cons_id']), money='10',
                                           number=i['fee_10'])
            if i['fee_15']!=0:
                insert_service_coupons_sql(couponsid=str(i['coupons_id']), consid=str(i['cons_id']), money='15',
                                           number=i['fee_15'])
            if i['fee_20']!=0:
                insert_service_coupons_sql(couponsid=str(i['coupons_id']), consid=str(i['cons_id']), money='20',
                                           number=i['fee_20'])
            if i['fee_25'] != 0:
                insert_service_coupons_sql(couponsid=str(i['coupons_id']), consid=str(i['cons_id']), money='25',
                                           number=i['fee_25'])
    except Exception as e:
        print(e)
if __name__=='__main__':
    coupons_data()