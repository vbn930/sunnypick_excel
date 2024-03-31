#python version 3.10.11
#used lib : pandas

import pandas as pd
import shutil
import os
from dataclasses import dataclass
from datetime import datetime
import time

@dataclass
class OrderData:
    order_date: str
    shop_name: str
    dealer_name: str
    brand_name: str
    item_name: str #상품명 - 옵션명(송장출력명)
    quantity: int
    order_num: str
    delivery_name: str
    delivery_num: str
    reciever_name: str
    reciever_phone: str
    selling_price: int
    supply_price: int #매출단가 - 공급단가
    total_supply_price: int #공급단가 총계
    delivery_fee: int
    extra_delivery_fee: int

def get_duration_from_csv():
    data = pd.read_csv("./setting.csv")
    start_date = data["시작 날짜"][0]
    end_date = data["종료 날짜"][0]

    return start_date, end_date

def get_spilted_list(org_data):
    dst_datas = []
    dst_data = []

    for i in range(len(org_data)):
        if i == 0:
            dst_data.append(org_data[i])
            if len(org_data) == 1:
                dst_datas.append(dst_data)
        else:
            if org_data[i-1].shop_name != org_data[i].shop_name:
                dst_datas.append(dst_data)
                dst_data = []
            dst_data.append(org_data[i])
    dst_datas.append(dst_data)
    return dst_datas

def get_input_data(file_path, start_date, end_date):
    order_datas = []
    datas = pd.read_excel(file_path)
    for idx, row in datas.iterrows():
        date = row["날짜"]
        if start_date <= date and date <= end_date:
            order_data = OrderData(row["날짜"], row["쇼핑몰명"], row["사업자명"], row["브랜드"], row["상품명"][:-5], row["수량"], 
                row["주문번호"], row["택배사"], row["운송장번호"], row["수취인명"], row["수취인휴대폰"], row["판매가"], row["매출단가"], row["매출총계"], row["결제배송비"], row["추가배송비"])
            order_datas.append(order_data)
    order_datas.sort(key=lambda x: x.shop_name)
    return  get_spilted_list(order_datas)

def get_shop_info(shop_data):

    shop_data.sort(key=lambda x: x.delivery_num)

    #운송장 번호가 일치하는 주문은 배송비 한번만 부과
    for i in range(len(shop_data)):
        if i != 0:
            if shop_data[i - 1].delivery_num == shop_data[i].delivery_num:
                shop_data[i].delivery_fee = 0
                shop_data[i].extra_delivery_fee = 0
    
    #브랜드 별로 각각 금액과 수량 계산
    shop_data.sort(key=lambda x: x.brand_name)
    total_price = 0
    total_delivery_fee = 0
    total_extra_delivery_fee = 0
    for data in shop_data:
        total_price += data.total_supply_price
        total_delivery_fee += data.delivery_fee
        total_extra_delivery_fee += data.extra_delivery_fee
    print(total_price, total_delivery_fee, total_extra_delivery_fee)

def get_brand_info(shop_data):
    brand_datas = get_spilted_list(shop_data)
    
def get_item_info(brand_data):
    item_datas = get_spilted_list(brand_data)


def main():
    start_date, end_date = get_duration_from_csv()
    start_date = datetime.strptime(start_date, "%Y-%m-%d")
    end_date = datetime.strptime(end_date, "%Y-%m-%d")

    shop_datas = get_input_data("./input.xlsx", start_date, end_date)
    for shop_data in shop_datas:
        print(shop_data[0].shop_name, len(shop_data))
        get_shop_info(shop_data)
    

if __name__ == "__main__":
    main()