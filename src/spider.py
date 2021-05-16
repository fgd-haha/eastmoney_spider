import datetime
import json
import os
import re
import requests

from lxml import etree

from src.save.save_excel import save_to_excel
from src.save.save_json import save_to_json
from src.save.save_ppt import save_to_ppt

base_path = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ppt_output = f"{base_path}/result/ppt"
json_output = f"{base_path}/result/json"
excel_output = f"{base_path}/result/excel"


def north_homepage():
    rp = requests.get('http://data.eastmoney.com/hsgtcg/list.html')
    rp.encoding = 'gbk'
    html = etree.HTML(rp.text)

    data = html.xpath('//script')[16].text.replace('\r\n', '')
    result_list = json.loads(re.findall('"data": (\[.*])', data)[0])
    return result_list


def north_specific_page(page, date_str):
    params = {
        "type": "HSGT20_GGTJ_SUM",
        "token": "894050c76af8597a853f5b408b759f5d",
        "st": "ShareSZ_Chg_One",
        "sr": -1,
        "p": page,
        "ps": 6000,
        "js": "var%20EGbVpzom={pages:(tp),data:(x)}",
        "filter": f"(Market=%27003%27%20and%20DateType=%271%27%20and%20HdDate=%27{date_str}%27)",
        "rt": 53682120
    }

    response = requests.get('http://dcfm.eastmoney.com/EM_MutiSvcExpandInterface/api/js/get', params=params)

    result_list = json.loads(re.findall('data:(\[\{.*\}\])', response.text)[0])
    return result_list


def south_homepage():
    rp = requests.get('http://data.eastmoney.com/hsgtcg/lz.html')
    rp.encoding = 'gbk'
    html = etree.HTML(rp.text)

    data = html.xpath('//script')[17].text.replace('\r\n', '')
    result_list = json.loads(re.findall('"data": (\[.*])', data)[0])
    return result_list


def south_specific_page(page, date_str):
    params = {
        "type": "HSGTHDSTA",
        "token": "894050c76af8597a853f5b408b759f5d",
        "st": "HDDATE",
        "sr": -1,
        "p": page,
        "ps": 1000,
        "js": "var HwcYhlwj={pages:(tp),data:(x)}",
        "filter": "(MARKET='S')",
        "rt": 53796688
    }
    response = requests.get('http://dcfm.eastmoney.com/EM_MutiSvcExpandInterface/api/js/get', params=params)

    result_list = json.loads(re.findall('data:(\[\{.*\}\])', response.text)[0])
    return result_list


def deduplication(data_list, type="north", date_str=str(datetime.date.today())):
    result_dict = dict()
    if type == "north":
        for code in data_list:
            if code["SCode"] not in result_dict and code["DateType"] == "1":
                result_dict[code['SCode']] = code
    elif type == "south":
        for code in data_list:
            if code["HDDATE"].startswith(date_str):
                if code['SCODE'] not in result_dict:
                    result_dict[code['SCODE']] = code
                else:
                    print(result_dict[code['SCODE']], end="\n")
                    print(code)
                    print("*" * 100)

    return result_dict


def spider_north(date_str):
    # 爬虫获取数据
    north_data_list = []
    north_data_list += north_homepage()

    for i in range(2):
        north_data_list += north_specific_page(i + 1, date_str)
    # 对获得的数据去重并排序
    data_dict = deduplication(north_data_list)
    north_sorted_dict = sorted(data_dict.items(), key=lambda kv: (kv[1]["ShareSZ_Chg_One"]))
    return north_sorted_dict


def spider_south(date_str):
    # 爬虫获取数据
    south_data_list = []
    south_data_list += south_homepage()
    for i in range(2):
        south_data_list += south_specific_page(i + 1, date_str)
    # 对获得的数据去重并排序
    data_dict = deduplication(south_data_list, type="south", date_str=date_str)
    south_sorted_dict = sorted(data_dict.items(), key=lambda kv: (kv[1]['SHAREHOLDPRICEONE']))
    return south_sorted_dict


def run(date_str=str(datetime.date.today())):
    north_sorted_dict = spider_north(date_str)
    south_sorted_dict = spider_south(date_str)

    # 结果保存到json文件中
    save_to_json(json_output, north_sorted_dict)
    # 结果保存到excel文件中
    save_to_excel(excel_output, north_sorted_dict)
    # 结果保存到ppt文件中
    save_to_ppt(ppt_output, date_str, north_sorted_dict, south_sorted_dict)


if __name__ == '__main__':
    run(date_str="2021-02-19")
