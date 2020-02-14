# coding:utf-8
# coding=gbk
import os
import json
import pandas
import logging
import codecs
import configparser
from datetime import datetime

import requests
from bs4 import BeautifulSoup


REQUEST_HEADER = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/79.0.3945.130 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,"
    "image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9",
    "Accept-Encoding": "gzip, deflate",
    "Accept-Language": "en,zh-TW;q=0.9,zh;q=0.8,en-US;q=0.7"
}


def load_config_file():
    config_filename = "config.ini"
    if os.path.exists(config_filename):
        config = configparser.ConfigParser()
        config.readfp(codecs.open(config_filename, "r", encoding="utf-8-sig"))
        return config
    else:
        raise FileNotFoundError("Config file not found -> config.ini")


def read_search_keywords_from_config_as_list():
    try:
        config = load_config_file()
        logging.debug(config)
        return config["default"]["search_keywords"].split(",")
    except Exception as e:
        logging.exception(e)
        raise ValueError("Search keywords format error")


def save_to_xlsx(xls_file, data):
    df = pandas.DataFrame(data).T
    df.to_excel(xls_file)


def get_today_date_string():
    today_string = datetime.strftime(datetime.now(), '%Y/%m/%d')
    year, month, day = today_string.split("/")
    tw_year = int(year) - 1911
    tw_date_string = "{}/{}/{}".format(tw_year, month, day)
    return tw_date_string


def get_start_date():
    config = load_config_file()
    start_date = config["default"]["search_start_date"]
    if not start_date:
        start_date = get_today_date_string()
    return start_date


def get_end_date():
    config = load_config_file()
    end_date = config["default"]["search_end_date"]
    if not end_date:
        end_date = get_today_date_string()
    return end_date


def gov_bid_parser(search_keyword):
    base_url = "http://web.pcc.gov.tw/tps"
    url = "{}/pss/tender.do?searchMode=common&searchType=basic".format(
        base_url)
    search_start_date = get_start_date()
    search_end_date = get_end_date()
    print("Search Date From [{}] - [{}]".format(
        search_start_date, search_end_date))
    params = {
        "method": "search",
        "searchMethod": "true",
        "tenderUpdate": "",
        "searchTarget": "",
        "orgName": "",
        "orgId": "",
        "hid_1": "1",
        "tenderName": search_keyword,
        "tenderId": "",
        "tenderType": "tenderDeclaration",
        "tenderWay": "1,2,3,4,5,6,7,10,12",
        "tenderDateRadio": "on",
        "tenderStartDateStr": search_start_date,
        "tenderEndDateStr": search_end_date,
        "tenderStartDate": search_start_date,
        "tenderEndDate": search_end_date,
        "isSpdt": "N",
        "proctrgCate": "",
        "btnQuery": "查詢",
        "hadUpdated": ""
    }
    resp = requests.post(url, headers=REQUEST_HEADER, data=params)
    dom = BeautifulSoup(resp.text, "lxml")
    tr_tags = dom.find_all("tr", {"onmouseover": "overcss(this);"})
    title_mapping = {1: "機關名稱", 2: "標案名稱",
                     4: "招標方式", 6: "公告日期",
                     7: "截止投標", 8: "預算金額", 9: "連結"}
    total_bids = {}
    counts = 0
    for tr_tag in tr_tags:
        td_tags = tr_tag.find_all("td")
        items = {}
        for i, td_tag in enumerate(td_tags):
            link = ""
            td_text = td_tag.text.strip().strip("\n")
            if td_text and i in title_mapping:
                link_tag = td_tag.find("a")
                text = td_text
                if link_tag:
                    link = link_tag["href"].replace("..", base_url)
                    text = link_tag.text.strip()

                items.update({title_mapping[i]: text})
                if link and i == 2:
                    items.update({title_mapping[i + 7]: link})

        total_bids.update({"{} {}".format(search_keyword, counts): items})
        counts += 1

    return total_bids


if __name__ == "__main__":
    search_list = read_search_keywords_from_config_as_list()
    summary_dict = {}
    for keyword in search_list:
        print("Try to search keywords: {}".format(keyword))
        result_dict = gov_bid_parser(keyword)
        summary_dict.update(result_dict)
    save_to_xlsx("gov_bids.xlsx", summary_dict)
    os.system("pause")
