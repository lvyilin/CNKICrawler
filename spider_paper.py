# -*- coding: utf-8 -*-
import socket
import time
import urllib
from configparser import ConfigParser

import requests
import xlwt
from bs4 import BeautifulSoup


def get_keyword(paper_id):
    if paper_id is None:
        return None
    cnki_net_url = "http://kns.cnki.net/KCMS/detail/detail.aspx?dbcode=CJFQ&filename=" + paper_id
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/75.0.3770.100 Safari/537.36"}
    html = requests.get(cnki_net_url, headers=headers)
    soup = BeautifulSoup(html.text, 'html.parser')
    txt = soup.find(id='catalog_KEYWORD').parent
    text_split = str(txt.get_text()).replace("关键词：", "").replace(";", " ").split()
    return " ".join(text_split)


def parse_url_id(url: str):
    id_start = url.index("CJFDTOTAL-") + len("CJFDTOTAL-")
    id_end = url.index(".htm")
    return url[id_start:id_end]


def spider_paper():
    start = time.clock()
    # f=urllib2.urlopen(url, timeout=5).read()
    # soup=BeautifulSoup(html)
    # tags=soup.find_all('a')
    file = open("data-detail.txt", encoding='utf8')
    cf = ConfigParser()
    cf.read("Config.conf", encoding='utf-8')
    keyword = cf.get('base', 'keyword')  # 关键词

    # 写入Excel
    wb = xlwt.Workbook("data_out.xls")
    sheet = wb.add_sheet("data-out")
    sheet.write(0, 0, '下载网址')
    sheet.write(0, 1, '标题')
    sheet.write(0, 2, '来源')
    sheet.write(0, 3, '引用')
    sheet.write(0, 4, '作者')
    sheet.write(0, 5, '作者单位')
    sheet.write(0, 6, '关键词')
    # sheet.write(0, 7, '摘要')
    # sheet.write(0, 8, '共引文献')

    raw_text_file = open("data_out.txt", 'w', encoding='utf8')

    lines = file.readlines()
    txt_num = 1
    lin_num = 1
    paper_list = []
    for line in lines:
        try:
            object = line.split('\t')
            paper_url = object[0]
            if paper_url in paper_list:
                continue
            paper_list.append(paper_url)
            attempts = 0
            success = False
            while attempts < 3 and not success:
                try:
                    html = urllib.request.urlopen(paper_url).read()
                    soup = BeautifulSoup(html, 'html.parser')
                    socket.setdefaulttimeout(10)  # 设置10秒后连接超时
                    success = True
                except socket.error:
                    attempts += 1
                    print("第" + str(attempts) + "次重试！！")
                    if attempts == 3:
                        break
                except urllib.error:
                    attempts += 1
                    print("第" + str(attempts) + "次重试！！")
                    if attempts == 3:
                        break
            title = soup.find_all('div',
                                  style="text-align:center; width:740px; font-size: 28px;color: #0000a0; font-weight:bold; font-family:'宋体';")
            abstract = soup.find_all('div', style='text-align:left;word-break:break-all')
            author = soup.find_all('div', style='text-align:center; width:740px; height:30px;')

            # 获取作者名字
            for item in author:
                author = item.get_text()
            # 获取作者单位，处理字符串匹配
            authorUnitScope = soup.find('div', style='text-align:left;', class_='xx_font')
            author_unit = ''
            author_unit_text = authorUnitScope.get_text()
            # print(author_unit_text)
            if '【作者单位】：' in author_unit_text:
                auindex = author_unit_text.find('【作者单位】：', 0)
            else:
                auindex = author_unit_text.find('【学位授予单位】：', 0)
            for k in range(auindex, len(author_unit_text)):
                if author_unit_text[k] == '\n' or author_unit_text[k] == '\t' or author_unit_text[k] == '\r' or \
                        author_unit_text[k] == '】':
                    continue
                if author_unit_text[k] == ' ' and author_unit_text[k + 1] == ' ':
                    continue
                if author_unit_text[k] != '【':
                    author_unit = author_unit + author_unit_text[k]
                if author_unit_text[k] == '【' and k != auindex:
                    break
            author_unit = author_unit.replace("作者单位：", "").replace(";", " ")

            key_word = get_keyword(parse_url_id(paper_url))
            if key_word is None:
                continue

            line = line.strip('\n')
            line = line + '\t' + str(author) + '\t' + str(author_unit) + '\t' + str(key_word) + '\n'
            outstring = line.split('\t')
            for i in range(len(outstring)):
                sheet.write(lin_num, i, outstring[i])
            print('写入第' + str(lin_num) + '行')
            raw_text_file.write("|".join(outstring))
            lin_num += 1
            wb.save('data_out_' + str(keyword) + '.xls')
        except Exception as e:
            print(e)
            continue

    raw_text_file.close()
    file.close()
    end = time.clock()
    print('Running time: %s Seconds' % (end - start))


if __name__ == '__main__':
    spider_paper()
