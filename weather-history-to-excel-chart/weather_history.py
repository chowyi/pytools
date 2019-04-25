# coding:utf-8
import os
import sys
import time
import datetime
from collections import namedtuple

import requests
import xlsxwriter
from bs4 import BeautifulSoup

Weather = namedtuple('Weather', 'date status high_temp low_temp wind')

strip_white = lambda x: x.replace(' ', '').replace('\n', '').replace('\r', '')


class MonthCode(object):
    fmt = '%Y%m'

    def __init__(self, code):
        try:
            self.code = code
            self.date = datetime.datetime.strptime(code, self.fmt)
        except:
            raise Exception('Invlid month format. Example: 201808 {f}'.format(f=self.fmt))

    def next_month(self):
        self.date = datetime.datetime(year=self.date.year, month=self.date.month + 1, day=1) if self.date.month < 12 else datetime.datetime(year=self.date.year + 1, month=1, day=1)
        self.code = self.date.strftime(self.fmt)

    def __lt__(self, other):
        return self.date < other.date

    def __gt__(self, other):
        return self.date > other.date

    def __le__(self, other):
        return self.date <= other.date

    def __ge__(self, other):
        return self.date >= other.date

    def __eq__(self, other):
        return self.date == other.date


class WeatherHistory(object):
    def __init__(self):
        self.server = 'http://www.tianqihoubao.com'
        self.session = requests.Session()
        self.session.headers.update({
            'Host': 'www.tianqihoubao.com',
            'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10.12; rv:66.0) Gecko/20100101 Firefox/66.0',
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
            'Accept-Language': 'zh-CN,zh;q=0.8,zh-TW;q=0.7,zh-HK;q=0.5,en-US;q=0.3,en;q=0.2',
            'Accept-Encoding': 'gzip, deflate',
            'Referer': 'http://www.tianqihoubao.com/lishi/hainan.html',
            'Connection': 'keep-alive',
            'Upgrade-Insecure-Requests': '1',
            'Pragma': 'no-cache',
            'Cache-Control': 'no-cache',
        })

    def is_city_avaiable(self, city):
        url = '{server}/lishi/{city}.html'.format(server=self.server, city=city)
        try:
            resp = self.session.get(url)
            return resp.status_code == requests.codes.ok
        except:
            return False

    def get_page_by_month(self, city, month):
        url = '{server}/lishi/{city}/month/{month}.html'.format(server=self.server, city=city, month=month)
        print('准备抓取 {m} 的数据... 来源: {url}'.format(m=month, url=url))
        resp = self.session.get(url, timeout=10, allow_redirects=False)
        if resp.status_code == requests.codes.ok:
            print('{m} 数据抓取成功. 来源: {url}'.format(m=month, url=url))
            return resp.content
        else:
            print('页面抓取失败: {url}'.format(url=url))
            raise Exception('{code} {msg}'.format(code=resp.status_code, msg=resp.content))

    def extract_data(self, html):
        data = []
        soup = BeautifulSoup(html, 'lxml')
        tr_list = soup.table.find_all('tr')

        # skip first line for ignore table head
        for tr in tr_list[1:]:
            td_list = tr.find_all('td')
            high_temp, low_temp = map(lambda x: int(x.strip('℃')), strip_white(td_list[2].string).split('/'))
            weather = Weather(strip_white(td_list[0].a.string), strip_white(td_list[1].string), high_temp, low_temp, strip_white(td_list[3].string))
            data.append(weather)

        return data


def data_to_excel(data, output):
    with xlsxwriter.Workbook(output) as workbook:
        worksheet = workbook.add_worksheet()
        for index, label in enumerate(('日期', '天气', '高温', '低温', '风向')):
            worksheet.write(0, index, label)
        for row, line in enumerate(data):
            for col, t in enumerate(line):
                worksheet.write(row + 1, col, t)

        line_chart = workbook.add_chart({'type': 'line'})
        date_region = ['Sheet1', 1, 0, len(data) + 1, 0]
        high_t_region = ['Sheet1', 1, 2, len(data) + 1, 2]
        low_t_region = ['Sheet1', 1, 3, len(data) + 1, 3]
        line_chart.add_series({
            'name': '=Sheet1!$C$1',
            'line': {'color': '#FF0000'},
            'categories': date_region,
            'values': high_t_region,
        })
        line_chart.add_series({
            'name': '=Sheet1!$D$1',
            'line': {'color': '#0000FF'},
            'categories': date_region,
            'values': low_t_region,
        })
        line_chart.set_title({'name': '最高温与最低温折线图'})
        line_chart.set_x_axis({'name': '日期'})
        line_chart.set_y_axis({'name': '气温（单位：℃）'})
        line_chart.set_size({'x_scale': 2, 'y_scale': 2})
        worksheet.insert_chart('J10', line_chart)


def main():
    args = sys.argv
    if len(args) not in (3, 4):
        print('Paramenters Error. Example: python {n} beijing 201808 201809'.format(n=args[0].split(os.path.sep)[-1]))
        return
    if len(args) == 3:
        _, city, begin = args
        end = begin
    else:
        _, city, begin, end = args

    wh = WeatherHistory()

    # 校验参数合法性
    begin_month = MonthCode(begin)
    end_month = MonthCode(end)
    if begin_month > end_month:
        print('结束日期不能早于开始日期')
        return
    if begin_month < MonthCode('201101') or end_month > MonthCode(datetime.datetime.now().strftime(MonthCode.fmt)):
        print('开始日期不能早于2011年')
        return

    if (wh.is_city_avaiable(city)):
        print('准备查找城市 {city} 的天气数据...'.format(city=city))
    else:
        print('没有找到关于指定城市 {city} 的数据.'.format(city=city))
        return

    t_month = MonthCode(begin_month.code)
    data = []

    # 按月份抓取页面，提取数据
    while t_month <= end_month:
        html = wh.get_page_by_month(city, t_month.code)
        data += wh.extract_data(html)

        t_month.next_month()
        print('Sleep 1 seconds...')
        time.sleep(1)

    # 保存至Excel并绘制折线图
    filename = '{n}.xlsx'.format(n='-'.join((city, begin, end)))
    data_to_excel(data, output=filename)
    print('数据已生成Excel文件：{f}'.format(f=filename))


if __name__ == '__main__':
    main()

