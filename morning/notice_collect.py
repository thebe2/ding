# *-* coding: utf-8 *-*
# 抓取东方财富上的上市公司公告
# http://data.eastmoney.com/notices/
# 代码版本 python 2.7 IDE：PyCharm

import requests
from random import random
import json
import xlrd
import xlwt
import time
import math
import urllib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.header import Header
import smtplib
import base64
import codecs
import ConfigParser
import logging

# 日志记录器
logger = logging.getLogger()
baseurl = "http://data.eastmoney.com"
apiurl = "http://data.eastmoney.com/notices/getdata.ashx?StockCode=&FirstNodeType=0&CodeType=%s&PageIndex=%s&PageSize=1000&jsObj=%s&SecNodeType=0&Time=&rt=%s"
noticeCate = 1
name = "eastmoneynotice"
path = "D:\\crawl\\eastmoney"

plateDic = [
    {
        "code": "hsa",
        "name": "沪深A股",
        "codeType": "1",
    },
    {
        "code": "zxb",
        "name": "中小板",
        "codeType": "4",
    },
    {
        "code": "cyb",
        "name": "创业板",
        "codeType": "5",
    }
]
# 模拟请求头
user_agent_list = [
    "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.1 (KHTML, like Gecko) Chrome/22.0.1207.1 Safari/537.1",
    "Mozilla/5.0 (X11; CrOS i686 2268.111.0) AppleWebKit/536.11 (KHTML, like Gecko) Chrome/20.0.1132.57 Safari/536.11",
    "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/536.6 (KHTML, like Gecko) Chrome/20.0.1092.0 Safari/536.6",
    "Mozilla/5.0 (Windows NT 6.2) AppleWebKit/536.6 (KHTML, like Gecko) Chrome/20.0.1090.0 Safari/536.6",
    "Mozilla/5.0 (Windows NT 6.2; WOW64) AppleWebKit/537.1 (KHTML, like Gecko) Chrome/19.77.34.5 Safari/537.1",
    "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/536.5 (KHTML, like Gecko) Chrome/19.0.1084.9 Safari/536.5",
    "Mozilla/5.0 (Windows NT 6.0) AppleWebKit/536.5 (KHTML, like Gecko) Chrome/19.0.1084.36 Safari/536.5",
    "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/536.3 (KHTML, like Gecko) Chrome/19.0.1063.0 Safari/536.3",
    "Mozilla/5.0 (Windows NT 5.1) AppleWebKit/536.3 (KHTML, like Gecko) Chrome/19.0.1063.0 Safari/536.3",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_8_0) AppleWebKit/536.3 (KHTML, like Gecko) Chrome/19.0.1063.0 Safari/536.3",
    "Mozilla/5.0 (Windows NT 6.2) AppleWebKit/536.3 (KHTML, like Gecko) Chrome/19.0.1062.0 Safari/536.3",
    "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/536.3 (KHTML, like Gecko) Chrome/19.0.1062.0 Safari/536.3",
    "Mozilla/5.0 (Windows NT 6.2) AppleWebKit/536.3 (KHTML, like Gecko) Chrome/19.0.1061.1 Safari/536.3",
    "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/536.3 (KHTML, like Gecko) Chrome/19.0.1061.1 Safari/536.3",
    "Mozilla/5.0 (Windows NT 6.1) AppleWebKit/536.3 (KHTML, like Gecko) Chrome/19.0.1061.1 Safari/536.3",
    "Mozilla/5.0 (Windows NT 6.2) AppleWebKit/536.3 (KHTML, like Gecko) Chrome/19.0.1061.0 Safari/536.3",
    "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/535.24 (KHTML, like Gecko) Chrome/19.0.1055.1 Safari/535.24",
    "Mozilla/5.0 (Windows NT 6.2; WOW64) AppleWebKit/535.24 (KHTML, like Gecko) Chrome/19.0.1055.1 Safari/535.24"
]

def download_get_html(url, charset="utf-8", timeout=10, num_retries=3):
    UA = random.choice(user_agent_list)
    headers = {
        'User-Agent': UA,
        'Content-Type': 'text/html; charset=' + charset
    }
    try:
        response = requests.get(url, headers=headers,
                                timeout=timeout)
        # 设置编码
        response.encoding = charset
        # 404 容错
        if response.status_code == 404:
            logger.debug('get 404: %s ', url)
            return None
        else:
            logger.debug('get : %s ', url)
            return response.text
    except:
        if num_retries > 0:
            time.sleep(10)
            logger.debug('正在尝试，10S后将重新获取倒数第 %d  次', num_retries)
            return download_get_html(url, charset, timeout, num_retries - 1)
        else:
            logger.debug('尝试也不好使了！取消访问')
            return None

# 打开excel
def open_excel(file='file.xls'):
    try:
        data = xlrd.open_workbook(file)
        return data
    except Exception, e:
        logger.debug(str(e))

# 分析excel 数据,获取分类数量
def analyze_excel(fileName):
    file = open_excel(fileName)
    workbook = xlwt.Workbook(encoding='utf-8')
    wsb = workbook.add_sheet("汇总")
    wsb.write(0, 0, label="板块")
    wsb.write(0, 1, label="公告类型")
    wsb.write(0, 2, label="数量")
    alignment = xlwt.Alignment()  # Create Alignment
    # May be: HORZ_GENERAL, HORZ_LEFT, HORZ_CENTER, HORZ_RIGHT,
    #HORZ_FILLED, HORZ_JUSTIFIED, HORZ_CENTER_ACROSS_SEL, HORZ_DISTRIBUTED
    alignment.horz = xlwt.Alignment.HORZ_CENTER  # 水平居中
    # May be: VERT_TOP, VERT_CENTER, VERT_BOTTOM, VERT_JUSTIFIED,
    # VERT_DISTRIBUTED
    alignment.vert = xlwt.Alignment.VERT_CENTER  # 垂直居中
    style = xlwt.XFStyle()  # Create Style
    style.alignment = alignment  # Add Alignment to Style
    x = 1
    for worksheet in file.sheets():
        nrows = worksheet.nrows  # 行数
        ncols = worksheet.ncols  # 列数
        docs = {}
        for rownum in range(1, nrows):
            row = worksheet.row_values(rownum)
            dataType = row[3]
            if dataType in docs:
                docs[dataType] = docs[dataType] + 1
            else:
                docs[dataType] = 1

        for name in docs:
            print worksheet.name, name, docs[name]
            # TODO sheet.write_merge(0, 0, 0, 1, 'Long Cell')
            # wsb.write(x, 0, label = worksheet.name)
            wsb.write(x, 1, label=name)
            wsb.write(x, 2, label=docs[name])
            x = x + 1
        print x, len(docs)
        if len(docs) == 0:
            wsb.write(x, 0, worksheet.name.encode("utf-8"), style)
        else:
            wsb.write_merge(x - len(docs), x - 1, 0, 0,
                            worksheet.name.encode("utf-8"), style)
        print "-------"
    workbook.save(u"汇总.xls")


# load_table_data.js getCode
def getCode(num=6):
    s = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"
    codes = list(s)
    code = ""
    for x in xrange(0, num):
        idx = int(math.floor(random() * 52))
        code += codes[idx]
    return code

# _url += parseInt(parseInt(new Date().getTime()) / 30000);
def getRightTime():
    r = int(time.time() / 30)
    return r


def getUrl(url, codetype, page, code, rt):
    return url % (codetype, page, code, rt)


def parser_data(data):
    temp = data["CDSY_SECUCODES"][0]
    noteicedate = data["NOTICEDATE"]
    date = noteicedate[0:noteicedate.index('T')]
    code = temp["SECURITYCODE"]
    name = temp["SECURITYSHORTNAME"]
    title = data["NOTICETITLE"]
    typeName = '公司公告'
    if data["ANN_RELCOLUMNS"] and len(data["ANN_RELCOLUMNS"]) > 0:
        typeName = data["ANN_RELCOLUMNS"][0]["COLUMNNAME"]
    namestr = unicode(name).encode("utf-8")
    detailLink = baseurl + '/notices/detail/' + code + '/' + \
        data["INFOCODE"] + ',' + \
        base64.b64encode(urllib.quote(namestr)) + '.html'
    # print date,code,name,title,typeName,detailLink
    if time_compare(date):
        return [code, name, title, typeName, detailLink, date]
    else:
        return None


def time_compare(notice_date):
    tt = time.mktime(time.strptime(notice_date, "%Y-%m-%d"))
    # 得到公告的时间戳
    if noticeCate == 1:
        # A股公告取当日
        # 得到本地时间（当日零时）的时间戳
        st = time.strftime("%Y-%m-%d", time.localtime(time.time()))
    else:
        # 新三板公告取前日
        # 得到本地时间（当日零时）的时间戳
        st = time.strftime(
            "%Y-%m-%d", time.localtime(time.time() - 60 * 60 * 24))
    t = time.strptime(st, "%Y-%m-%d")
    now_ticks = time.mktime(t)
    # 周一需要是大于
    if tt >= now_ticks:
        return True
    else:
        return False


def do_notice(notices, plate):
    for page in xrange(1, 10):
        rt = getRightTime()
        code = getCode(8)
        url = getUrl(apiurl, plate["codeType"], page, code, rt)
        jsdata = download_get_html(url)
        if jsdata != None:
            json_str = jsdata[15:-1]
            datas = json.loads(json_str)["data"]
            for data in datas:
                # 公告日期
                notice = parser_data(data)
                if notice != None:
                    notices.append(notice)
                else:
                    logger.debug("page end notices %s %d"& (plate["name"], len(notices)))
                    return
        else:
            logger.debug("no  notices %s %d"& (plate["name"], len(notices)))
            return

# 写excel
def write_sheet(workbook, sheetName, rows):
    worksheet = workbook.add_sheet(sheetName)
    worksheet.write(0, 0, label="代码")
    worksheet.write(0, 1, label="名称")
    worksheet.write(0, 2, label="公告标题")
    worksheet.write(0, 3, label="公告类型")
    for x in xrange(0, len(rows)):
        row = rows[x]
        for y in xrange(0, 4):
            if y == 2:
                alink = 'HYPERLINK("%s";"%s")' % (row[4], row[2])
                worksheet.write(x + 1, y, xlwt.Formula(alink))
            else:
                item = row[y]
                worksheet.write(x + 1, y, item)


def render_mail(name, rows):
    html_mail = ""
    header_tpl = """
    <h2>%s</h2>
        <table>
            <thead>
                <tr>
                    <th style="width: 60px; padding: 0px;text-align:center">代码</th>
                    <th style="width: 110px; padding: 0px;text-align:center">名称</th>
                    <th style="width: 385px; padding: 0px;text-align:center">公告标题</th>
                    <th style="width: 110px; padding: 0px;text-align:center">公告类型</th>
                    <th style="width: 80px; padding: 0px;text-align:center">公告日期</th>
                </tr>
            </thead>
            <tbody>
    """
    html_mail = html_mail + header_tpl % (name)
    tr_tpl = """
        <tr>
                    <td style="text-align:center">
                        %s
                    </td>
                    <td style="text-align:center">
                        %s
                    </td>
                    <td>
                        <a  style="text-align:left;width:350px" href="%s" title="%s">%s</a>
                    </td>
                    <td style="text-align:center">
                        %s
                    </td>
                    <td style="text-align:center">
                        %s
                    </td>
                </tr>
    """
    for row in rows:
        trs = tr_tpl % (unicode(row[0]).encode("utf-8"), unicode(row[1]).encode("utf-8"), unicode(row[4]).encode("utf-8"), unicode(
            row[2]).encode("utf-8"), unicode(row[2]).encode("utf-8"), unicode(row[3]).encode("utf-8"), unicode(row[5]).encode("utf-8"))
        html_mail = html_mail + trs
    footer = "</tbody></table>"
    html_mail = html_mail + footer
    return html_mail


def write_html(now, html):
    f = codecs.open(name + "-" + now + '.html', 'a', 'utf-8')
    f.write(unicode(html, "utf-8"))
    f.close()


def read_html(now):
    ipath = name + "-" + now + '.html'
    f = open(ipath)
    html = ""
    for text in f.readlines():
        html = html + text.decode('utf-8')
    f.close()
    return html




def send_notice_mail(fileName, now):
    cf = get_config_parser()
    to_list = cf.get("mailconf", "to_list").split(",")
    mail_host = cf.get("mailconf", "mail_host")
    mail_username = cf.get("mailconf", "mail_username")
    mail_user = cf.get("mailconf", "mail_user")
    mail_pass = cf.get("mailconf", "mail_pass")
    mail_postfix = cf.get("mailconf", "mail_postfix")
    me = "AStockMarketNoticeWatcher" + "<" + \
         mail_username + "@" + mail_postfix + ">"
    msg = MIMEMultipart()
    subject = now + ' 日 - 二级市场公告信息每日更新'
    msg['Subject'] = Header(subject, 'utf-8')
    msg['From'] = me
    msg['To'] = ";".join(to_list)

    mail_msg = read_html(now)
    # 邮件正文内容
    # msg.attach(MIMEText(now+" 日,二级市场公告信息。详情请见附件excel", 'plain', 'utf-8'))
    msg.attach(MIMEText(mail_msg, 'html', 'utf-8'))

    # 构造附件2，传送当前目录下的 xls 文件
    att2 = MIMEText(open(fileName, 'rb').read(), 'base64', 'utf-8')
    att2["Content-Type"] = 'application/octet-stream'
    # 解决中文附件下载时文件名乱码问题
    att2.add_header('Content-Disposition', 'attachment', filename='=?utf-8?b?' +
                    base64.b64encode(fileName.encode('UTF-8')) + '?=')
    msg.attach(att2)

    try:
        server = smtplib.SMTP()
        server.connect(mail_host)
        server.ehlo()
        server.starttls()
        server.login(mail_user, mail_pass)
        server.sendmail(me, to_list, msg.as_string())
        server.close()
        logger.debug('sent mail successfully')
    except smtplib.SMTPException, e:
        # 参考http://www.cnblogs.com/klchang/p/4635040.html
        logger.debug('Error: 无法发送邮件', repr(e))

def get_config_parser():
    config_file_path = "notice_montage.ini"
    cf = ConfigParser.ConfigParser()
    cf.read(config_file_path)
    return cf

# 解析配置
def init_config():
    cf = get_config_parser()
    global DEBUG, INTERVAL, WEBSITE
    INTERVAL = int(cf.get("timeconf", "interval"))
    DEBUG = cf.get("urlconf", "debug") == 'True'
    WEBSITE = cf.get("urlconf", "website")


def init_log():
    if DEBUG:
        # 测试日志输出到流
        handler = logging.StreamHandler()
    else:
        # 正式日志输出到文件，备查
        handler = logging.FileHandler("notice_montage.log")
    formatter = logging.Formatter(
        '%(asctime)s %(name)-12s %(levelname)-8s %(message)s')
    handler.setFormatter(formatter)
    logger.addHandler(handler)
    logger.setLevel(logging.DEBUG)



def main(fileName, now):
    workbook = xlwt.Workbook(encoding='utf-8')
    for plate in plateDic:
        if plate["code"] == "sb":
            global noticeCate
            noticeCate = 2
        else:
            global noticeCate
            noticeCate = 1
        notices = []
        do_notice(notices, plate)
        if len(notices) > 0:
            write_sheet(workbook, plate["name"], notices)
            html = render_mail(plate["name"], notices)
            write_html(now, html)
    workbook.save(fileName)
    # send_notice_mail(fileName, now)


def run(fileName, now, num_retries=3):
    try:
        main(fileName, now)
    except Exception, e:
        logger.debug(str(e))
        if num_retries > 0:
            time.sleep(10)
            logger.debug('公告抓取正在尝试，10S后将重新获取倒数第', num_retries, '次')
            run(fileName, now, num_retries - 1)
        else:
            logger.debug('公告抓取尝试也不好使了！取消运行')

if __name__ == "__main__":
    logger.debug("start")
    num_retries = 3
    now = time.strftime("%Y-%m-%d", time.localtime())
    fileName = "gg-" + now + ".xls"
    run(fileName, now)
    # analyze_excel(fileName)
    logger.debug("end")
