# *-* coding: utf-8 *-*
# 抓取证券日报上的每日交易公告集锦
# http://www.ccstock.cn/meiribidu/jiaoyitishi/
# python 2.7
import requests
import random
import os
import sys
import time
from bs4 import BeautifulSoup
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.header import Header
import smtplib
import codecs
from datetime import date, timedelta
import re
import ConfigParser
import logging

# 日志记录器
logger = logging.getLogger()
# 默认配置信息
DEBUG = True
WEBSITE = "http://www.ccstock.cn/meiribidu/jiaoyitishi"
INTERVAL = 5
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


# 获取当期集锦的url
def parser_list_page(html_doc, now):
    soup = BeautifulSoup(html_doc, 'lxml', from_encoding='utf-8')
    # 只找第一个标签
    tag = soup.find("div", class_="listMain").find("li")
    link_tag = tag.find("a")
    span_tag = tag.find("span")
    page_url = link_tag['href']
    # 截取日期
    date_string = span_tag.string[0:10]
    if date_string == now:
        return page_url
    else:
        return None


def parser_item_page(html_doc, now):
    soup = BeautifulSoup(html_doc, 'lxml', from_encoding='utf-8')
    # title = soup.find("h1").string
    newscontent = soup.find("div", id="newscontent")
    html = newscontent.prettify()
    return html


def get_now():
    if DEBUG:
        yesterday = date.today() - timedelta(1)
        return yesterday.strftime('%Y-%m-%d')
    else:
        now = time.strftime("%Y-%m-%d", time.localtime())
        return now


def write_file(content, now):
    fileName = "morning-" + now + '.txt'
    # full_path = os.path.join(path, fileName)
    f = codecs.open(fileName, 'a', 'utf-8')
    f.write(content)
    f.close()


def read_file(now):
    fileName = "morning-" + now + '.txt'
    # full_path = os.path.join(path, fileName)
    f = open(fileName)
    html = ""
    i = 1
    fileName2 = "morning-" + now + '-f.txt'
    f2 = codecs.open(fileName2, 'a', 'utf-8')
    for text in f.readlines():
        line = text.decode('utf-8')
        newline = transform_number(line, i)
        f2.write(newline)
        html = html + newline
        i = i + 1
    f.close()
    f2.close()
    return html

# 格式化人民币计数
def transform_yuan(lineText):
    # 中标金额313，000，000.00元
    p = re.compile(u"[\d*，]*\d*[.]?\d*万元")
    searchObj = re.findall(p, lineText)
    if searchObj:
        for x in xrange(0, len(searchObj)):
            s1 = searchObj[x]
            ns = filter(lambda ch: ch in '0123456789.', s1)
            nb = float(ns)
            if nb >= 10000:
                s2 = str(nb / 10000) + "亿元"
                lineText = lineText.replace(s1, s1 + "(" + s2 + ")")
    p = re.compile(u"[\d*，]*\d*[.]?\d+元")
    searchObj = re.findall(p, lineText)
    if searchObj:
        for x in xrange(0, len(searchObj)):
            s1 = searchObj[x]
            ns = filter(lambda ch: ch in '0123456789.', s1)
            nb = float(ns)
            if nb >= 100000000:
                s2 = str(nb / 100000000) + "亿元"
                lineText = lineText.replace(s1, s1 + "(" + s2 + ")")
            elif nb >= 10000:
                s2 = str(nb / 10000) + "万元"
                lineText = lineText.replace(s1, s1 + "(" + s2 + ")")
    return lineText

# 格式化股份计数
def transform_gu(lineText):
    p = re.compile(u"[\d*，]*\d+股")
    searchObj = re.findall(p, lineText)
    if searchObj:
        for x in xrange(0, len(searchObj)):
            s1 = searchObj[x]
            ns = filter(lambda ch: ch in '0123456789', s1)
            nb = float(ns)
            if nb >= 100000000:
                s2 = str(nb / 100000000) + "亿股"
                lineText = lineText.replace(s1, s1 + "(" + s2 + ")")
            elif nb >= 10000:
                s2 = str(nb / 10000) + "万股"
                lineText = lineText.replace(s1, s1 + "(" + s2 + ")")
    return lineText


def transform_number(lineText, line):
    lineText = transform_yuan(lineText)
    lineText = transform_gu(lineText)
    # TODO 增减持
    # 行首空格
    return lineText


def send_notice_mail(html, now):
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
    subject = now + ' 日 - 二级市场重要公告集锦'
    msg['Subject'] = Header(subject, 'utf-8')
    msg['From'] = me
    msg['To'] = ";".join(to_list)
    mail_msg = html
    # 邮件正文内容
    msg.attach(MIMEText(mail_msg, 'html', 'utf-8'))

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
        logger.debug('Error: 无法发送邮件 %s ', repr(e))

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


def main(num_retries=3):
    now = get_now()
    logger.debug("now %s", now)
    list_html_doc = download_get_html(WEBSITE)
    page_url = parser_list_page(list_html_doc, now)
    logger.debug("page URL ： %s", page_url)
    if page_url == None:
        # 今日公告还未生成，暂停10分钟后再去尝试
        if num_retries > 0:
            time.sleep(60 * INTERVAL)
            main(num_retries - 1)
        else:
            logger.debug('3次尝试后取消运行')
    else:
        page_html_doc = download_get_html(page_url)
        content = parser_item_page(page_html_doc, now)
        # 本地文件存一份,备查
        write_file(content, now)
        html = read_file(now)
        # 发送邮件通知
        send_notice_mail(html, now)


if __name__ == "__main__":
    reload(sys)
    sys.setdefaultencoding("utf-8")
    init_config()
    init_log()
    logger.debug('start notice montage')
    main()
    logger.debug("notice montage run end")
