'''
目的：获取东方财富网的机构列表，并爬出这些机构分别持有哪些股票
目标网址：http://data.eastmoney.com/hsgtcg/InstitutionStatistics.aspx

用的方法：python3 + selenium webdriver3
1.要安装的库：selenium,openpyxl,pandas
命令行输入：pip install selenium
        pip install openpyxl
        pip install pandas
2.上网下载chromedriver.exe，chrome driver 安装版本要跟chrome一样（在浏览器的设置里面找）,放在Python安装目录下即可，我在D盘放了，就是下面参数的path
    网址：https://npm.taobao.org/mirrors/chromedriver
'''

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from helpFunction import *
import math
import re
from time import sleep, ctime
import threading
import pandas as pd


#需要改的参数
path = 'D:\chromedriver.exe'  # 要改，chromedriver.exe在哪里
date = "2021-05-10"  # 确保是交易日
date_range = "上一个交易日"  # 根据你要的日期改：“上一个交易日”（就是今天爬昨天的数）、“近3日”、“近5日”、“近10日”、“近30日”
# 页数根据机构列表有多少页手动改，特别是节假日前后，数据只有1页，线程只开1；跑其他日期的时候，选择了“近3日”，“近5日”要手动看网站的页数
page_number = 4
N = 5  # 假设分N个线程（实际是N+1个，线程0是是持股数量少于50的机构，因为不需要翻页，可以用requests方法跑）

# 不需要改的参数

dir_path = "save/" + date + "/"
Company_list_url = "http://data.eastmoney.com/hsgtcg/InstitutionStatistics.html"  # 不用改！！
date_range_loc = '//li[text()="%s"]' % date_range
date_range_check_loc = By.XPATH, '//li[@class="linklab spe-padding at" and text()="%s"]' % date_range
init_dirs(dir_path)  # 创建初始文件夹


# 文件名参数
excel_name_for_companys = dir_path + date+"机构总数表.xlsx"  # 机构总数表

def detail_url(code):  # 详情页的url
    return "http://data.eastmoney.com/hsgtcg/StockHdDetail/%s/%s.html" % (code, date)

# 输入输出的表格命名
def excel_name_for_get_codeList(n):
    return dir_path + "线程%d机构名单.xlsx" % n
def excel_name_for_save_shareDetail(n):
    if type(n) == str:
        return dir_path + "/success/%s.xlsx" % n
    return dir_path + "/success/线程%d机构持股明细.xlsx" % n

# 以下是定位参数
def nextLoc(page):
    return '//div[@class="pagerbox"]/a[text()="%s"]' % str(page)
def now_page_loc(page):
    return By.XPATH, '//div[@class="pagerbox"]/a[@class="active" and text()="%s"]' % str(page)


# 启动chrome
def setupDriver():
    chrome_options = Options()
    prefs = {
        'profile.default_content_setting_values': {
            'images': 2,
            'javascript': 2
        }
    }
    chrome_options.add_experimental_option('prefs', prefs)

    chrome_options.add_argument('--headless')  # 静默模式
    chrome_options.add_argument('--disable-gpu')
    driver = webdriver.Chrome(options=chrome_options, executable_path=path)
    # driver.set_page_load_timeout(10)   #没了这2句也能跑，这个是页面加载设置超时的时间
    # driver.set_script_timeout(10)
    return driver


# 获得机构代号和持股数量
def getCompanyAndAmount():
    result = []
    driver = setupDriver()
    BP = BasePage(driver, dir_path, "获取机构名单")
    BP.log("获取机构log.txt\n", "")
    # 翻页
    for page in range(1, page_number + 1):
        BP.log("获取机构第%d页" % page, "")
        print("获取机构第%d页" % page)
        if page == 1:
            if page_number == 1:
                T = BP.getUrl(Company_list_url, "")
            else:
                T = BP.getUrl(Company_list_url, now_page_loc(page))
            if T and date_range != "上一个交易日":
                T = BP.click(date_range_loc, date_range_check_loc)
                if T:
                    BP.log("点击%s button成功！" % date_range, "i")
                else:
                    BP.log("点击%s button失败！" % date_range, "e")
        else:
            T = BP.click(nextLoc(page), now_page_loc(page))

        print("T=", T)
        if T:  # 确认网页是在正确的位置，再获取信息
            # 下面是真正要的东西，获取代码和名称
            try:
                date_loc_list = BP.driver_find_elements("//tbody/tr/td[1]")
                name_loc_list = BP.driver_find_elements("//tbody/tr/td[2]/a")
                code_loc_list = BP.driver_find_elements("//tbody/tr/td[2]/a")
                total_loc_list = BP.driver_find_elements("//tbody/tr/td[4]")  # 当日持有股票总个数
                for i in range(len(date_loc_list)):
                    if date_loc_list[i].text == date:
                        result.append([date_loc_list[i].text, name_loc_list[i].text,
                                       code_loc_list[i].get_attribute("href").split("/")[-2], total_loc_list[i].text])
                sleep(5)
            except Exception as e:
                BP.log("机构第%d页，获取内容失败：%s" % (page, repr(e)), "e")
                BP.quit_browser()
                return
        else:
            BP.log("重要：需要重新运行", "e")
            BP.quit_browser()
            return
    if len(result) < 158:
        print("北向机构数量异常，建议检查一下")
    else:
        print("北向机构共%d个" % len(result))
    r = pd.DataFrame(result,
                     columns=['日期', '机构名称', '机构编号', '持股数量'])
    r.to_excel(excel_name_for_companys, index=False)
    BP.quit_browser()
    return True


def main1(code, count, l):  # 分到这里的所有机构的持股数量都不超过50，所以不用翻页
    l.log("now is %s,共%s个========================" % (code, count), "")
    data = []
    response = getResponse(date, code, l)
    if response:
        # DATE = re.findall('"HDDATE":"(.*?)",', response.text)  # 日期
        participantName = re.findall('"PARTICIPANTNAME":"(.*?)",', response.text)  # 机构名称
        SCODE = re.findall('"SCODE":"(.*?)",', response.text)  # 股票编号
        SNAME = re.findall('"SNAME":"(.*?)",', response.text)  # 股票名称
        CLOSEPRICE = re.findall('"CLOSEPRICE":(.*?),', response.text)  # 当日收盘价
        SHAREHOLDSUM = re.findall('"SHAREHOLDSUM":(.*?),', response.text)  # 持股数量
        SHAREHOLDPRICE = re.findall('"SHAREHOLDPRICE":(.*?),', response.text)  # 持股市值
        ZDF = re.findall('"ZDF":(.*?),', response.text)  # 当日涨跌幅(%)
        ONE = re.findall('"SHAREHOLDPRICEONE":(.*?),', response.text)  # 一日市值变化
        # FIVE = re.findall('"SHAREHOLDPRICEFIVE":(.*?),', response.text)  # 5日市值变化
        if len(SCODE) == count:
            for j in range(count):
                data.append([date, code, participantName[j], str(SCODE[j]), SNAME[j], CLOSEPRICE[j], SHAREHOLDSUM[j],
                             SHAREHOLDPRICE[j], ZDF[j], ONE[j]])
                # 【日期，机构编号，机构名称，股票编号，股票名称，当天收盘价，持股数量，持股市值，当日涨跌幅，一日市值变化】
            l.log("结果=%d个" % count, "i")
            return data
        else:
            l.log("数量不对，持股数量应为%d,实际获取%d" % (count, len(SCODE)), "e")
            return []
    else:
        l.log("%s 的网址连接不上" % code, "e")
        return []


def main2(code, count, l):
    pages = math.ceil(count / 50)
    data = []
    url = detail_url(code)

    for page in range(1, pages + 1):
        l.log("获取第%d页" % page, "i")
        # print("获取第%d页" % page, "i")
        if page == 1:
            T = l.getUrl(url, now_page_loc(page))#
        else:
            T = l.click(nextLoc(page), now_page_loc(page))
            if T:
                l.log("点击成功！", "i")
            else:
                l.log("点击失败！", "e")
        if T:  # 确认网页是在正确的位置，再获取信息
            # 下面是真正要的东西，获取代码和名称
            # if page >10:# 翻页过多的话就等久一点
            #     time.sleep(10)
            try:
                participantName = l.driver_find_element('//span[@class="jgname"]').text  # 机构名称
                table = l.driver_find_elements('//div[@class="dataview-body"]/table/tbody/tr')
            except NameError as e:
                l.log("机构第%d页，获取participantName、table,失败：%s" % (page, repr(e)), "e")
                print("这里:\n%s" % e)
                break
            for tr in table:
                try:
                    SCODE = tr.find_element(By.XPATH, 'td[2]/a').text  # 股票编号
                    SNAME = tr.find_element(By.XPATH, 'td[3]/a').text  # 股票名称
                except Exception as e:
                    l.log("机构第%d页，获取SCODE、SNAME失败：%s" % (page, repr(e)), "e")
                    break
                try:
                    CLOSEPRICE = tr.find_element(By.XPATH, 'td[4]/span').text  # 当日收盘价
                except Exception as e:
                    l.log("code=%s，page=%d，CLOSEPRICE 获取异常" % (code, page), "e")
                    CLOSEPRICE = tr.find_element(By.XPATH, 'td[4]').text
                    l.log("code=%s，page=%d，CLOSEPRICE=%s 获取成功" % (code, page, CLOSEPRICE), "i")
                SHAREHOLDSUM = tr.find_element(By.XPATH, 'td[7]').text  # 持股数量
                SHAREHOLDPRICE = tr.find_element(By.XPATH, 'td[8]').text  # 持股市值
                ZDF = tr.find_element(By.XPATH, 'td[9]').text  # 当日涨跌幅(%)
                ONE = tr.find_element(By.XPATH, 'td[10]/span').text  # 一日市值变化
                data.append([date, code, participantName, str(SCODE), SNAME, CLOSEPRICE,
                             SHAREHOLDSUM, SHAREHOLDPRICE, ZDF, ONE])
                # print([date, code, participantName, str(SCODE), SNAME, CLOSEPRICE,
                #              SHAREHOLDSUM, SHAREHOLDPRICE, ZDF, ONE])
                # 【日期，机构编号，机构名称，股票编号，股票名称，当天收盘价，持股数量，持股市值，当日涨跌幅，一日市值变化】

        else:
            break
    if len(data) == count:
        return data
    else:
        l.log("%s 个数不匹配，失败！" % code, "e")
        return []

def getDetail_1(*args):  # requests方法
    result = pd.DataFrame()
    w = []
    sum_of_success = 0
    open_name = excel_name_for_get_codeList(0)
    save_excel_path = excel_name_for_save_shareDetail(0)
    l = BasePage("", dir_path, 0)
    l.log(open_name + "\n", "")
    print(open_name + " 开始！")
    codeList, countList = get_excel(open_name)
    amount_of_this_thread = len(codeList)
    for i in range(amount_of_this_thread):
        data = main1(codeList[i], countList[i], l)
        if data:
            result = result.append(data)
            l.log("%s成功" % codeList[i], "i")
            sum_of_success += 1
        else:

            w.append([codeList[i], countList[i]])
    l.log("线程0结束，机构共有%d个，跑成功%d个，失败%d个，失败的机构列表：%s\n持股数量应该有%d，实际有%d" % (
        amount_of_this_thread, sum_of_success, len(w), w, sum([int(i) for i in countList]), result.shape[0]), "")
    print("线程0结束，机构共有%d个，跑成功%d个，失败%d个，失败的机构列表：%s" % (amount_of_this_thread, sum_of_success, len(w), w))
    result.columns = ['日期', '机构编号', '机构名称', '股票编号', '股票名称', '当天收盘价', '持股数量', '持股市值', '当日涨跌幅', '一日市值变化']
    result.to_excel(save_excel_path, index=False)


def getDetail_2(arg):  # selenium 方法
    if type(arg) == int:  # arg= n,  线程n
        open_name = excel_name_for_get_codeList(arg)
        n = arg
    else:  # arg=xx.xlsx，
        open_name = arg
        n = "其他"  # 如果是跑自定义的excel表，n=log记录名

    codeList, countList = get_excel(open_name)
    result = pd.DataFrame()
    w = []
    sum_of_success = 0
    save_excel_path = excel_name_for_save_shareDetail(n)
    l = BasePage(setupDriver(), dir_path, n)
    l.log(open_name + "开始！\n", "")
    print(open_name + " 开始！")
    amount_of_this_thread = len(codeList)
    for i in range(amount_of_this_thread):
        code, count = codeList[i], countList[i]
        l.log("现在是第%d个，一共%d个，编号为：%s,持股数量%s个========================" % (i + 1, amount_of_this_thread, code, count), "")
        data = main2(code, count, l)
        if data:
            result = result.append(data)
            l.log("%s成功" % code, "i")
            sum_of_success += 1
        else:
            w.append([code, count])
    l.quit_browser()
    l.log("线程%s结束，机构共有%d个，跑成功%d个，失败%d个，失败的机构列表：%s\n持股数量应该有%d，实际有%d" % (
        str(n), amount_of_this_thread, sum_of_success, len(w), w, sum([int(i) for i in countList]),
        result.shape[0]), "")
    print("线程%s结束，机构共有%d个，跑成功%d个，失败%d个，失败的机构列表：%s" % (str(n), amount_of_this_thread, sum_of_success, len(w), w))
    result.columns = ['日期', '机构编号', '机构名称', '股票编号', '股票名称', '当天收盘价', '持股数量', '持股市值', '当日涨跌幅', '一日市值变化']
    result.to_excel(save_excel_path, index=False)


# if __name__ == '__main__':
def Tag1():
    startTime = ctime()
    '''1. 获得机构代号和持股数量，得机构总表。如果已经存在某日期的机构总数表，则这一步不会运行，所以，如果全部重来就要全部删掉 '''
    if not os.path.exists(excel_name_for_companys):
        getCompanyAndAmount()  # getCompanyAndAmount()获得机构代号和持股数量，输入：，输出：将所有机构保存在一个总excel

    T=divide_excel(dir_path,excel_name_for_companys, N)  # 分n+1个名单, 前提：已有机构总表，输出：n个线程跑的分表
    # 说明： 线程0，是持股数量少于50的机构，即不需要翻页，用requests方法跑
    #       线程1-n，用selenium方法跑
    if T:
        '''2. 多线程跑持股明细，得多个持股明细表和多个log'''
        threads = [threading.Thread(target=getDetail_1, args=(0,))]
        for t in range(1, N + 1):
            threads.append(threading.Thread(target=getDetail_2, args=(t,)))
        # 启动线程
        for t in range(N + 1):
            threads[t].start()
            sleep(2)
        # 守护线程
        for t in range(N + 1):
            threads[t].join()
        sleep(2)
    else:
        getDetail_1(0)
    print(' 主线程开始和结束:%s，%s' % (startTime, ctime()))

    '''3. 合并跑成功的excel表(success文件夹下)，合并log记录'''
    merge_excel(dir_path, date)
    merge_log(dir_path)

    # 保存失败的机构编码
    # if wrong_list:
    #     df = pd.DataFrame(wrong_list, columns=['机构编号', '持股数量'])
    #     df.to_excel(dir_path +"跑失败的机构.xlsx", index=False)

def Tag2():
    excel_path = dir_path + "跑失败的机构.xlsx"
    getDetail_2(excel_path)  #
    l = BasePage(setupDriver(), dir_path, 1)

    print(" 开始！")
    data = main2("C00019", 1753, l)
    l.quit_browser()
    result = pd.DataFrame()
    result = result.append(data)
    result.columns = ['日期', '机构编号', '机构名称', '股票编号', '股票名称', '当天收盘价', '持股数量', '持股市值', '当日涨跌幅', '一日市值变化']
    result.to_excel("1.xlsx", index=False)

