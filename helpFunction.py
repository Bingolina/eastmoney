import os
import requests
import os.path
import pandas as pd
import time
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver import ActionChains

def getResponse(date, code, l):  # 连接失败的话，尝试3次
    month, day = date.split("-")[1], date.split("-")[2]
    url = "http://dcfm.eastmoney.com/em_mutisvcexpandinterface/api/js/get?callback" \
          "=jQuery1123004297432063462936_1618504179999&st=HDDATE&sr=-1&ps=50&p=1&token" \
          "=894050c76af8597a853f5b408b759f5d&js=%7B%22data%22%3A(x)%2C%22pages%22%3A(tp)%2C%22font%22%3A(" \
          "font)%7D&filter=(PARTICIPANTCODE%3D%27" + code + "%27)(MARKET+in+(%27001%27%2C%27003%27))(" \
                                                            "HDDATE%3D%5E2021%2F" + str(
        month) + "%2F" + str(day) + "%5E)&type=HSGTNHDDET"


    i = 0
    while i < 3:
        try:
            response = requests.get(url, timeout=(5, 10))
            l.log("页面连接成功！状态码=%s" % response.status_code, "i")
            return response
        except requests.exceptions.RequestException as e:
            i += 1
            l.log("第%d次连接不上:%s" % (i, e), "i")
    l.log("%s连接不上了" % code, "e")
    return ""


def init_dirs(dir_path):
    if not os.path.exists(dir_path + "/success/"):
        os.makedirs(dir_path + "/success/")


def divide_excel(dir_path,excel_name_for_companys, slice_num):
    df = pd.read_excel(excel_name_for_companys)
    df50 = df[df['持股数量'] < 51]  # 持股数量少于50的机构，即不需要翻页，用requests方法跑,所以单独放
    df50.to_excel(dir_path + "线程0机构名单.xlsx", encoding='gbk', index=False)
    df1 = df[df['持股数量'] > 50]
    if df1.shape[0]!=0:
        df1.reset_index(drop=True, inplace=True)
        code_list = df1['机构编号'].tolist()
        count_list = df1['持股数量'].tolist()
        total = len(code_list)
        for i in range(slice_num):
            c1 = []
            c2 = []
            for j in range(i, total, slice_num):
                c1.append(code_list[j])
                c2.append(count_list[j])
            d = pd.DataFrame({'机构编号': c1, '持股数量': c2}, columns=['机构编号', '持股数量'])
            d.to_excel(dir_path + "线程" + str(i+1) + "机构名单.xlsx", encoding='gbk', index=False)
    else:
        print("只有线程0，其他不需要")
        return ""
    return True


def get_excel(path):
    df = pd.read_excel(path)
    codeList = df['机构编号'].tolist()
    countList = df['持股数量'].tolist()
    return codeList, countList


def merge_excel(dir_path, date):
    files = os.listdir(dir_path + "success/")
    print(files)

    try:
        df = pd.read_excel(dir_path + "success/"+files[0])
        if len(files) != 1:
            for i in range(1, len(files)):
                df1 = pd.read_excel(dir_path + "success/"+files[i])
                df = df.append(df1)
        df.to_excel(dir_path + date + "北向机构持股数据汇总.xlsx")
        print("合并表格成功。")
    except NameError as e:
        print("合并表格失败： %s" % e)


def merge_log(dir_path):
    files = os.listdir(dir_path)
    logs_files = []
    for file in files:
        if ".txt" in file:
            logs_files.append(file)
    lines = []
    try:
        for i in logs_files:
            with open(dir_path+i, 'r', encoding='utf-8') as f:
                lines.append(f.readlines())
        with open(dir_path + "log汇总.txt", 'w', encoding='utf-8') as txt:
            for line in lines:
                txt.writelines(line)
                txt.write("\n\n\n")
        print("合并log成功")
        try:
            for log_name in logs_files:
                os.remove(dir_path + log_name)
        except NameError as e:
            print("删除log失败： %s" % e)
    except NameError as e:
        print("合并或删除log失败： %s" % e)


class BasePage(object):
    """
    定义一个页面基类，让所有页面都继承这个类，封装一些常用的页面操作方法到这个类
    """

    def __init__(self, driver, dir_path, n):
        self.driver = driver
        self.dir_path = dir_path
        self.n = n  # 第n个线程，

    # 打开网址
    def getUrl(self, url, check_loc):
        for i in range(3):
            try:
                self.driver.get(url)
                time.sleep(4)
                return True
            except NameError as e:
                self.log("网址打开失败第%d次，url=%s： %s" % (i, url, e), "e")
                print("网址打开失败第%d次，url=%s： %s" % (i, url, e))
        return False

                # if check_loc!="":
                #     try:
                #         WebDriverWait(self.driver, 30, 1).until(EC.presence_of_element_located((check_loc)))
                #     except:
                #         print("网页刷新失败,url=",url)
                # else:
                #     self.log("网址打开成功,url=%s" % url, "i")
                #     return True


    # quit browser 关闭全部浏览器
    def quit_browser(self):
        self.driver.quit()

    # 关闭当前浏览器
    def close_browser(self):
        self.driver.close()

    # 题外话：*arg 可变参数，允许你传入0个或任意个参数，这些可变参数在函数调用时自动组装为一个tuple
    #       **arg 关键字参数允许你传入0个或任意个含参数名的参数，这些关键字参数在函数内部自动组装为一个dict

    # 找到定位的元素
    def driver_find_element(self, loc):
        try:
            el = self.driver.find_element(By.XPATH, loc)
            return el
        except Exception as e:
            self.log("定位不到元素(%s)： %s" % (loc, e), "e")
            return False

    # 找到定位的元素
    def driver_find_elements(self, loc):
        try:
            el = self.driver.find_elements(By.XPATH, loc)
            return el
        except Exception as e:
            self.log("定位不到元素(%s)： %s" % (loc, e), "e")

    # 点击元素，并验证
    def click(self, loc, check_loc):
        next_page = self.driver_find_element(loc)
        if not next_page:
            return False
        # print("现在点击=", loc)
        try:
            next_page.click()
            time.sleep(4)
            return True
        except NameError as e:
            self.log("点击失败： %s" % ( e), "e")
            print("点击失败次： %s" % ( e))
            return False
        # for _ in range(3):
        #     check = self.driver_find_element(check_loc)
        #     if check:
        #         self.log("点击成功且check_loc也验证了", "i")
        #         return True
        #     time.sleep(3)
        # self.log("点击成功但页面更新不成功", "e")
        # return False




    # 记录每个线程的log输出
    def log(self, text, tag):
        if type(self.n) == int:
            log_path = self.dir_path + "/log-线程" + str(self.n) + '.txt'
        else:  # n 如果不是线程数字，那么就是保存的log文件名
            log_path = self.dir_path + "/" + self.n + "-log.txt"
        with open(log_path, 'a', encoding='utf-8') as f:
            if tag == "i":
                line = "        " + text + "\n"
            elif tag == "e":
                line =  "        **Error: " + text + "\n"
            else:
                line = text + "\n"
            f.write(line)

    #

    # 下面的暂时用不到
    # 浏览器前进操作
    def forward(self):
        self.driver.forward()

    # 浏览器后退操作
    def back(self):
        self.driver.back()

    # 输入
    def type(self, location, text):
        el = self.driver_find_element(location)
        # 清除文本框
        el.clear()
        try:
            el.send_keys(text)
            self.log("Had type \' %s \' in inputBox" % text, "i")
        except NameError as e:
            self.log("Failed to type in input box with %s" % e, "e")

    # 网页标题
    def get_page_title(self):
        self.log("Current page title is %s" % self.driver.title, "i")
        return self.driver.title
