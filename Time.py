import time
from CompanyToCode import Tag1


def set_time():
    while True:
        now_hour = time.strftime("%H", time.localtime())
        now_min = time.strftime("%M", time.localtime())
        if now_hour < "05":
            rest = 5 - int(now_hour)
            sleeptime = (rest - 1) * 3600 + (60 - int(now_min)) * 60
            print("启动时北京时间为：" + time.strftime("%H:%M", time.localtime()), "\t软件将在", rest - 1, "小时",
                  int((sleeptime - (rest - 1) * 3600) / 60), "分钟后发送数据")
            time.sleep(sleeptime)
        elif now_hour > "05":
            rest = 5 - int(now_hour) + 24
            sleeptime = (rest - 1) * 3600 + (60 - int(now_min)) * 60
            print("启动时北京时间为：" + time.strftime("%H:%M", time.localtime()), "\t软件将在", rest - 1, "小时",
                  int((sleeptime - (rest - 1) * 3600) / 60), "分钟后发送数据")
            time.sleep(sleeptime)
        elif now_hour == "05":
            print("启动时北京时间为：" + time.strftime("%H:%M", time.localtime()), "\t软件将在每天5点运行！")
            Tag1()
            break
            # time.sleep(86400-int(now_min)*60)


if __name__ == '__main__':
    # 修改参数还是去CompanyToCode，就在那个代码前面
    # 1. 定时跑
    # set_time()
    Tag1()
    # 2.单独跑，这里的输入是excel表，路径是：save\(交易日期文件夹)\跑失败的机构.xlsx，具体哪个交易日的，改date参数就行了
    # 输出是”其他.xlsx“表，在路径：save\(交易日期文件夹)\其他.xlsx
    # Tag2()
