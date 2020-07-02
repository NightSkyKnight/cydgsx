# coding=utf-8

'''
闯越顶岗实习签到程序
'''


import requests
import sys
import json
import win32api
import win32con
import os
import _thread
import datetime


cookie_login = {}
cookie_index = {}


# 登录
def get_login():
    global cookie_login
    global cookie_index
    login_headers = {
        'Host': 'hl.cydgsx.com',
        'Connection': 'keep-alive',
        'Accept': 'application/json, text/javascript, */*; q=0.01',
        'X-Requested-With': 'XMLHttpRequest',
        'User-Agent': 'Mozilla/5.0 (Linux; Android 10; Redmi K20 Pro Build/QKQ1.190825.002; wv) AppleWebKit/537.36 (KHTML, like Gecko) Version/4.0 Chrome/80.0.3987.99 Mobile Safari/537.36',
        'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
        'Origin': 'https://hl.cydgsx.com',
        'Sec-Fetch-Site': 'same-origin',
        'Sec-Fetch-Mode': 'cors',
        'Sec-Fetch-Dest': 'empty',
        'Referer': 'https://hl.cydgsx.com/m/Home/Index',
        'Accept-Encoding': 'gzip, deflate, br',
        'Accept-Language': 'zh-CN,zh;q=0.9,en;q=0.8,ja;q=0.7'
    }

    # 账号
    username = '440602199910020931'
    # 密码
    UserPwd = '440602199910020931'

    login_data = 'username={}&UserPwd={}&wxInfo=&openid='.format(
        username, UserPwd)
    login = requests.post(
        'https://hl.cydgsx.com/m/Home/CheckLoginJson', headers=login_headers, data=login_data)
    cookie_login = requests.utils.dict_from_cookiejar(login.cookies)

    index_headers = {
        'Host': 'hl.cydgsx.com',
        'Connection': 'keep-alive',
        'Upgrade-Insecure-Requests': '1',
        'User-Agent': 'Mozilla/5.0 (Linux; Android 10; Redmi K20 Pro Build/QKQ1.190825.002; wv) AppleWebKit/537.36 (KHTML, like Gecko) Version/4.0 Chrome/80.0.3987.99 Mobile Safari/537.36',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
        'Sec-Fetch-Site': 'same-origin',
        'Sec-Fetch-Mode': 'navigate',
        'Sec-Fetch-User': '?1',
        'Sec-Fetch-Dest': 'document',
        'Referer': 'https://hl.cydgsx.com/m/Home/Index',
        'Accept-Encoding': 'gzip, deflate, br',
        'Accept-Language': 'zh-CN,zh;q=0.9,en;q=0.8,ja;q=0.7',
        'Cookie': 'LoginUser_Id={0}; loginUserName={1}'.format(cookie_login['LoginUser_Id'], cookie_login['loginUserName'])
    }
    get_index = requests.get(
        'https://hl.cydgsx.com/m/s/Home/Index', headers=index_headers)
    cookie_index = requests.utils.dict_from_cookiejar(get_index.cookies)
    # 保存cookie
    save_cookie()


# 保存cookie
def save_cookie():
    global cookie_index
    global cookie_login
    file_path = os.path.dirname(__file__)      # 分离文件路径与后缀
    with open('{}\cookie_index.txt'.format(file_path), "w") as index:
        json.dump(cookie_index, index)
    # pprint.pprint(cookie_index, width=5)
    with open("{}\cookie_login.txt".format(file_path), "w") as login:
        json.dump(cookie_login, login)
    # pprint.pprint(cookie_login, width=5)


# 读取cookie
def open_cookie():
    global cookie_index
    global cookie_login
    try:
        with open("cookie_index.txt", "r") as index:
            cookie_index = json.load(index)
        with open("cookie_login.txt", "r") as login:
            cookie_login = json.load(login)
    except:
        get_login()


# 获取当前日期
def get_datetime():
    # 当前日期
    now = datetime.datetime.now().date()
    year, month, day = str(now).split("-")  # 切割
    # 年月日，转换为数字
    year = int(year)
    month = int(month)
    day = int(day)

    # 判断当前是不是这个星期最后一天
    weekday = datetime.datetime.now().weekday()
    if weekday == 6:
        try:
            _thread.start_new_thread(zhou_qiandao, (now,))
        except:
            print('无法启动周签到线程')

    # 判断当前是不是这个月最后一天
    def last_day(any_day):
        # any_day的天数改为28并加4
        next_month = any_day.replace(day=28) + datetime.timedelta(days=4)
        # 计算的日期减去天数等于上个月最后一天
        return next_month - datetime.timedelta(days=next_month.day)
    # 获取这个月最后一天
    last_day = last_day(datetime.date(year, month, day))
    if now == last_day:
        yue_qiandao(now)


# 当天签到
def qiandao():
    # LoginTimeCooikeName 未发现如何生成
    headers = {
        'Host': 'hl.cydgsx.com',
        'Connection': 'keep-alive',
        'Content-Length': '292',
        'Accept': '*/*',
        'Origin': 'https://hl.cydgsx.com',
        'X-Requested-With': 'XMLHttpRequest',
        'Sec-Fetch-Dest': 'empty',
        'User-Agent': 'Mozilla/5.0 (Linux; Android 10; Redmi K20 Pro Build/QKQ1.190825.002; wv) AppleWebKit/537.36 (KHTML, like Gecko) Version/4.0 Chrome/80.0.3987.99 Mobile Safari/537.36',
        'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
        'Sec-Fetch-Site': 'same-origin',
        'Sec-Fetch-Mode': 'cors',
        'Referer': 'https://hl.cydgsx.com/m/s/Log/wLog',
        'Accept-Encoding': 'gzip, deflate',
        'Accept-Language': 'zh-CN,zh;q=0.9,en-US;q=0.8,en;q=0.7',
        'Cookie': 'ASP.NET_SessionId={0}; jxnApp=0; giveCard_10269={1}; loginUserName={2}; LoginTimeCooikeName=8ee100ae0db2c3afd71eaf3501ec36b1; LoginUser_Id={3}'.format(cookie_index['ASP.NET_SessionId'], cookie_index['giveCard_10269'], cookie_login['loginUserName'], cookie_login['LoginUser_Id'])
    }
    # interContent 文本内容
    # posAddress 中文地址
    # posLong posLati 经纬度
    qiang_data = 'InternStateId=4&interContent=%E7%B4%AF&logImg=&posAddress=%E4%B8%AD%E5%9B%BD%E5%B9%BF%E4%B8%9C%E7%9C%81%E4%BD%9B%E5%B1%B1%E5%B8%82%E7%A6%85%E5%9F%8E%E5%8C%BA%E8%BD%BB%E5%B7%A5%E5%8C%97%E4%B8%83%E8%A1%978%E5%8F%B7&posLong=113.09530631503551&posLati=23.03589761307942&locationType=1&ArticleId=0'

    get_qian = requests.post(
        'https://hl.cydgsx.com/m/s/Log/SaveWriteLog', headers=headers, data=qiang_data)
    if get_qian.text.find('请重新登录') != -1:
        get_login()
        qiandao()
    get_json = json.loads(get_qian.text)
    try:
        if get_json['state'] == 1:
            print('已签到')
        else:
            print('重复签到')
            print(get_json['meg'])
    except KeyError as e:
        print(e)
        print('尝试重新运行')


# 补签
def buqian():
    headers = {
        'Host': 'hl.cydgsx.com',
        'Connection': 'keep-alive',
        'Accept': '*/*',
        'X-Requested-With': 'XMLHttpRequest',
        'User-Agent': 'Mozilla/5.0 (Linux; Android 10; Redmi K20 Pro Build/QKQ1.190825.002; wv) AppleWebKit/537.36 (KHTML, like Gecko) Version/4.0 Chrome/80.0.3987.99 Mobile Safari/537.36',
        'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
        'Origin': 'https://hl.cydgsx.com',
        'Sec-Fetch-Site': 'same-origin',
        'Sec-Fetch-Mode': 'cors',
        'Sec-Fetch-Dest': 'empty',
        'Referer': 'https://hl.cydgsx.com/m/s/Log/wReplacementLog?date={}'.format(time_day),
        'Accept-Encoding': 'gzip, deflate, br',
        'Accept-Language': 'zh-CN,zh;q=0.9,en;q=0.8,ja;q=0.7',
        'Cookie': 'LoginUser_Id=10269&logintype=2&RoleOId=295&UserName=440602199910020931&Name=%e9%bb%84%e4%bf%8a%e8%b1%aa&unGuid=28533bf0ec1f429eb9b188481218ac31; loginUserName=440602199910020931; ASP.NET_SessionId=dsmjtf1dngi51weqds2pjtzx; giveCard_10269=%7b%22id%22%3a0%2c%22isnew%22%3a0%2c%22title%22%3a%22%22%7d'
    }

    time_day_data = 'date={}&content=%E7%B4%AF&content_image=&address=%E4%B8%AD%E5%9B%BD%E5%B9%BF%E4%B8%9C%E7%9C%81%E4%BD%9B%E5%B1%B1%E5%B8%82%E7%A6%85%E5%9F%8E%E5%8C%BA%E8%BD%BB%E5%B7%A5%E5%8C%97%E4%B8%83%E8%A1%978%E5%8F%B7&longitude=113.09549430903348&latitude=23.03578462101444'.format(
        time_day)

    get_day = requests.post(
        'https://hl.cydgsx.com/m/s/Log/SavewReplacementLog', headers=headers, data=time_day_data)


# 周记
def zhou_qiandao(today):
    headers = {
        'Host': 'hl.cydgsx.com',
        'Connection': 'keep-alive',
        'Accept': 'application/json, text/javascript, */*; q=0.01',
        'Origin': 'https://hl.cydgsx.com',
        'X-Requested-With': 'XMLHttpRequest',
        'Sec-Fetch-Dest': 'empty',
        'User-Agent': 'Mozilla/5.0 (Linux; Android 10; Redmi K20 Pro Build/QKQ1.190825.002; wv) AppleWebKit/537.36 (KHTML, like Gecko) Version/4.0 Chrome/80.0.3987.99 Mobile Safari/537.36',
        'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
        'Sec-Fetch-Site': 'same-origin',
        'Sec-Fetch-Mode': 'cors',
        'Referer': 'https://hl.cydgsx.com/m/s/Log/wWeekSmy?date={}'.format(today),
        'Accept-Encoding': 'gzip, deflate',
        'Accept-Language': 'zh-CN,zh;q=0.9,en-US;q=0.8,en;q=0.7',
        'Cookie': 'ASP.NET_SessionId={0}; jxnApp=0; giveCard_10269={1}; loginUserName={2}; LoginTimeCooikeName=8ee100ae0db2c3afd71eaf3501ec36b1; LoginUser_Id={3}'.format(cookie_index['ASP.NET_SessionId'], cookie_index['giveCard_10269'], cookie_login['loginUserName'], cookie_login['LoginUser_Id'])
    }
    zhou_data = 'logImg=&smyDate={}&summaryType=%E5%91%A8%E5%B0%8F%E7%BB%93&summaryInfo=%E7%B4%AF'.format(
        today)
    zhou = requests.post(
        'https://hl.cydgsx.com/m/s/Log/SaveSmyJson', headers=headers, data=zhou_data)
    zhou_json = json.loads(zhou.text)
    try:
        if zhou_json['state'] == 1:
            print('已填写周记')
        else:
            print('重复填写周记：{}'.format(zhou_json['meg']))
    except KeyError as e:
        print(e)
        print('尝试重新运行')


# 月记
def yue_qiandao(today):
    headers = {
        'Host': 'hl.cydgsx.com',
        'Connection': 'keep-alive',
        'Accept': 'application/json, text/javascript, */*; q=0.01',
        'Origin': 'https://hl.cydgsx.com',
        'X-Requested-With': 'XMLHttpRequest',
        'Sec-Fetch-Dest': 'empty',
        'User-Agent': 'Mozilla/5.0 (Linux; Android 10; Redmi K20 Pro Build/QKQ1.190825.002; wv) AppleWebKit/537.36 (KHTML, like Gecko) Version/4.0 Chrome/80.0.3987.99 Mobile Safari/537.36',
        'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
        'Sec-Fetch-Site': 'same-origin',
        'Sec-Fetch-Mode': 'cors',
        'Referer': 'https://hl.cydgsx.com/m/s/Log/wMonthSmy?date={}'.format(today),
        'Accept-Encoding': 'gzip, deflate',
        'Accept-Language': 'zh-CN,zh;q=0.9,en-US;q=0.8,en;q=0.7',
        'Cookie': 'ASP.NET_SessionId={0}; jxnApp=0; giveCard_10269={1}; loginUserName={2}; LoginTimeCooikeName=8ee100ae0db2c3afd71eaf3501ec36b1; LoginUser_Id={3}'.format(cookie_index['ASP.NET_SessionId'], cookie_index['giveCard_10269'], cookie_login['loginUserName'], cookie_login['LoginUser_Id'])
    }

    yue_data = 'logImg=&smyDate={}&summaryType=%E6%9C%88%E5%B0%8F%E7%BB%93&summaryInfo=%E8%BF%98%E4%B8%8D%E9%94%99'.format(
        today)
    yue = requests.post(
        'https://hl.cydgsx.com/m/s/Log/SaveSmyJson', headers=headers, data=yue_data)
    yue_json = json.loads(yue.text)
    try:
        if yue_json['state'] == 1:
            print('已填写月记')
        else:
            print('重复填写月记：{}'.format(yue_json['meg']))
    except KeyError as e:
        print(e)
        print('尝试重新运行')


# 获取键值
def get_keyvalue(key):
    try:
        i = 0
        while True:
            # 循环枚举值
            yield win32api.RegEnumValue(key, i)
            i += 1
    except Exception as e:
        pass
    # 无论try语句中是否抛出异常，finally中的语句一定会被执行
    finally:
        key.close()


# 注册表操作
def open_win():
    file_name = os.path.basename(__file__)  # 当前文件名的名称
    file_path = os.path.splitext(file_name)[0]      # 分离文件路径与后缀
    file_exe = str(os.path.dirname('./'))+'\\'+file_path
    path = os.path.abspath(file_exe)  # 获取文件的绝对路径
    print(path)
    # 注册表项名
    KeyName = 'Software\\Microsoft\\Windows\\CurrentVersion\\Run'
    # 设置注册表键
    keyHandle = win32api.RegConnectRegistry(None, win32con.HKEY_CURRENT_USER)

    # 查找注册表是否创建过键值
    # 打开注册表键
    key = win32api.RegOpenKeyEx(keyHandle, KeyName, 0, win32con.KEY_ALL_ACCESS)
    # 获取键值并输出为列表
    reg_data = list(get_keyvalue(key))
    # 遍历键值file_path
    for i in range(len(reg_data)):
        if reg_data[i][0] == 'cydgsx':
            print('已添加为自启动。。。')
            return True
        else:
            pass
    print('添加为自启动。。。')
    # 注册为自启动
    try:
        # 打开现有注册表
        key = win32api.RegOpenKey(
            keyHandle, KeyName, 0, win32con.KEY_ALL_ACCESS)
        # 设置指定项的值
        win32api.RegSetValueEx(key, file_path, 0, win32con.REG_SZ, path)
        # 关闭系统注册表中的一个项（或键）
        win32api.RegCloseKey(key)
    except Exception as e:
        print('添加失败')
        print(e)
        return False
    print('添加成功')


if __name__ == "__main__":
    open_cookie()
    try:
        _thread.start_new_thread(open_win, ())
        _thread.start_new_thread(get_datetime, ())
    except Exception as e:
        print('无法启动线程')
        print(e)

    # time_day = '2019-06-22'
    qiandao()

    os.system('pause')
