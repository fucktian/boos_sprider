import requests
from bs4 import BeautifulSoup
import datetime
import json
import xlwings as xw
from selenium import webdriver
import time
import pandas as pd
from selenium.webdriver import Chrome, ChromeOptions, ActionChains
from selenium.webdriver.common.keys import Keys
import csv
import multiprocessing
import os


# 获取招聘职位信息
def jobMesssage(html):
    df_jobMesssage = pd.DataFrame()
    df = pd.DataFrame()
    # with open('jobhtml.html', 'r', encoding='utf-8') as f:
    #     html = BeautifulSoup(f, 'html.parser')
    # html.list = html.find_all('div', attrs={'class': 'left-list-box'})
    for i, item in enumerate(html):
        item.list = item.find_all('div', attrs={'class': 'job-detail-box'})
        for i, item in enumerate(item.list):
            # print(item, i, sep=',')
            # print(item.find('div', attrs={'class': 'job-detail-header-box'}).find('span', attrs={'class': 'job-salary'}).text,i,sep=',')
            try:
                df_jobMesssage['招聘职位网址'] = item.find('a', attrs={'data-nick': 'job-detail-job-info'}).get('href'),
                df_jobMesssage['岗位名称'] = item.find('a', attrs={'data-nick': 'job-detail-job-info'}).find('div', attrs={
                    'class': 'job-title-box'}).text.strip('').replace('\n', '').replace('\t', ''),
                df_jobMesssage['工作地及要求'] = item.find('a', attrs={'data-nick': 'job-detail-job-info'}).find('div',
                                                                                                           attrs={
                                                                                                               'class': 'job-labels-box'}).text.strip(
                    '').replace('\n', '').replace('\t', ''),  #
                df_jobMesssage['公司名称'] = item.find('div', attrs={'data-nick': 'job-detail-company-info'}).find('div',
                                                                                                               attrs={
                                                                                                                   'class': 'job-company-info-box'}).text.strip(
                    '').replace('\n', '').replace('\t', '')
                df_jobMesssage['薪资'] = item.find('div', attrs={'class': 'job-detail-header-box'}).find('span', attrs={
                    'class': 'job-salary'}).text

                # print(df_jobMesssage)
                df_jobMesssage.to_csv('job.csv', mode='a+', header=None, index=True, encoding='utf-8-sig', sep=',')
                df = pd.concat([df, df_jobMesssage], axis=0)
                # df.to_json('jobliepin.json', orient='records', indent=1, force_ascii=False)

                print(str(i), '招聘职位写入正常')
            except:
                print(str(i), '招聘职位写入正常')

    return df


# 获取招聘要求和公司信息
def jobRequire(url):
    df = {}  # 定义字典
    # url='https://www.liepin.com/a/29686195.shtml?d_sfrom=search_prime&d_ckId=c8f01cee484fdfafc8e1e5d047a1e1d1&d_curPage=0&d_pageSize=40&d_headId=6ae8e76ae415c8d307347eef4182b4e4&d_posi=38'
    cookie = 'Cookie: __uuid=1632571874000.95; __s_bid=11011704223d5f9c92ff5bd3e81bc8334a74; __tlog=1632611231431.79%7C00000000%7C00000000%7C00000000%7C00000000; Hm_lvt_a2647413544f5a04f00da7eee0d5e200=1632571900,1632611231; Hm_lpvt_a2647413544f5a04f00da7eee0d5e200=1632615070; __session_seq=12; __uv_seq=12'
    headers = {
        'user-agent': 'User-Agent: Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/92.0.4515.159 Safari/537.36',
        'Cookie': cookie,
        'Connection': 'keep - alive'
    }
    # 新闻链接
    # session = requests.session()
    res = requests.get(url=url, headers=headers, timeout=30)
    res.encoding = 'utf-8'
    res.raise_for_status()
    res.encoding = res.apparent_encoding
    html = BeautifulSoup(res.text, 'html.parser')
    time.sleep(0.1)
    # print(html)
    # 存入本地
    # with open('jobhtmlText.html','w',encoding='utf-8-sig') as f:#gbk,utf-8-sig\gb2312
    #     f.write(res.text)
    # with open('jobhtmlText.html', 'r', encoding='utf-8') as f:
    #     html = BeautifulSoup(f, 'html.parser')
    html.list = html.find_all('content')  # 整体框架
    for i, item in enumerate(html):
        # item.list = item.find_all('section', attrs={'class': 'company-intro-container'})[0].text#上级框架
        # print(item.list)
        try:
            df['招聘要求'] = item.find_all('section', attrs={'class': 'job-intro-container'})[0].text.strip('\n'),
            df['公司信息'] = item.find_all('section', attrs={'class': 'company-intro-container'})[0].text.strip('\n'),
            # df.to_csv('job.csv', mode='a+', header=None, index=None, encoding='utf-8-sig', sep=',')
            # df.to_json('jobliepin.json', orient='records', indent=1, force_ascii=False)
            print(df)
            print(str(i), '招聘职位写入正常')
        except:
            print(str(i), '招聘职位写入正常')

    return df


class Web:
    def __init__(self, url):
        self.url = url

    # 获取招聘职位信息
    def web(self):
        driver.back()
        time.sleep(0.3)
        driver.get(self.url)  # 加载网址
        time.sleep(1)
        source = driver.page_source  # 页面内容实例化
        html = BeautifulSoup(source, 'html.parser')  # 获取页面内容
        html.list = html.find_all('div', attrs={'class': 'left-list-box'})
        # with open('jobhtml.html','w',encoding='utf-8-sig') as f:#gbk,utf-8-sig\gb2312
        #     f.write(source)
        # print(html)
        return html.list

    # 获取招聘要求和公司信息
    def web_a(self, url):
        driver.back()
        time.sleep(0.3)
        driver.get(url)  # 加载网址
        time.sleep(1)
        source = driver.page_source  # 页面内容实例化
        html = BeautifulSoup(source, 'html.parser')  # 获取页面内容
        html.list = html.find_all('content')  # 整体框架
        # with open('jobhtmlText.html','w',encoding='utf-8-sig') as f:#gbk,utf-8-sig\gb2312
        #     f.write(source)
        # print(html)
        return html.list


class writeExcel:
    def __init__(self, data):
        self.data = data
        # print(data)

    def wE_r(self):
        app = xw.App(visible=False, add_book=False)
        new_workbook = xw.Book()
        new_worksheet = new_workbook.sheets.add('worksheet')
        app.display_alerts = False
        app.screen_updating = False
        title = ["序号", "岗位名称", "发布日期", "薪资", "工作地及要求", "公司名称", "公司规模", "所属行业", "招聘职位网址", "招聘要求",
                 "招聘公司网址", "公司信息", '福利', '关键字', '薪资范围', '标记', '顺序', '记录日期']
        new_worksheet['A1'].value = title

        for i in range(len(self.data)):
            try:
                df_w = jobRequire(data[i]['招聘职位网址'])
                print(data[i]['招聘职位网址'])
                if i%9==8:
                    time.sleep(20)#每取8个停下8秒应对反扒
                else:
                    time.sleep(0.2)

                new_worksheet.cells[i + 1, 0].value = i + 1
                new_worksheet.cells[i + 1, 1].value = data[i]['岗位名称']
                new_worksheet.cells[i + 1, 2].value = ''  # data[i]['发布日期']
                new_worksheet.cells[i + 1, 3].value = data[i]['薪资']
                new_worksheet.cells[i + 1, 4].value = data[i]['工作地及要求']
                new_worksheet.cells[i + 1, 5].value = data[i]['公司名称']
                new_worksheet.cells[i + 1, 6].value = ''  # data[i]['公司规模']
                new_worksheet.cells[i + 1, 7].value = ''  # data[i]['所属行业']
                new_worksheet.cells[i + 1, 8].value = data[i]['招聘职位网址']
                new_worksheet.cells[i + 1, 9].value = df_w[
                    '招聘要求']  # str(df_w['招聘要求'].values).strip("['").strip("']").strip('')
                new_worksheet.cells[i + 1, 10].value = ''  # data[i]['招聘公司网址']
                new_worksheet.cells[i + 1, 11].value = df_w[
                    '公司信息']  # str(df_w['公司信息'].values).strip("['").strip("']").strip('')
                new_worksheet.cells[i + 1, 12].value = ''  # data[i]['福利']
                # 修改项目
                new_worksheet.cells[i + 1, 13].value = key  # 关键字
                new_worksheet.cells[i + 1, 14].value = salary  # 薪资范围
                new_worksheet.cells[i + 1, 17].value = datetime.date.today()  # 薪资范围

            except:
                print(str(i), 'Excel数据写入异常')

        new_worksheet.autofit()
        new_workbook.save('jobliepin_m.xlsx')
        new_workbook.close()
        app.quit()

    def run(self):
        pf = multiprocessing.Process(target=self.wE_r())
        pf.start()
        pf.join()


df = pd.DataFrame()  # 定义    全局变量
key = '质量管理'  # 职位名称
salary = '10$20'  # 20$40#10$20
if __name__ == "__main__":
    # jobRequire()
    opt = ChromeOptions()  # 创建chrome参数
    opt.headless = False  # 显示浏览器
    driver = Chrome(options=opt)  # 浏览器实例化
    # driver=webdriver.Chrome()
    driver.set_window_size(300, 700)
    for i in range(12):  # +str(i);key=
        try:
            print(str(i), '获取第{}页数据'.format(i + 1))

            job_url = 'https://www.liepin.com/zhaopin/?headId=9f577a23fdb5d9437efff7679944c610&key=' + str(
                key) + '&dq=410&salary=' + salary + '&pubTime=240&currentPage=' + str(i)
            print(job_url)
            'https://www.liepin.com/zhaopin/?headId=12baac27653545ffceb6a268fc0c82aa&ckId=12baac27653545ffceb6a268fc0c82aa&key=%E7%89%A9%E6%B5%81%E8%BF%90%E8%90%A5&dq=410&salary=20$40&pubTime=240&currentPage=1'
            'https://www.liepin.com/zhaopin/?headId=12baac27653545ffceb6a268fc0c82aa&key=%E7%89%A9%E6%B5%81%E8%BF%90%E8%90%A5&dq=410&salary=10$20&pubTime=240'
            'https://www.liepin.com/zhaopin/?headId=9f577a23fdb5d9437efff7679944c610&key=%E7%89%A9%E6%B5%81%E7%AE%A1%E7%90%86&dq=410&salary=20$40&pubTime=240'
            # job_url_a='https://www.liepin.com/a/30216633.shtml?d_sfrom=search_prime&d_ckId=10e193c94fdc8095c14815c02246e6e7&d_curPage=0&d_pageSize=40&d_headId=6ae8e76ae415c8d307347eef4182b4e4&d_posi=2'
            time1 = time.time()  # 计算时长

            # 获取招聘职位信息
            myWeb = Web(job_url)  # 实例化类
            html = myWeb.web()  # 招聘要求和公司信息
            time.sleep(0.5)
            # print(html)
            df1 = jobMesssage(html)
            df = pd.concat([df1, df], axis=0)
            df.to_json('jobliepin.json', orient='records', indent=1, force_ascii=False)

            time2 = time.time()  # 计算时长
            print(str(i), '数据正常'.format(i + 1))
            print('总耗时：{}'.format(time2 - time1))
        except:
            print(str(i), '数据异常'.format(i + 1))

    # 写入excel
    with open('jobliepin.json', 'r', encoding='utf-8') as f:
        data = json.load(f)
        # print(data)
        myWe = writeExcel(data)  # 写入excel
        myWe.run()  # 执行多线程

    try:  # 关闭后台浏览器
        driver.close()
        driver.quit()
        os.system('taskkill /F /IM chromedriver.exe')  # 关闭进程浏览器
        sreach_windows = driver.current_window_handle
        # 获得当前所有打开的窗口的句柄
        all_handles = driver.window_handles
        for handle in all_handles:
            driver.switch_to.window(handle)
            driver.close()
            time.sleep(1.2)
    except:
        print('已完后台毕浏览器')