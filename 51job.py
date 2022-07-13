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


# 获取职位信息
def jobMesssage(item):
    df = pd.DataFrame()
    item.list = item.find_all('a', attrs={'class': 'el'})  # 获取招聘岗位信息
    for i, item in enumerate(item.list):
        try:
            df['招聘职位网址'] = item.get('href'),
            df['岗位名称'] = item.find_all('span')[0].text,
            df['发布日期'] = item.find_all('span')[1].text,
            df['薪资'] = item.find_all('span')[2].text,  #
            df['工作地及要求'] = item.find_all('span')[3].text,  #
            # df_all=pd.concat([df,df_all],axis=1)
            item.list = item.find_all('p', attrs={'class': 'tags'})
            for i, item.list in enumerate(item.list):
                df['福利'] = item.list.get('title'),  #
            print(str(i), '招聘职位写入正常')
        except:
            print(str(i), '招聘职位写入正常')
    return df


# 获取职位对应公司信息
def jobFirm(item):
    df = pd.DataFrame()
    item.list = item.find_all('div', attrs={'class': 'er'})  # 获取招聘公司信息
    for i, item in enumerate(item.list):
        # print(item,i,sep=',')
        # print(item.find_all('p')[1].text)
        try:
            df['招聘公司网址'] = item.find('a').get('href'),
            df['公司名称'] = item.find('a').text,
            df['公司规模'] = item.find_all('p')[0].text,
            df['所属行业'] = item.find_all('p')[0].text,
            print(str(i), '招聘公司写入正常')
        except:
            print(str(i), '招聘公司写入异常')
    return df


# 职位要求
def jobRequire(html):
    df = pd.DataFrame()
    # with open('jobhtmlText.html', 'r', encoding='utf-8') as f:
    #     html = BeautifulSoup(f, 'html.parser')
    # html.list = html.find_all('div', attrs={'class': 'tHeader tHjob'})
    html.list = html.find_all('div', attrs={'class': 'tCompany_main'})
    # print(html.list)
    for i, item in enumerate(html.list):
        try:
            # contactInf=item.find_all('div', attrs={'class': 'tBorderTop_box'})[1].find('span').text.strip('') #联系方式
            # officeAddress=item.find_all('div', attrs={'class': 'tBorderTop_box'})[1].find('p').text#上班地址
            jobRequir_a = item.find('div', attrs={'class': 'tBorderTop_box'}).text.strip('').replace('\n', '').replace(
                '\t', '') \
                .replace(' ', '')  # 任职要求
            # print(jobRequir_a, i, sep='')
            item.list = item.find('div', attrs={'class': 'tBorderTop_box'}).find_all('p')
            jobRequir = []  # 职位要求
            for i, item in enumerate(item.list):
                jobRequir.append(item.text.strip('') + '\n')
                jobRequirText = ''.join(jobRequir)
                # print(jobRequirText)
                # print(jobRequirText.find('任职要求'))
                if jobRequirText.find('任职要求') > 0:
                    df['招聘要求'] = jobRequirText,
                else:
                    df['招聘要求'] = jobRequir_a,
            # print(df)
            print(str(i), '职位信息写入正常')
        except:
            print(str(i), '职位信息写入异常')
    return df


# 招聘公司信息获取
def firmMeessage(html):
    df = pd.DataFrame()
    # with open('jobhtmlText.html', 'r', encoding='utf-8') as f:
    #     html = BeautifulSoup(f, 'html.parser')
    html.list = html.find_all('div', attrs={'class': 'tCompany_full'})
    # print(html.list)
    for i, item in enumerate(html.list):
        item.list = item.find_all('div', attrs={'class': 'tBorderTop_box'})
        # print(item.list[0].text.strip('').replace('\n', '').replace('\t', '').replace(' ', ''))
        # for i, item in enumerate(item.list):
        #     print(item.text,i,sep='')
        try:
            df['公司信息'] = item.list[0].text.strip('').replace('\n', '').replace('\t', '').replace(' ', ''),
            # print(df)
            print(str(i), '公司信息写入正常')
        except:
            print(str(i), '公司信息写入异常')

    return df


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
            new_worksheet.cells[i + 1, 0].value = i + 1
            new_worksheet.cells[i + 1, 1].value = data[i]['岗位名称']
            new_worksheet.cells[i + 1, 2].value = data[i]['发布日期']
            new_worksheet.cells[i + 1, 3].value = data[i]['薪资']
            new_worksheet.cells[i + 1, 4].value = data[i]['工作地及要求']
            new_worksheet.cells[i + 1, 5].value = data[i]['公司名称']
            new_worksheet.cells[i + 1, 6].value = data[i]['公司规模']
            new_worksheet.cells[i + 1, 7].value = data[i]['所属行业']
            new_worksheet.cells[i + 1, 8].value = data[i]['招聘职位网址']
            new_worksheet.cells[i + 1, 9].value = data[i]['招聘要求']
            new_worksheet.cells[i + 1, 10].value = data[i]['招聘公司网址']
            new_worksheet.cells[i + 1, 11].value = data[i]['公司信息']
            new_worksheet.cells[i + 1, 12].value = data[i]['福利']
            # 修改项目
            new_worksheet.cells[i + 1, 13].value = key  # 关键字
            new_worksheet.cells[i + 1, 14].value = '20-30k' if salary == '09' else '15-20K'  # 薪资范围
            new_worksheet.cells[i + 1, 17].value = datetime.date.today()  # 薪资范围

            print(str(i), 'Excel数据写入正常')
        new_worksheet.autofit()
        new_workbook.save('job51_1.xlsx')
        new_workbook.close()
        app.quit()

    def run(self):
        pf = multiprocessing.Process(target=self.wE_r())
        pf.start()
        pf.join()


class Web:
    def __init__(self, url):
        self.url = url

    def web(self):
        # with open('jobhtml.html', 'r', encoding='utf-8') as f:
        # job_url = 'https://search.51job.com/list/080200,000000,0000,00,9,99,%25E7%2589%25A9%25E6%25B5%2581,2,1.html?'
        driver.back()
        time.sleep(0.3)
        driver.get(self.url)  # 加载网址
        time.sleep(1)
        source = driver.page_source  # 页面内容实例化
        html = BeautifulSoup(source, 'html.parser')  # 获取页面内容
        html.list = html.find_all('div', attrs={'class': 'j_joblist'})
        return html.list

    # 招聘需求信息获取
    def web_a(self, url):
        # job_url = 'https://jobs.51job.com/hangzhou/119721744.html?s=sou_sou_soulb&t=0_0'
        driver.back()
        time.sleep(0.3)
        driver.get(url)  # 加载网址
        time.sleep(1.2)
        source = driver.page_source  # 页面内容实例化
        html = BeautifulSoup(source, 'html.parser')  # 获取页面内容
        # with open('jobhtmlText.html','w',encoding='utf-8-sig') as f:#gbk,utf-8-sig\gb2312
        #     f.write(source)
        # print(html)
        return html

    # 招聘公司信息获取
    def web_b(self, url):
        # job_url = 'https://jobs.51job.com/all/co3836624.html'
        driver.back()
        time.sleep(0.5)
        driver.get(url)  # 加载网址
        time.sleep(1.2)
        source = driver.page_source  # 页面内容实例化
        html = BeautifulSoup(source, 'html.parser')  # 获取页面内容
        # print(html)
        # with open('jobhtmlText.html','w',encoding='utf-8-sig') as f:#gbk,utf-8-sig\gb2312
        #     f.write(source)
        # print(html)
        return html


key = '质量管理工程师'  # 质量管理#质量主管
salary = '08'  # 08表示1.5-20K，09表示20-30k
if __name__ == "__main__":

    opt = ChromeOptions()  # 创建chrome参数
    opt.headless = False  # 显示浏览器
    driver = Chrome(options=opt)  # 浏览器实例化
    # job_url = 'https://search.51job.com/list/080200,000000,0000,00,9,99,%25E7%2589%25A9%25E6%25B5%2581,2,1.html?'
    # 杭州，2-3万'https://search.51job.com/list/080200,000000,0000,00,9,09,%25E7%2589%25A9%25E6%25B5%2581%25E8%25BF%2590%25E8%2590%25A5,2,1.html?lang=c&postchannel=0000&workyear=99&cotype=99&degreefrom=99&jobterm=99&companysize=99&ord_field=0&dibiaoid=0&line=&welfare='
    # 杭州1.5-2'https://search.51job.com/list/080200,000000,0000,00,9,08,%25E7%2589%25A9%25E6%25B5%2581%25E8%25BF%2590%25E8%2590%25A5,2,1.html?lang=c&postchannel=0000&workyear=99&cotype=99&degreefrom=99&jobterm=99&companysize=99&ord_field=0&dibiaoid=0&line=&welfare='

    # # 招聘需求信息获取
    # myWeb = Web(job_url)  # 实例化类
    # time.sleep(0.2)
    # html = myWeb.web_a('https://jobs.51job.com/hangzhou-scq/125683481.html?s=sou_sou_soulb&t=0_0')  # 'https://jobs.51job.com/hangzhou/135496109.html?s=sou_sou_soulb&t=0_0') # 实例化网址
    # # df4 = jobRequire(html)  # 获取职位需求信息
    # df4 = jobRequire()
    # print(df4)
    # time.sleep(0.3)

    # 取前三页数据
    df = pd.DataFrame()  # 定义pands整理表格
    for i in range(12):
        try:  # '+str(i+1)+'#08表示1.5-20K，09表示20-30k#1表示近三天，2表示近一周
            print(str(i), '获取第{}页数据'.format(i + 1))
            job_url = 'https://search.51job.com/list/010000,000000,0000,00,2,' + salary + ',' + key + ',2,' + str(
                i + 1) + '.html?lang=c&postchannel=0000&workyear=99&cotype=99&degreefrom=99&jobterm=99&companysize=99&ord_field=1&dibiaoid=0&line=&welfare='
            print(job_url)
            # 'https://search.51job.com/list/080200,000000,0000,00,2,09,%25E7%2589%25A9%25E6%25B5%2581%25E7%25AE%25A1%25E7%2590%2586,2,1.html?           lang=c&postchannel=0000&workyear=99&cotype=99&degreefrom=99&jobterm=99&companysize=99&ord_field=1&dibiaoid=0&line=&welfare='
            # 'https://search.51job.com/list/080200,000000,0000,00,9,99,%25E7%2589%25A9%25E6%25B5%2581%25E8%25BF%2590%25E8%2590%25A5,2,1.html?lang=c&postchannel=0000&workyear=99&cotype=99&degreefrom=99&jobterm=99&companysize=99&ord_field=0&dibiaoid=0&line=&welfare='
            # with open('jobhtml.html', 'r', encoding='utf-8') as f:
            #     html = BeautifulSoup(f, 'html.parser')
            #     html.list = html.find_all('div', attrs={'class': 'j_joblist'})
            time1 = time.time()  # 计算时长
            myWeb = Web(
                job_url)  # 实例化类  # 'https://jobs.51job.com/hangzhou-yhq/135494019.html?s=sou_sou_soulb&t=0_0')  # 实例化网址
            time.sleep(1)
            html = myWeb.web()
            # print(html)
            for i, item in enumerate(html):
                # print(item,i,sep=',')
                item.list = item.find_all('div', attrs={'class': 'e'})  # 获取每个招聘岗位条目
                for i, item in enumerate(item.list):
                    df1 = jobMesssage(item)  # 获取岗位
                    # print(df1['招聘职位网址'])
                    df2 = jobFirm(item)  # 获取公司
                    url = str(df1['招聘职位网址'].values).strip("['").strip("']").strip('')
                    print(url)
                    url_b = str(df2['招聘公司网址'].values).strip("['").strip("']").strip('')
                    print(url_b)

                    # 招聘需求信息获取
                    myWeb = Web(job_url)  # 实例化类
                    time.sleep(0.3)
                    html = myWeb.web_a(
                        url)  # 'https://jobs.51job.com/hangzhou/135496109.html?s=sou_sou_soulb&t=0_0') # 实例化网址
                    df4 = jobRequire(html)  # 获取职位需求信息
                    print(df4)
                    time.sleep(0.5)

                    # 招聘公司信息获取
                    myWeb = Web(job_url)  # 实例化类
                    time.sleep(0.3)
                    html = myWeb.web_b(url_b)  # 'https://jobs.51job.com/all/co3836624.html')  # 实例化网址
                    df5 = firmMeessage(html)  # 获取职位需求信息
                    print(df5)
                    time.sleep(0.5)

                    df3 = pd.concat([df1, df2], axis=1)
                    df6 = pd.concat([df3, df4], axis=1)
                    df7 = pd.concat([df5, df6], axis=1)
                    df7.to_csv('job.csv', mode='a+', header=None, index=None, encoding='utf-8-sig', sep=',')
                    df = pd.concat([df, df7], axis=0)
                    print(df)
                    df.to_json('jobGain.json', orient='records', indent=1, force_ascii=False)
                    time.sleep(0.5)
            time.sleep(0.5)
            print(str(i), '数据正常'.format(i + 1))

            time2 = time.time()  # 计算时长
            print('总耗时：{}'.format(time2 - time1))
        except:
            print(str(i), '数据异常'.format(i + 1))

    # key = '物流管理'  # 物流经理#物流运营
    # salary = '08'  # 08表示1.5-20K，09表示20-30k
    with open('jobGain.json', 'r', encoding='utf-8') as f:
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