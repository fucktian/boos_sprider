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


def jobMesssage(html):
    df_jobMesssage = pd.DataFrame()
    df = pd.DataFrame()
    # with open('jobhtml.html', 'r', encoding='utf-8') as f:
    #     html = BeautifulSoup(f, 'html.parser')
    # html.list = html.find_all('div', attrs={'class': 'job-list'})
    # print(html.list)
    for i, item in enumerate(html):
        item.list = item.find_all('div', attrs={'class': 'job-primary'})
        # print(item,i,sep=',')
        for i, item in enumerate(item.list):  # 获取每个招聘条目
            # print(item, i, sep=',')
            try:
                item.list = item.find('div', attrs={'class': 'info-append clearfix'}).text.replace(' ', '').replace(
                    '\n', ' ')
                print(item.list, i, sep=',')

                df_jobMesssage['招聘职位网址'] = 'https://www.zhipin.com' + item.find('div',
                                                                                attrs={'class': 'primary-box'}).get(
                    'href'),
                df_jobMesssage['岗位名称'] = item.find('div', attrs={'class': 'job-title'}).find('span', attrs={
                    'class': 'job-name'}).text,
                df_jobMesssage['工作地及要求'] = item.find('div', attrs={'class': 'job-title'}).find('span', attrs={
                    'class': 'job-area-wrapper'}).text.strip('\n'),  #
                df_jobMesssage['公司名称'] = item.find('div', attrs={'class': 'info-company'}).text.replace(' ',
                                                                                                        '').replace(
                    '\n', ' '),
                df_jobMesssage['薪资'] = item.find('div', attrs={'class': 'job-limit clearfix'}).text.strip('').replace(
                    '\n', ' '),
                df_jobMesssage['福利'] = item.find('div', attrs={'class': 'info-append clearfix'}).text.replace(' ',
                                                                                                              '').replace(
                    '\n', ' '),
                # print(df_jobMesssage)
                df_jobMesssage.to_csv('job.csv', mode='a+', header=None, index=True, encoding='utf-8-sig', sep=',')
                df = pd.concat([df, df_jobMesssage], axis=0)
                df.to_json('jobBoss.json', orient='records', indent=1, force_ascii=False)
                print(str(i), '招聘职位写入正常')
            except:
                print(str(i), '招聘职位写入正常')
    return df


def jobRequire(html):
    # df = pd.DataFrame()
    df = {}  # 定义字典
    # # url='https://www.zhipin.com/job_detail/c3aea253a5b3b2501nJ92d-9GFBR.html'
    # url='https://www.zhipin.com/job_detail/c2b2f449e3c613a71nN72NS1FlpW.html'
    # # url='https://www.zhipin.com/job_detail/1635c904e28317c31nN63ti0FlJY.html'
    # cookie = 'Cookie: __guid=95203226.4063907470298592000.1630401055947.081; _bl_uid=tIkzmsaaz8bup1qepsempvm87k3z; wt2=Dt6B1sNjfS9mOw2rOUcWz7LnE65oG5AcG7C-7iuSGQ10DZgwjtuGdrBZlKOJt5QsEu8DWRIOSeNQ2a7qP7q1yRQ~~; lastCity=101210100; __g=-; Hm_lvt_194df3105ad7148dcf2b98a91b5e727a=1630888771,1632789052,1632907583,1632959098; acw_tc=0bdd34ba16329610479403976e01a46b6a653805d48cc356c7a1254d2d5375; __c=1632959098; __a=66278464.1630401067.1632907554.1632959098.52.6.7.47; Hm_lpvt_194df3105ad7148dcf2b98a91b5e727a=1632962530; __zp_stoken__=0138dGiMjNjETJHpLDRQ2VDBYbnMRPGxPGRFeJC8TJ0Y%2FASEDIHMxYwBwZi8AHjN%2BTxwJVQgkUkJCHRMVQ3ACZm0YMWV2U1EgOHM5WnAVdzxse017agxTPj5JZUd4Q1w1DSU7fXVbUEcKIRY%3D; __zp_sseed__=iVynj0LLIRVDsqGYxrY8A2rJBiqFMuzEYl1KvBTzD1Q=; __zp_sname__=e948d594; __zp_sts__=1632962688132; monitor_count=40'
    # # cookie ='Cookie: HMACCOUNT_BFESS=399A131593FFAEE5; BDUSS_BFESS=VpjS3U5Q1hQd3ktdkMwand3N3k1ekppN1FJSUhSc2EtdVBEMGhBaU0zSEdYbEpoRVFBQUFBJCQAAAAAAAAAAAEAAADB320FNTMzNTg5NDkzAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAMbRKmHG0SphW; BAIDUID_BFESS=DA74B922ACBBFCBDF71367A36C973898:FG=1'
    # # cookie ='set-cookie: __zp_sseed__=iVynj0LLIRVDsqGYxrY8A7QRlGL1xd7z8VDrvc0yURg=; Path=/; Domain=.zhipin.com'
    # headers = {
    #     'user-agent': 'user-agent: Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/86.0.4240.198 Safari/537.36',
    #     'Cookie': cookie,
    #     'Connection': 'keep - alive',
    #     'Accept':'Accept: image / avif, image / webp, image / apng, image / *, * / *;q = 0.8',
    # }
    # # 新闻链接
    # # session = requests.session()
    # res = requests.get(url=url, headers=headers, timeout=30)
    # res.encoding = 'utf-8'
    # res.raise_for_status()
    # res.encoding = res.apparent_encoding
    # html = BeautifulSoup(res.text, 'html.parser')
    # time.sleep(3)
    # print(html)
    # # 存入本地
    # with open('jobhtmlText.html','w',encoding='utf-8-sig') as f:#gbk,utf-8-sig\gb2312
    #     f.write(res.text)
    # with open('jobhtmlText.html', 'r', encoding='utf-8') as f:
    #     html = BeautifulSoup(f, 'html.parser')
    # html.list = html.find_all('div', attrs={'class': 'job-detail'})  # 整体框架
    for i, item in enumerate(html):
        # print(item,1,sep=',')
        item.list = item.find_all('div', attrs={'class': 'text'})[0].text.strip('').replace(' ', '')
        print(item.list, i, sep=',')
        try:
            df['招聘要求'] = item.find_all('div', attrs={'class': 'text'})[0].text.strip('\n').replace(' ', '').replace(
                '\n', ' ').replace('\r', ' ').replace('\t', ' '),  # 上级框架,
            df['公司信息'] = item.find_all('div', attrs={'class': 'job-sec company-info'})[0].text.strip('\n').replace(' ',
                                                                                                                   ''),
            # df.to_csv('job.csv', mode='a+', header=None, index=None, encoding='utf-8-sig', sep=',')
            # df.to_json('jobBoss.json', orient='records', indent=1, force_ascii=False)
            # print(df)
            print(str(i), '招聘职位写入正常')
        except:
            print(str(i), '招聘职位写入正常')

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
            try:
                # df_w = jobRequire(data[i]['招聘职位网址'])
                # print(data[i]['招聘职位网址'])
                new_worksheet.cells[i + 1, 0].value = i + 1
                new_worksheet.cells[i + 1, 1].value = data[i]['岗位名称']
                new_worksheet.cells[i + 1, 2].value = ''  # data[i]['发布日期']
                new_worksheet.cells[i + 1, 3].value = data[i]['薪资']
                new_worksheet.cells[i + 1, 4].value = data[i]['工作地及要求']
                new_worksheet.cells[i + 1, 5].value = data[i]['公司名称']
                new_worksheet.cells[i + 1, 6].value = ''  # data[i]['公司规模']
                new_worksheet.cells[i + 1, 7].value = ''  # data[i]['所属行业']
                new_worksheet.cells[i + 1, 8].value = data[i]['招聘职位网址']
                # new_worksheet.cells[i + 1, 9].value =df_w['招聘要求']#str(df_w['招聘要求'].values).strip("['").strip("']").strip('')
                new_worksheet.cells[i + 1, 10].value = ''  # data[i]['招聘公司网址']
                # new_worksheet.cells[i + 1, 11].value = df_w['公司信息']#str(df_w['公司信息'].values).strip("['").strip("']").strip('')
                new_worksheet.cells[i + 1, 12].value = ''  # data[i]['福利']
                # 修改项目
                new_worksheet.cells[i + 1, 13].value = key  # 关键字
                new_worksheet.cells[i + 1, 14].value = '20-30k' if salary == 6 else '15-20K'  # 薪资范围
                new_worksheet.cells[i + 1, 17].value = datetime.date.today()  # 薪资范围

                print(str(i), 'Excel数据写入正常')
            except:
                print(str(i), 'Excel数据写入异常')
        # 招聘公司信息获取
        # for i in range(len(self.data)):
        #     try:
        #         # 招聘公司信息获取
        #         time1 = time.time()  # 计算时长
        #         myWeb = Web(url)  # 实例化类
        #         time.sleep(0.5)
        #         html = myWeb.web_a(data[i]['招聘职位网址'])  # 'https://jobs.51job.com/all/co3836624.html')  # 实例化网址
        #         df_w = jobRequire(html)  # 获取职位需求信息
        #         print(df_w)
        #         time.sleep(2.5)
        #         new_worksheet.cells[i + 1, 9].value = df_w['招聘要求']
        #         new_worksheet.cells[i + 1, 11].value = df_w['公司信息']
        #         print(str(i), 'Excel数据-2模块写入正常')
        #         time2 = time.time()  # 计算时长
        #         print('总耗时：{}'.format(time2 - time1))
        #     except:
        #         print(str(i), 'Excel数据-2模块写入异常')

        new_worksheet.autofit()
        new_workbook.save('jobBoss.xlsx')
        new_workbook.close()
        app.quit()

    def wE_r_a(self):
        app = xw.App(visible=True, add_book=False)
        wb = app.books.open('jobBoss.xlsx')
        sh = wb.sheets['worksheet']
        # print(sh.range('i2').value)
        rng = [i for i in sh.range("i:i").value if i != None]  # 单元格内容招聘网址
        print(rng)
        # j = sh.range('a1').expand('table').rows.count
        # print(j)
        app.display_alerts = False
        app.screen_updating = False
        for i in range(len(rng) - 1):
            try:
                # 招聘公司信息获取
                time1 = time.time()  # 计算时长
                myWeb = Web(url)  # 实例化类
                time.sleep(0.5)
                html = myWeb.web_a(rng[i + 1])  # 'https://jobs.51job.com/all/co3836624.html')  # 实例化网址
                df_w = jobRequire(html)  # 获取职位需求信息
                print(df_w)
                time.sleep(2.5)
                sh.cells[i + 1, 9].value = df_w['招聘要求']
                sh.cells[i + 1, 11].value = df_w['公司信息']
                print(str(i), 'Excel数据-2模块写入正常')
                time2 = time.time()  # 计算时长
                print('总耗时：{}'.format(time2 - time1))
            except:
                print(str(i), 'Excel数据-2模块写入异常')

        sh.autofit()
        wb.save('jobBoss.xlsx')
        wb.close()
        app.quit()

    def run(self):
        pf = multiprocessing.Process(target=self.wE_r())
        pf.start()
        pf.join()

    def run_a(self):
        pf = multiprocessing.Process(target=self.wE_r_a())
        pf.start()
        pf.join()


class Web:
    def __init__(self, url):
        self.url = url

    # 获取招聘职位信息
    def web(self):
        driver.back()
        # driver.refresh()
        time.sleep(0.5)
        driver.get(self.url)  # 加载网址
        time.sleep(1.5)
        source = driver.page_source  # 页面内容实例化
        html = BeautifulSoup(source, 'html.parser')  # 获取页面内容
        html.list = html.find_all('div', attrs={'class': 'job-list'})
        # with open('jobhtml.html','w',encoding='utf-8-sig') as f:#gbk,utf-8-sig\gb2312
        #     f.write(source)
        # print(html)
        return html.list

        # 获取招聘要求和公司信息

    def web_a(self, url):
        driver.back()
        # driver.refresh()
        time.sleep(0.5)
        driver.get(url)  # 加载网址
        time.sleep(1.5)
        source = driver.page_source  # 页面内容实例化
        html = BeautifulSoup(source, 'html.parser')  # 获取页面内容
        html.list = html.find_all('div', attrs={'class': 'job-detail'})  # 整体框架
        # with open('jobhtmlText.html','w',encoding='utf-8-sig') as f:#gbk,utf-8-sig\gb2312
        #     f.write(source)
        # print(html)
        return html.list


df = pd.DataFrame()  # 定义    全局变量
key = '物流管理'  # 物流经理#物流运营
salary = '6'  # 5表示15-20K，6表示20-30k
if __name__ == '__main__':
    # jobMesssage()
    # jobRequire()
    # opt = ChromeOptions()  # 创建chrome参数
    # opt.headless = False  # 显示浏览器
    # driver = Chrome(options=opt)  # 浏览器实例化
    # # driver=webdriver.Chrome()
    # driver.set_window_size(300, 700)
    # url='https://www.zhipin.com/job_detail/?query=%E7%89%A9%E6%B5%81%E8%BF%90%E8%90%A5&city=101210100&industry=&position='
    # url_b='https://www.zhipin.com/job_detail/63a31859fef2dbbc1nJy0tS8EFJY.html'
    # # 招聘公司信息获取
    # myWeb = Web(url)  # 实例化类
    # time.sleep(0.3)
    # html = myWeb.web_a(url_b)  # 'https://jobs.51job.com/all/co3836624.html')  # 实例化网址
    # df5 = jobRequire(html)  # 获取职位需求信息
    # print(df5)
    # time.sleep(0.5)

    opt = ChromeOptions()  # 创建chrome参数
    # 不加载图片
    prefs = {"profile.managed_default_content_settings.images": 2}
    opt.add_experimental_option("prefs", prefs)
    opt.headless = False  # 显示浏览器
    driver = Chrome(options=opt)  # 浏览器实例化
    # driver=webdriver.Chrome()
    driver.set_window_size(300, 700)
    url = 'https://www.zhipin.com/c101210100/y_6/?query=%E7%89%A9%E6%B5%81%E8%BF%90%E8%90%A5&ka=sel-salary-6'
    '''
    for i in range(3):  # +str(i);key=
        try:
            print(str(i), '获取第{}页数据'.format(i + 1))
            url='https://www.zhipin.com/c101210100/y_'+salary+'/?query='+key+'&city=101210100&industry=&position=&ka=sel-salary-'+salary+'&page='+str(i+1)+'&ka=page-'+str(i+1)
            print(url)
            #'https://www.zhipin.com/job_detail/?query=%E7%89%A9%E6%B5%81%E8%BF%90%E8%90%A5&city=101210100&industry=&position='
            #'https://www.zhipin.com/c101210100/y_6/?query=%E7%89%A9%E6%B5%81%E8%BF%90%E8%90%A5&ka=sel-salary-6'
            #'https://www.zhipin.com/c101210100/y_5/?query=%E7%89%A9%E6%B5%81%E8%BF%90%E8%90%A5&ka=sel-salary-5'
            #‘https://www.zhipin.com/c101210100/y_5/?query=%E7%89%A9%E6%B5%81%E8%BF%90%E8%90%A5&page=2&ka=page-2’
            time1 = time.time()  # 计算时长
            # 获取招聘职位信息
            myWeb=Web(url)
            html=myWeb.web()#获取招聘岗位信息
            # html=myWeb.web_a('https://www.zhipin.com/job_detail/c2b2f449e3c613a71nN72NS1FlpW.html')# 获取招聘要求和公司信息
            time.sleep(0.5)
            # print(html)
            df1 = jobMesssage(html)
            df = pd.concat([df1, df], axis=0)
            df.to_json('jobBoss.json', orient='records', indent=1, force_ascii=False)
            # url_b = str(df1['招聘公司网址'].values).strip("['").strip("']").strip('')
            # print(url_b)
            # # 招聘公司信息获取
            # myWeb = Web(url)  # 实例化类
            # time.sleep(0.3)
            # html = myWeb.web_a(url_b)  # 'https://jobs.51job.com/all/co3836624.html')  # 实例化网址
            # df2 = jobRequire(html)  # 获取职位需求信息
            # print(df2)
            # time.sleep(0.5)
            #
            # df3 = pd.concat([df1, df2], axis=1)
            # df3.to_csv('job.csv', mode='a+', header=None, index=None, encoding='utf-8-sig', sep=',')
            # df = pd.concat([df, df3], axis=0)
            # print(df)
            # df.to_json('jobBoss.json', orient='records', indent=1, force_ascii=False)
            # time.sleep(0.5)
            time2 = time.time()  # 计算时长
            print(str(i), '数据正常'.format(i + 1))
            print('总耗时：{}'.format(time2 - time1))
        except:
            print(str(i), '数据异常'.format(i + 1))

    # 写入excel
    with open('jobBoss.json', 'r', encoding='utf-8') as f:
        data = json.load(f)
        # print(data)
        myWe = writeExcel(data)  # 写入excel
        myWe.run()  # 执行多线程
    '''
    # 写入excel_a
    with open('jobBoss.json', 'r', encoding='utf-8') as f:
        data = json.load(f)
        # print(data)
        myWe = writeExcel(data)  # 写入excel
        myWe.run_a()  # 执行多线程

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