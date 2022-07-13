import asyncio, random
from pyppeteer import launch
from lxml import etree
import pandas as pd


class ss_xz(object):
    def __init__(self):
        self.data_list = list()

    def screen_size(self):
        """使用tkinter获取屏幕大小"""
        import tkinter
        tk = tkinter.Tk()
        width = tk.winfo_screenwidth()
        height = tk.winfo_screenheight()
        tk.quit()
        return width, height

    # width, height = 1366, 768
    async def main(self):
        try:
            browser = await launch(headless=False,
                                   args=['--disable-infobars', '--window-size=1366,768', '--no-sandbox'])

            page = await browser.newPage()
            width, height = self.screen_size()
            await page.setViewport({'width': width, 'height': height})
            await page.goto(
                'https://www.zhipin.com/?ka=city-sites-101190100')
            await page.evaluateOnNewDocument(
                '''() =>{ Object.defineProperties(navigator, { webdriver: { get: () => false } }) }''')
            await asyncio.sleep(5)
            # 查询岗位
            await page.type(
                '#wrap > div.column-search-panel > div > div > div.search-form > form > div.search-form-con > p > input',
                '质量管理', {'delay': self.input_time_random() - 50})
            await asyncio.sleep(2)
            # 点击搜索
            await page.click('#wrap > div.column-search-panel > div > div > div.search-form > form > button')


            # print(await page.content())
            # 获取页面内容
            i = 0
            while True:
                await asyncio.sleep(4)
                content = await page.content()
                html = etree.HTML(content)
                # 解析内容
                self.parse_html(html)
                # 翻页
                await page.click('#main > div > div.job-list > div.page > a.next')
                i += 1
                print(i)
                # 跑一天的数据 i >= 3 够用了，跑很多天数据，就建议把3变为更大的数字，也就是多抓几页的数据
                if i >= 30:
                    break
            df = pd.DataFrame(self.data_list)
            # df['职位'] = df.职位.str.extract(r'[(.*?)]', expand=True)

            df.to_excel('C:/Users/15695171918/Desktop/nj-job.xlsx', index=False)
            print(df)

        except Exception as a:
            print(a)


    def input_time_random(self):
        return random.randint(100, 151)

    def parse_html(self, html):

        li_list = html.xpath('//div[@class="job-list"]//ul/li')
        data_df = []
        for li in li_list:
            # 获取文本
            items = {}
            items['职位'] = li.xpath('.//span[@class="job-name"]/a/@title')
            items['薪酬'] = li.xpath('.//div[@class="job-limit clearfix"]/span/text()')
            items['地区'] = li.xpath('.//span[@class="job-area"]/text()')
            items['公司名称'] = li.xpath('.//div[@class="info-company"]//h3[@class="name"]/a/text()')
            items['公司类型'] = li.xpath('.//div[@class="info-company"]/div[@class="company-text"]//p/a/text()')
            items['公司规模'] = li.xpath('.//div[@class="info-company"]/div[@class="company-text"]//p/text()')
            items['福利'] = li.xpath('.//div[@class="info-desc"]/text()')
            items['工作经验及学历要求'] = li.xpath('.//div[@class="job-limit clearfix"]//p/text()')

            span_list = li.xpath('.//div[@class="tags"]')
            for span in span_list:
                items['技能要求'] = span.xpath('./span/text()')
            # print(items)
                self.data_list.append(items)




    def run(self):
        asyncio.get_event_loop().run_until_complete(self.main())


if __name__ == '__main__':

    comment = ss_xz()
    comment.run()
