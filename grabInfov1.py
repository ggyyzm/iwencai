# -*- coding: utf-8 -*-
import time
import json
import math
from base64 import decodebytes
from selenium import webdriver
from selenium.webdriver import ChromeOptions, Chrome
from selenium.webdriver import ActionChains
from selenium.webdriver.support.select import Select


class GrabStockInfo():
    def __init__(self, keywords, driverPath, showWindow):
        self.baseUrl = "http://www.iwencai.com/stockpick/search?typed=1&preParams=&ts=1&f=1&qs=result_original&selfsectsn=&querytype=stock&searchfilter=&tid=stockpick&w="
        self.keywords = ""
        for key in keywords:
            self.keywords += str(key) + " "

        self.driver = None
        self.driverPath = driverPath
        self.showWindow = showWindow
        # 启动chrome 并完成初始化搜索
        self._start()

    def _start(self):
        options = ChromeOptions()
        # 添加无界面参数
        options.add_argument('--headless')
        if self.showWindow:
            self.driver = Chrome(self.driverPath)
        else:
            self.driver = Chrome(self.driverPath, options=options)
        self.driver.get(self.baseUrl + self.keywords.strip())
        time.sleep(0.5)
        # islogin = self.load_cookies()
        # time.sleep(0.5)
        # self.login()
        #
        # valid = None
        # with open('valid.txt', 'r') as f:
        #     valid = f.readline()
        # self.do_login('mx_536962324', 'gui9624983', valid)
        #
        # islogin = self.load_cookies()

    # 判断是某个股票信息 还是 多个股票信息
    def _judgeIsSingleStock(self):
        try:
            self.driver.find_element_by_id('newworkWrap')
            return False
        except:
            return True

    # 加载本地cookie
    def load_cookies(self):
        try:
            with open('cookies.json', 'r', encoding='utf-8') as f:
                listCookies = json.loads(f.read())
                flag = False
                for cookie in listCookies:
                    if cookie['name'] == 'escapename':
                        flag = True
                        cur_time = int(time.time())
                        # 如果当前时间大于过期时间， 就需要重新登录
                        if cur_time > int(cookie['expiry']):
                            return False
                        else:
                            break
                if flag is False:
                    return flag
                for cookie in listCookies:
                    self.driver.add_cookie({
                        'domain': cookie['domain'],
                        'httpOnly': cookie['httpOnly'],
                        'name': cookie['name'],
                        'path': cookie['path'],
                        'secure': cookie['secure'],
                        'value': cookie['value'],
                    })
                self.driver.refresh()
                # 先设置当前页面为70页
                el = self.driver.find_element_by_id('showPerpage').find_element_by_tag_name('select')
                Select(el).select_by_value('70')
                time.sleep(1)
                return True
        except FileNotFoundError:
            return False

    def save_cookies(self):
        cookies = self.driver.get_cookies()
        jsonCookies = json.dumps(cookies)
        with open('cookies.json', 'w', encoding='utf-8') as f:
            f.write(jsonCookies)

    def login(self):
        ActionChains(self.driver).click(self.driver.find_element_by_id('topbarCon').find_element_by_tag_name('a')).perform()
        time.sleep(0.5)
        # 切换到login_iframe
        self.driver.switch_to.frame('login_iframe')
        ActionChains(self.driver).click(self.driver.find_element_by_link_text('账号密码登录')).perform()
        valid_imageBase64 = self.driver.find_element_by_tag_name('img').screenshot_as_base64
        with open('./veri.png', 'wb') as img:
             img.write(decodebytes(bytes(valid_imageBase64, encoding='utf-8')))
        # return valid_imageBase64

    def do_login(self, username, password, valid):
        if username == '' or password == '' or valid == '':
            return False, '不能为空'
        self.driver.find_element_by_id('uname').clear()
        self.driver.find_element_by_id('passwd').clear()
        self.driver.find_element_by_id('account_captcha').clear()
        self.driver.find_element_by_id('uname').send_keys(str(username))
        self.driver.find_element_by_id('passwd').send_keys(str(password))
        self.driver.find_element_by_id('account_captcha').send_keys(str(valid))
        self.driver.find_element_by_css_selector('div.submit_btn').click()
        time.sleep(2)
        try:
            message = self.driver.find_element_by_css_selector('div.msg_box').text
            if message == '':
                self.save_cookies()
                print('写入cookie成功')
                return True, message
            else:
                valid_imageBase64 = self.driver.find_element_by_tag_name('img').screenshot_as_base64
                with open('./veri.png', 'wb') as img:
                    img.write(decodebytes(bytes(valid_imageBase64, encoding='utf-8')))
                return False, message
        except:
            self.save_cookies()
            return True, ''

    # 获取当前页的股票信息
    def _getCurrentPageStocksInfo(self):
        self.stockList['name_and_code'] = []
        self.stockList['is_single'] = False
        # 先设置当前页面为70页
        el = self.driver.find_element_by_id('showPerpage').find_element_by_tag_name('select')
        Select(el).select_by_value('70')
        time.sleep(1)

        # 获取所有股票数
        all_stocks = self.driver.find_element_by_id('boxTitle').find_element_by_css_selector('span.num').text
        # 返回当前页所有股票的信息
        try:
            for el in self.driver.find_element_by_css_selector('div.static_con').find_elements_by_tag_name('tr'):
                temp = el.find_elements_by_tag_name('td')
                name = temp[3].text
                code = temp[2].text
                self.stockList['name_and_code'].append([name, code])
        except:
            time.sleep(1)
            for el in self.driver.find_element_by_css_selector('div.static_con').find_elements_by_tag_name('tr'):
                temp = el.find_elements_by_tag_name('td')
                name = temp[3].text
                code = temp[2].text
                self.stockList['name_and_code'].append([name, code])
        # 获取可选择的页

        for el in self.driver.find_element_by_id('pageBar').find_elements_by_tag_name('a'):
            if el.text != '...':
                self.stockList['optional_page'].append(el.text)
        # 获取当前页
        try:
            current_page = self.driver.find_element_by_id('pageBar').find_element_by_css_selector('a.current').text
        except:
            current_page = '1'

        self.stockList['all_stocks'] = all_stocks
        self.stockList['current_page'] = current_page
        self.stockList['all_pages'] = str(int(all_stocks) // 70 + 1)

    def allStockList(self):
        # stockList Info
        self.stockList = {
            'is_single': None,
            'name_and_code': [],
            'current_page': '0',
            'optional_page': [],
            'all_stocks': '0',
            'all_pages': '0',
        }
        if self._judgeIsSingleStock():
            try:
                self.stockList['is_single'] = True
                # 返回当前股票的信息
                name = self.driver.find_element_by_css_selector('.pickName').text
                code = self.driver.find_element_by_css_selector('.pickCode').text
                self.stockList['name_and_code'].append([name, code])
                self.stockList['all_stocks'] = 1
                return self.stockList
            except:
                self.driver.close()
                return None
        else:
            try:
                self._getCurrentPageStocksInfo()
                return self.stockList
            except:
                self.driver.close()
                return None

    def getStocksListByPage(self, page):
        if type(page) == int:
            page = str(page)
        # print('page = ' + page)
        if not self._judgeIsSingleStock():
            if page == self.stockList['current_page'] or page not in self.stockList['optional_page'] \
                    or (page == '上一页' and self.stockList['current_page'] == '1') \
                    or (page == '下一页' and self.stockList['current_page'] == self.stockList['all_pages']):
                # print('error, return self.stockList, 没有翻页')
                return self.stockList
            else:
                # 查找需要翻转的页面元素
                click_el = None
                for el in self.driver.find_element_by_id('pageBar').find_elements_by_tag_name('a'):
                    if el.text == page:
                        # print('el.text = ')
                        # print(el.text)
                        # print('page = ')
                        # print(page)
                        click_el = el
                        break

                # 触发获取某一页面点击事件
                ActionChains(self.driver).move_to_element(click_el).click(click_el).perform()
                time.sleep(5)
                if not self._judgeIsSingleStock():
                    self._getCurrentPageStocksInfo()
                    return self.stockList


    def getRetrivelStockInfo(self, selectName):
        if self._judgeIsSingleStock():
            yield self.getSingleInfo(self.stockList['name_and_code'][0][0])
        else:
            yield from self.getMultipleInfo(selectName)

    def getSingleInfo(self, kwords):
        # 单页信息中只包含文字信息，不包含图
        """
                主要文字信息包括：
                    股票简称  stock_abbreviation
                    股票代码    stock_code
                    每股价格   pick_price
                    涨跌幅     price_limit
                    涨跌停原因    reasons_limit
                    总股本 （股） general_capital
                    流通比例（%） circulation_proportion
                    总市值（亿元） total_market_value
                    流通市值（亿元） market_value
                    市盈率（倍）  pe_ratio
                    市净率（倍）  net_ratio
                    报告期       report_period
                    净利润(元)   net_margin
                    净利润同比(%) net_profit
                    每股收益(元)   earnings_per_share

                    主营产品名称  main_product_name
                    所属概念    affiliation_concept
                """
        options = webdriver.ChromeOptions()
        # 添加无界面参数
        options.add_argument('--headless')
        if self.showWindow:
            driver = Chrome(self.driverPath)
        else:
            driver = Chrome(self.driverPath, options=options)
        driver.get(self.baseUrl + kwords)
        time.sleep(2)
        self.textInfo = {
            'stock_code': '',
            'stock_abbreviation': '',
            'main_product_name': [],
            'affiliation_concept': [],
            'pick_price': '',
            'price_limit': '',
            # 跌涨幅事件  多重列表： 事件时间 事件名称 事件内容
            'reasons_limit': [],
            'financial_all': {
                'report_period': '',
                'net_margin': '',
                'net_profit': '',
                'earnings_per_share': ''
            },

            'general_capital': '',
            'circulation_proportion': '',
            'total_market_value': '',
            'market_value': '',
            'pe_ratio': '',
            'net_ratio': ''
        }

        # ---------------------- 基本概况中的信息 ------------------------------------
        basic_overview = driver.find_element_by_id("dp_block_0")
        # 先点击 ·展开·
        for el in basic_overview.find_elements_by_link_text('更多'):
            ActionChains(driver).click(el).perform()
        # 股票代码
        stock_code = basic_overview.find_element_by_class_name("item").find_element_by_class_name("em").text
        self.textInfo['stock_code'] = stock_code
        # 股票简称
        stock_abbreviation = basic_overview.find_element_by_css_selector(
            ".em.alignCenter.graph").find_element_by_tag_name("a").text
        self.textInfo['stock_abbreviation'] = stock_abbreviation
        # 主营产品名称
        main_product_name = []
        for el in basic_overview.find_element_by_css_selector(".em.split").find_elements_by_tag_name("span"):
            text = el.find_element_by_tag_name("a").text
            main_product_name.append(text)
        self.textInfo['main_product_name'] = main_product_name
        # 所属概念
        try:
            affiliation_concept = []
            for el in basic_overview.find_element_by_css_selector(".em.alignCenter.split").find_elements_by_tag_name(
                    "span"):
                text = el.find_element_by_tag_name("a").text
                affiliation_concept.append(text)
            self.textInfo['affiliation_concept'] = affiliation_concept
        except:
            print("无所属概念")

        # --------------------------------涨跌幅 与 近期重要事件---------------------------------------------
        self.textInfo['pick_price'] = driver.find_element_by_css_selector('.pick_price').text
        self.textInfo['price_limit'] = driver.find_element_by_css_selector('.pick_zd_zdf').text
        try:
            important_events = driver.find_element_by_id("dp_block_1")
            for el in important_events.find_element_by_tag_name('tbody').find_elements_by_tag_name('tr'):
                if el.find_elements_by_tag_name('td')[3].text.strip() in ['涨停', '跌停']:
                    date = el.find_elements_by_tag_name('td')[2].text
                    event = el.find_elements_by_tag_name('td')[3].text
                    content = el.find_elements_by_tag_name('td')[4].find_element_by_tag_name('span').get_attribute(
                        'fullstr')
                    self.textInfo['reasons_limit'].append([date, event, content])
        except:
            print('无近期重要事件')

        # ------------------------------财务指标中的信息---------------------------------------------

        financial_index = driver.find_element_by_id("dp_block_2")
        financial_all = []
        for el in financial_index.find_element_by_tag_name("tbody").find_elements_by_tag_name("tr"):
            #  报告期       report_period
            #  净利润(元)   net_margin
            #  净利润同比(%) net_profit
            #  每股收益(元)   earnings_per_share
            report_period = el.find_elements_by_css_selector(".item")[0].find_element_by_tag_name("div").text
            net_margin = el.find_elements_by_css_selector(".item")[1].find_element_by_tag_name("div").text
            net_profit = el.find_elements_by_css_selector(".item")[2].find_element_by_tag_name("div").text
            earnings_per_share = el.find_elements_by_css_selector(".item")[5].find_element_by_tag_name("div").text
            financial_all.append([report_period, net_margin, net_profit, earnings_per_share])
        print(financial_all)
        self.textInfo['financial_all'] = financial_all

        # ------------------------------常用指标中的信息---------------------------------------------
        common_index = driver.find_element_by_id("dp_block_3")
        el = common_index.find_element_by_tag_name("tbody").find_element_by_tag_name("tr").find_elements_by_tag_name(
            "td")
        # 总股本 （股）
        general_capital = el[0].find_element_by_tag_name("div").text
        self.textInfo['general_capital'] = general_capital
        # 流通比例（ % ）
        circulation_proportion = el[3].find_element_by_tag_name("div").text
        self.textInfo['circulation_proportion'] = circulation_proportion
        # 总市值（亿元）
        total_market_value = el[1].find_element_by_tag_name("div").text
        self.textInfo['total_market_value'] = total_market_value
        # 流通市值（亿元）
        market_value = el[2].find_element_by_tag_name("div").text
        self.textInfo['market_value'] = market_value
        # 市盈率（倍）
        pe_ratio = el[4].find_element_by_tag_name("div").text
        self.textInfo['pe_ratio'] = pe_ratio
        # 市净率（倍）
        net_ratio = el[6].find_element_by_tag_name("div").text
        self.textInfo['net_ratio'] = net_ratio

        driver.close()
        return self.textInfo

    def getMultipleInfo(self, selectName):
        stocks = self.driver.find_element_by_css_selector('div.static_con').find_elements_by_tag_name('tr')
        index = 0
        retri = 0
        while index < len(stocks):
            try:
                # 多页中先获取图然后再获取相应的图片
                images = []
                if selectName[index]:
                    ActionChains(self.driver).move_to_element(
                        self.driver.find_element_by_link_text(self.stockList['name_and_code'][index][0])).perform()
                    time.sleep(0.5)
                    img_el = self.driver.find_element_by_css_selector("div.quotation-wrap")
                    for index1, el in enumerate(img_el.find_elements_by_css_selector("li.quotation-tab-item")):
                        if index1 < 4:
                            # el.click()
                            # self.driver.execute_script("arguments[0].click();", el)
                            ActionChains(self.driver).move_to_element(
                                self.driver.find_element_by_link_text(self.stockList['name_and_code'][index][0])).perform()
                            time.sleep(1)
                            ActionChains(self.driver).move_to_element(el).click(el).perform()
                            time.sleep(0.5)
                            im_base64 = img_el.screenshot_as_base64
                            images.append(im_base64)

                    # 获取文字信息
                    textInfo = self.getSingleInfo(self.stockList['name_and_code'][index][0])
                    yield textInfo, images
                    index += 1
                else:
                    index += 1
            except:
                if retri > 5:
                    print('{0}获取当前股票内容失败！！！获取下一股票'.format(retri+1))
                    index += 1
                    retri = 0
                else:
                    print('{0}获取当前股票内容失败！！！重新获取'.format(retri+1))
                    retri += 1


if __name__ == "__main__":
    grab = GrabStockInfo(keywords=["涨幅"], driverPath=r'./chromedriver.exe', showWindow=True)
    # 获取当前页面的列表
    list = grab.allStockList()

    # 通过传入选中的值
    for data in grab.getRetrivelStockInfo([True, True]):
        if len(data) == 2:
            print(data)
        if len(data) == 14:
            print(data)

    list1 = grab.getStocksListByPage('2')

    for data in grab.getRetrivelStockInfo([True, True, False, True]):
        if len(data) == 2:
            print(data)
        if len(data) == 14:
            print(data)

