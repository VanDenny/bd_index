import time
from selenium import webdriver
import urllib
import pandas as pd
import os
from selenium.webdriver.common.action_chains import ActionChains
from PIL import Image
import pytesseract

class Bdindex_clawer:
    passport_url = 'https://passport.baidu.com/v2/?login&tpl=mn&u=http%3A%2F%2Fwww.baidu.com%2F'
    index_url = 'http://index.baidu.com'
    province_code = {
        "山东": "#901", "贵州": "#902", "江西": "#903",
        "重庆": "#904", "内蒙古": "#905", "湖北": "#906",
        "辽宁": "#907", "湖南": "#908", "福建": "#909",
        "上海": "#910", "北京": "#911", "广西": "#912",
        "广东": "#913", "四川": "#914", "云南": "#915",
        "江苏": "#916", "浙江": "#917", "青海": "#918",
        "宁夏": "#919", "河北": "#920", "黑龙江": "#921",
        "吉林": "#922", "天津": "#923", "陕西": "#924",
        "甘肃": "#925", "新疆": "#926", "河南": "#927",
        "安徽": "#928", "山西": "#929", "海南": "#930",
        "台湾": "#931", "西藏": "#932", "香港": "#933",
        "澳门": "#934"
    }
    city_code = {
        "济南": "1", "青岛": "77", "潍坊": "80",
        "烟台": "78", "临沂": "79", "淄博": "81",
        "泰安": "353", "济宁": "352", "聊城": "83",
        "东营": "82", "威海": "88", "德州": "86",
        "滨州": "76", "莱芜": "356", "枣庄": "85",
        "菏泽": "84", "日照": "366"
    }
    cookie_valid = True
    def __init__(self):
        self.driver = webdriver.Chrome()

    def openweb(self):
        pass

    def login(self):
        self.driver.get(self.passport_url)
        self.driver.find_element_by_xpath("//p[@title='用户名登录']").click()
        name_element = self.driver.find_element_by_id("TANGRAM__PSP_3__userName")
        passport_element = self.driver.find_element_by_id("TANGRAM__PSP_3__password")
        submit_element = self.driver.find_element_by_id("TANGRAM__PSP_3__submit")
        name_element.send_keys('575548935@qq.com')
        passport_element.send_keys('baidu@890529')
        submit_element.click()
        verify_input = self.driver.find_elements_by_id("TANGRAM__PSP_3__verifyCode")
        if verify_input:
            input('是否人工通过验证：')
        self.cookie_list = self.get_cookie()
        df = pd.DataFrame(self.cookie_list).set_index('name')
        print(df)
        df.to_excel('cookie.xlsx')


    def get_cookie(self):
        cookie_list = self.driver.get_cookies()
        return cookie_list



    def get_index(self, province, city, period):
        if self.cookie_valid:
            if os.path.exists('cookie.xlsx'):
                self.cookie_list = pd.read_excel('cookie.xlsx').to_dict(orient='records')
                print(self.cookie_list)
                self.driver.delete_all_cookies()
                for cookie in self.cookie_list:
                    if cookie['name'] in ['BDUSS', 'BAIDUID']:
                        self.driver.add_cookie(cookie)
            else:
                print('找不到cookie')
                Bdindex_clawer.cookie_valid = False
        else:
            self.login()
            Bdindex_clawer.cookie_valid = True
        key_word = urllib.parse.quote(city, encoding='GBK')
        url = 'http://index.baidu.com/?tpl=trend&type=0&area=%s&time=13&word=%s%s' % (self.city_code[city], key_word, self.province_code[province])
        print(url)
        self.driver.get(url)
        current_url = self.driver.current_url
        if current_url == url:
            self.driver.maximize_window()
            time.sleep(2)
            sel = '//a[@rel="%s"]' % period
            self.driver.find_element_by_xpath(sel).click()
            self.driver.find_element_by_id('trend-meanline').click()
        else:
            Bdindex_clawer.cookie_valid = False
            self.get_index(province, city, period)









if __name__ == "__main__":
    bd_clawer = Bdindex_clawer()
    bd_clawer.get_index('山东', '威海', 'all')