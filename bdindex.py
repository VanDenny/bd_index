import time
from selenium import webdriver
import urllib
import pandas as pd
import os
import pyautogui as pag
from PIL import Image
from selenium.webdriver import ActionChains
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
    gd_code ={
        "广州": "95", "深圳": "94", "佛山": "196",
        "惠州": "199", "汕头": "212", "东莞": "133",
        "茂名": "203", "江门": "198", "珠海": "200",
        "湛江": "197", "肇庆": "209", "揭阳": "205",
        "中山": "207", "韶关": "201", "阳江": "202",
        "云浮": "195", "梅州": "211", "清远": "208",
        "潮州": "204", "汕尾": "213", "河源": "210"
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
        name_element.send_keys('18702088767')
        passport_element.send_keys('18702088767')
        # name_element.send_keys('18620807798')
        # passport_element.send_keys('denglifan1989')
        # name_element.send_keys('575548935@qq.com')
        # passport_element.send_keys('baidu@890529')
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



    def get_index(self, province, city, key_word, period):
        if self.cookie_valid:
            if os.path.exists('cookie.xlsx'):
                self.cookie_list = pd.read_excel('cookie.xlsx').to_dict(orient='records')
                # print(self.cookie_list)
                self.driver.delete_all_cookies()
                for cookie in self.cookie_list:
                    if cookie['name'] in ['BDUSS', 'BAIDUID']:
                        # print(cookie)
                        self.driver.add_cookie({'name': cookie['name'], 'value': cookie['value']})
            else:
                print('找不到cookie')
                Bdindex_clawer.cookie_valid = False
        else:
            self.login()
            Bdindex_clawer.cookie_valid = True
        new_key_word = urllib.parse.quote(key_word, encoding='GBK')
        url = 'http://index.baidu.com/?tpl=trend&type=0&area=%s&time=13&word=%s%s' % (self.city_code[city], new_key_word, self.province_code[province])
        print(url)
        self.driver.get(url)
        current_url = self.driver.current_url
        denny_pic = self.driver.find_elements_by_xpath('//img[@src="/static/imgs/deny.png"]')
        if current_url == url and len(denny_pic) == 0:
            self.driver.maximize_window()
            time.sleep(3)
            if str(period) in ['24h', '7', '30', '90', 'all']:
                sel = '//a[@rel="%s"]' % period
                self.driver.find_element_by_xpath(sel).click()
            elif str(period) in ['2011', '2012', '2013', '2014', '2015', '2016', '2017']:
                self.driver.find_element_by_xpath('//a[@rel="diy"]').click()
                year_selects = self.driver.find_elements_by_class_name('yearA')
                month_selects = self.driver.find_elements_by_class_name('monthA')
                for i in range(2):
                    year_selects[i].click()
                    time.sleep(1)
                    herf = './/a[@href="#%s"]' % period
                    year_ele = year_selects[i].find_element_by_xpath(herf)
                    print(year_ele)
                    year_ele.click()
                    month_selects[i].click()
                    m_herf = './/a[@href="#%s"]' % ['01', '12'][i]
                    mon_ele = month_selects[i].find_element_by_xpath(m_herf)
                    time.sleep(1)
                    mon_ele.click()
                self.driver.find_element_by_xpath('//input[@value="确定"]').click()
            else:
                input('是否输入好日期：')
            time.sleep(1)
            self.driver.find_element_by_id('trend-meanline').click()
            time.sleep(2)
            # rect = self.driver.find_element_by_xpath("//*[name()='svg']/*[name()='rect'][11]")
            # rect_location = rect.location
            # rect_size = rect.size
            # print("location: %s, size: %s" % (rect_location, rect_size))
            # ActionChains(self.driver).move_to_element(rect).click().perform()
            js = 'window.scrollTo(0,100);'
            self.driver.execute_script(js)
            time.sleep(1)
            pag.moveTo(245, 1000, 0.5)
            pag.moveTo(245, 950, 0.5)
            pag.click()
            time.sleep(2)
            content_word = self.driver.find_element_by_id("trendBarVal")
            index_location = content_word.location
            index_size = content_word.size
            print("location: %s, size: %s" % (index_location, index_size))
            time.sleep(3)
            self.driver.save_screenshot('pic/%s_%s_%s.png' % (key_word, city, period))
            img = Image.open('pic/%s_%s_%s.png' % (key_word, city, period))
            # 跨浏览器兼容
            # scroll = self.driver.execute_script("return window.scrollY;")
            # print(scroll)
            # top = index_location['y'] - scroll
            # print(top)
            rangle = (
                int(index_location['x'] + 50),
                int(index_location['y'] + 30),
                int(index_location['x'] + index_size['width'] + 50),
                int(index_location['y'] + index_size['height'] + 28)
            )
            img = img.crop(rangle)
            x_rz = index_size['width'] * 3
            y_rz = index_size['height'] * 3
            img_rz = img.resize((x_rz, y_rz), Image.ANTIALIAS)
            img_rz = img_rz.convert('RGB')
            img_rz.save('pic/%s_%s_%s.jpg' % (key_word, city, period), quality=100)
            index_dict = {}
            index_dict['keyword'] = key_word
            index_dict['city'] = city
            index_dict['period'] = period
            try:
                code = pytesseract.image_to_string(img_rz)
                if code:
                    index_dict['index'] = code
                else:
                    index_dict['index'] = ''
            except:
                index_dict['index'] = ''
            print(index_dict)
            return index_dict
        elif denny_pic:
            time.sleep(20)
            self.driver.delete_all_cookies()
            self.get_index(province, city, key_word, period)
        else:
            Bdindex_clawer.cookie_valid = False
            self.get_index(province, city, key_word, period)




    def process(self):
        city_name = [i for i in self.city_code.keys()]
        res_list = []
        for ord, key_word in enumerate(city_name[9:]):
            for city in city_name:
                time.sleep(3)
                res_list.append(self.get_index('山东', city, key_word, '2017'))
        df = pd.DataFrame(res_list)
        df.to_excel('result/shandong.xlsx')


class Pic_to_str:
    def __init__(self, folder_path):
        self.folder_path = folder_path

    def process(self):
        res_list = []
        for root, dirs, files in os.walk(self.folder_path):
            for file in files:
                print(file)
                res_dict = {}
                if 'jpg' in file:
                    file_name = os.path.splitext(file)[0].split('_')
                    # print(file_name)
                    res_dict['key_word'], res_dict['city'], res_dict['period'] = file_name
                    file_path = os.path.join(root, file)
                    img = Image.open(file_path)
                    res_dict['index'] = pytesseract.image_to_string(img)
                    print(res_dict['index'])
                res_list.append(res_dict)
        df = pd.DataFrame(res_list)
        df.to_excel(os.path.join(root, 'index.xlsx'))













if __name__ == "__main__":
    bd_clawer = Bdindex_clawer()
    bd_clawer.process()
    # pic_to_str = Pic_to_str('D:\program_lib\Baiduindex\pic\山东省')
    # pic_to_str.process()