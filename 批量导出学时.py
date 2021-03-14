from selenium import webdriver
from bs4 import BeautifulSoup
import openpyxl
import xlwt
import time

# 登陆网址
url = 'https://tx.jci.edu.cn/a/front/index'
url_zhanshi = 'https://tx.jci.edu.cn/a/front/grzxCjd'
xueshi_1 = '思想政治引领'
xueshi_2 = '体育活动强化'
xueshi_3 = '创意创新创业实训'
xueshi_4 = '社会活动模块'
xueshi_5 = '审美与人文素养提升'

password = '000000'

driver = webdriver.Chrome('chromedriver.exe')
driver.get(url)

# 建立统计表
xueshibiaoge = xlwt.Workbook(encoding='utf-8',style_compression=0)
sheet = xueshibiaoge.add_sheet('学时统计表')
sheet.write(0,0,'姓名')
sheet.write(0,1,'学号')
sheet.write(0,2,xueshi_1)
sheet.write(0,3,xueshi_2)
sheet.write(0,4,xueshi_3)
sheet.write(0,5,xueshi_4)
sheet.write(0,6,xueshi_5)
sheet.write(0,7,'总学时')
sheet.write(0,8,'备注')

savepath = 'E:/桌面/20智能制造学时统计表.xls'


for i in range(1,48):
    xuehao = "1200408001" + str(i).zfill(2)
    # 写入学号
    sheet.write(i,1,xuehao)


    driver.find_element_by_xpath('//*[@id="topmenu"]/div/div/a').click()
    time.sleep(1)
    # 输入学号
    driver.find_element_by_xpath('//*[@id="username"]').send_keys(xuehao)
    # 输入密码
    driver.find_element_by_xpath('//*[@id="password"]').send_keys(password)
    driver.find_element_by_xpath('//*[@id="loginForm"]/table/tbody/tr[3]/td/button').click()
    if(driver.find_element_by_xpath('//*[@id="topmenu"]/div/div/a[1]').text=='登录'):
        sheet.write(i,8,'修改了密码')
        print(xuehao + "修改了密码")
        i = i + 1
        driver.find_element_by_xpath('//*[@id="topmenu"]/div/div/a').click()
        time.sleep(1)
        # 输入学号
        driver.find_element_by_xpath('//*[@id="username"]').send_keys(xuehao)
        # 输入密码
        driver.find_element_by_xpath('//*[@id="password"]').send_keys(password)
        driver.find_element_by_xpath('//*[@id="loginForm"]/table/tbody/tr[3]/td/button').click()
    else:
        # 进入第二课堂学时展示页面
        driver.get(url_zhanshi)

        # 获取学时页面
        # 获取网页源码
        pageSource = driver.page_source
        pageSource_1 = BeautifulSoup(pageSource, 'lxml')

        # 查询

        name = driver.find_element_by_xpath('//*[@id="topmenu"]/div/div/a[1]').text

        sheet.write(i,0,name)

        print(name + ":" + xuehao)
        if(driver.find_elements_by_tag_name('strong')[0].text=='0.0'):
            xueshi1 = '0'
            xueshi2 = '0'
            xueshi3 = '0'
            xueshi4 = '0'
            xueshi5 = '0'
            sheet.write(i, 2, xueshi1)
            sheet.write(i, 3, xueshi2)
            sheet.write(i, 4, xueshi3)
            sheet.write(i, 5, xueshi4)
            sheet.write(i, 6, xueshi5)
            print(xueshi_1 + ":" + xueshi1)
            print(xueshi_2 + ":" + xueshi2)
            print(xueshi_3 + ":" + xueshi3)
            print(xueshi_4 + ":" + xueshi4)
            print(xueshi_5 + ":" + xueshi5)

        else:
            for j in range(len(driver.find_elements_by_class_name('text-sxdd'))):
                if(driver.find_elements_by_class_name('text-sxdd')[j].text==xueshi_1):
                    xueshi1 = driver.find_elements_by_tag_name('strong')[j].text
                    sheet.write(i, 2, xueshi1)
                    print(xueshi_1 + ":" + xueshi1)
                else:

                    if (driver.find_elements_by_class_name('text-sxdd')[j].text == xueshi_2):
                        xueshi2 = driver.find_elements_by_tag_name('strong')[j].text
                        sheet.write(i, 3, xueshi2)
                        print(xueshi_2 + ":" + xueshi2)
                    else:

                        if (driver.find_elements_by_class_name('text-sxdd')[j].text == xueshi_3):
                            xueshi3 = driver.find_elements_by_tag_name('strong')[j].text
                            sheet.write(i, 4, xueshi3)
                            print(xueshi_3 + ":" + xueshi3)
                        else:

                            if (driver.find_elements_by_class_name('text-sxdd')[j].text == xueshi_4):
                                xueshi4 = driver.find_elements_by_tag_name('strong')[j].text
                                sheet.write(i, 5, xueshi4)
                                print(xueshi_4 + ":" + xueshi4)
                            else:

                                if (driver.find_elements_by_class_name('text-sxdd')[j].text == xueshi_5):
                                    xueshi5 = driver.find_elements_by_tag_name('strong')[j].text
                                    sheet.write(i, 6, xueshi5)
                                    print(xueshi_5 + ":" + xueshi5)

        driver.find_element_by_xpath('//*[@id="topmenu"]/div/div/a[3]').click()
        time.sleep(2)
xueshibiaoge.save(savepath)

