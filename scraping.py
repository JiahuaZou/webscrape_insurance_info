#!/usr/bin/env python
# coding: utf-8

# In[1]:


print('预计运行时间30分钟，初始化中...')

import warnings
warnings.filterwarnings("ignore")
import sys
import http.client
from decimal import *
import bs4
import time
import pandas as pd
import re
import numpy as np
import urllib
from urllib.request import Request, urlopen
import selenium
from selenium import webdriver
from bs4 import BeautifulSoup
import requests
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
#try:
#    from PIL import Image
#except ImportError:
#    import Image
#import pytesseract


# In[3]:


driver = webdriver.PhantomJS("phantomjs-2.1.1-macosx/bin/phantomjs")
year = sys.argv[1]
month = sys.argv[2]
target_year = str(year)
target_month = str(month)

getcontext().prec = 8

if len(target_month) == 1:
    targetre = target_year + "年0" + target_month + "月"
else:
    targetre = target_year + "年" + target_month + "月"
targetre2 = target_year + "年" + target_month + "月"
    
chinalife = "http://www.e-chinalife.com/chinalife-ech/news/wannenggonggao/index.jsp"
pingan = 'http://life.pingan.com/cms-tmplt/interestrateList.shtml'
cpic = 'http://life.cpic.com.cn/xrsbx/jggg/wnx/jsllgg/?subMenu=1&inSub=1'
cpic_root = 'http://life.cpic.com.cn'
taikang = 'http://www1.taikanglife.com/service/bdsearch/pro_price/6c5268b7-1.shtml'
huatai = 'http://shop.ehuatai.com/ECHT/productSearch/multi/multiInterestRateSearch.jsp'
taiping = 'http://life.cntaiping.com/ggzx-jggg/index.html?productid=7'
fude = 'http://www.sino-life.com/jggg/wnxjsllgg/'
fude_root = 'http://www.sino-life.com'
zhongying = 'http://www.aviva-cofco.com.cn/website/khfw/lycx/xxcx/lljtldwjgcx/grwnxjsllcx/index.shtml'
guangda = 'http://www.sunlife-everbright.com/eportal/ui?pageId=597652'
yangguang = 'https://wecare.sinosig.com/common/new_customerservice/html/baodanfuwu/zhjzcx_index.html'
youbang_root = 'https://www.aia.com.cn'
youbang = 'http://www.aia.com.cn/zh-cn/customer-support/service-information/ul-settlement-rate/aia_ul_settlement_rate.html'
metlife = 'https://mls.metlife.com.cn/mls/website/AlmightyRateList.jsp'
metlife_root = 'https://mls.metlife.com.cn/mls/website/AlmightyRateDetail.jsp?type='
generali = 'http://www.generalichina.com/wlls/index.jhtml'
ruitai = 'http://www.oldmutual-guodian.com/public/wanneng/'
allianz = 'http://www.allianz.com.cn/interest-notice.php'
allianz_root = 'https://www.allianz.com.cn'
zhonghong = 'http://manulife-sinochem.com/notice/view_lilv.html'
renbao = 'http://www.picclife.com/interestRate/index.jhtml'
anbang = 'http://www.anbang-life.com/lsgg/index.htm'
anbang_root = 'http://www.anbang-life.com/lsgg/'

all_df = []


# In[4]:


#pytesseract.pytesseract.tesseract_cmd = r'E:\tesseract\tesseract.exe'
#print(pytesseract.image_to_string(Image.open('E:/scrape/interet_rate_may_ruitai.jpg'), lang = 'chi_sim'))


# In[5]:


try:
    print("爬取中国人寿中")
    driver.get(chinalife)
    table = driver.find_element_by_id(id_='searchDate_table_id')
    tablehtml = table.get_attribute('innerHTML')

    soup = BeautifulSoup(tablehtml)
    lst = []
    for s in soup.find_all("td"):
        lst.append(s.get_text())
    temp = {"company" : "中国人寿", "product" : lst[::4], 
            "daily rate" : [str(Decimal(i)/10000) for i in lst[1::4]], 
            "annual rate" : [str(Decimal(i)/100) for i in lst[2::4]], 
            "time of annoucement" : lst[3::4]}
    chinalife_df = pd.DataFrame(data = temp)
    #all_df.append(chinalife_df)
    print("爬取中国人寿成功")
except Exception as e:
    print(e)
    print("爬取中国人寿失败，继续爬其他公司")


# In[6]:


try:
    print("爬取平安保险中")
    driver.get(pingan)
    table2 = driver.find_element_by_xpath(xpath = '//body')
    tablehtml2 = table2.get_attribute('innerHTML')

    soup2 = BeautifulSoup(tablehtml2)
    lst2 = []
    for s in soup2.find_all("td"):
        temp = s.get_text()
        lst2.append(re.sub(r'(\t)|(\n)|(\s)', '', temp))
    temp = {"company" : "平安保险", "product" : lst2[::5], "daily rate" : lst2[2::5], "annual rate" : lst2[1::5], "time of annoucement" : lst2[3::5]}
    pingan_df = pd.DataFrame(data = temp)
    all_df.append(pingan_df)
    print("爬取平安保险成功")
except Exception as e:
    print(e)
    print("爬取平安保险失败，继续爬其他公司")


# In[7]:


try:    
    print("爬取太平洋人寿中")
    productNameLst = []
    dailyRateLst = []
    annualRateLst = []
    annoucementTimeLst = []
    i = 0
    soup3 = BeautifulSoup(urllib.request.urlopen(cpic))
    for element in soup3.find_all('a'):
        if target_year in element.get_text() and target_month in element.get_text():
            cpic_true_link = element['href']
            break
    soup3_2 = BeautifulSoup(urllib.request.urlopen(cpic_root + cpic_true_link), 'lxml')

    for s in soup3_2.find_all('td'):
        for temp in s.findChildren('p', recursive = False):
            if "保险" in temp.get_text() or "寿险" in temp.get_text():
                productNameLst.append(temp.get_text())
                annualRateLst.append(temp.find_next(text = re.compile('[0-9]{1}\.[0-9]*%$')))
                dailyRateLst.append(str(float(temp.find_next(text = re.compile('[0-9]{1}\.[0-9]*$')))/10000)[:8])
                annoucementTimeLst.append(temp.find_next(text = re.compile('.*月$')))
    temp = {"company" : "太平洋人寿", "product" : productNameLst, "daily rate" : dailyRateLst, 
            "annual rate" : annualRateLst, "time of annoucement" : annoucementTimeLst}
    cpic_df = pd.DataFrame(data = temp)
    cpic_df.loc[16:19, 'daily rate'] = cpic_df.loc[15, 'daily rate']
    cpic_df.loc[16:19, 'annual rate'] = cpic_df.loc[15, 'annual rate']
    all_df.append(cpic_df)
    print("爬取太平洋人寿成功")
except Exception as e:
    print(e)
    print("爬取太平洋人寿失败，继续爬其他公司")


# In[8]:


try:
    print("爬取泰康人寿中，预计十分钟")
    pageLinkLst = [taikang]
    for p in range(2, 6):
        temp = list(pageLinkLst[0])
        temp[-7] = str(p)
        pageLinkLst.append(''.join(temp))
    productNameLst = []
    productLinkLst = []
    for page in pageLinkLst:
        soup_temp = BeautifulSoup(urllib.request.urlopen(page))
        time.sleep(1)
        for s in soup_temp.find_all('a'):
            if "万能型" in s.get_text() and '保证利率的公告' not in s.get_text():
                productNameLst.append(s.get_text())
                productLinkLst.append(s['href'])
    textlst = []
    dailyRateLst = []
    annualRateLst = []
    annoucementTime = []
    i = 0
    for link in productLinkLst:
        soup_temp = BeautifulSoup(urllib.request.urlopen(link))
        time.sleep(abs(np.random.normal(4, 1, 1)[0]))
        for s in soup_temp.find_all('p'):
            i += 1
            if i == 41:
                temp = re.sub(r'(\r)|(\n)|(\s)|(\u3000)', '', s.get_text())
                textlst.append(temp)
                tempLst = re.findall(r'[0-9]{1}\.[0-9]*', temp)
                annualRateLst.append(str(Decimal(tempLst[0])/Decimal(100)))
                dailyRateLst.append('0')
                annoucementTime.append(temp[25:33])
            else:
                temp = re.sub(r'(\r)|(\n)|(\s)|(\u3000)', '', s.get_text())
                textlst.append(temp)
                tempLst = re.findall(r'[0-9]{1}\.[0-9]*\%', temp)
                dailyRateLst.append(tempLst[0])
                annualRateLst.append(tempLst[1])
                annoucementTime.append(temp[25:33])
            break
        print(i)
    temp = {"company" : "泰康人寿", "product" : productNameLst, "daily rate" : dailyRateLst, 
            "annual rate" : annualRateLst, "time of annoucement" : annoucementTime}
    taikang_df = pd.DataFrame(data = temp)
    taikang_df = taikang_df.loc[taikang_df['time of annoucement'] == targetre, :]
    all_df.append(taikang_df)
    print("爬取泰康人寿成功")
except Exception as e:
    print(e)
    print("爬取泰康人寿失败，继续爬其他公司")


# In[9]:


try:
    print("爬取华泰保险中")
    driver.get(huatai)
    table = driver.find_element_by_id(id_='fraInput')
    tablehtml = table.get_attribute('innerHTML')

    soup5 = BeautifulSoup(tablehtml)
    lst = []
    lst2 = []

    for s in soup5.find_all("td"):
        lst.append(re.sub(r'(\n)|(\t)', '', s.get_text()))

    for element in lst:
        if len(element) > 20 or len(element) < 5:
            None
        else:
            lst2.append(element)

    lst2 = lst2[13:]
    temp = {"company" : "华泰保险", "product" : lst2[::4], 
            "daily rate" : lst2[1::4], #[str(Decimal(i)/10000) for i in lst2[1::4]], 
            "annual rate" : lst2[2::4], #[str(Decimal(i)/100) for i in lst2[2::4]], 
            "time of annoucement" : lst2[3::4]}
    huatai_df = pd.DataFrame(data = temp)
    huatai_df
    all_df.append(huatai_df)
    print("爬取华泰保险成功")
except Exception as e:
    print(e)
    print("爬取华泰保险失败，继续爬其他公司")


# In[10]:


try:
    print("爬取太平人寿中")
    productNameLst = []
    dailyRateLst = []
    annualRateLst = []
    annoucementTimeLst = []
    i = 0
    delay = 10
    
    driver.get(taiping)
    table = driver.find_element_by_id('jgggxqzs')
    tablehtml = table.get_attribute('innerHTML')
    soup6 = BeautifulSoup(tablehtml)
    
    for s in soup6.find_all('option'):
        if "投连" in s.get_text():
            None
        else:
            num_try = 5
            while num_try != 0:
                try:
                    i += 1
                    print(s.get_text())
                    temp = list(taiping)
                    temp[-1] = s['value']
                    temp = ''.join(temp)
                    time.sleep(1)

                    #req = Request(temp, headers={'User-Agent': 'Mozilla/5.0'})
                    driver.get(temp)
                    table = WebDriverWait(driver, delay).until(EC.presence_of_element_located((By.ID, 'jgggxqzs')))

                    tablehtml = table.get_attribute('innerHTML')
                    soup_temp = BeautifulSoup(tablehtml)

                    text = soup_temp.find_all('li')
                    #print(s)
                    lst_temp = re.findall(r'[0-9]{1}\.[0-9]*%', str(text))
                    #print(lst_temp)
                    dailyRateLst.append(lst_temp[0])
                    annualRateLst.append(lst_temp[1])
                    annoucementTimeLst.append(str(text)[5:13])
                    productNameLst.append(s.get_text())
                    num_try = 0
                except Exception as e:
                    print(e)
                    num_try -= 1

    temp = {"company" : "太平人寿", "product" : productNameLst, "daily rate" : dailyRateLst, 
            "annual rate" : annualRateLst, "time of annoucement" : annoucementTimeLst}
    taiping_df = pd.DataFrame(data = temp)
    all_df.append(taiping_df)
    print("爬取太平保险成功")
except Exception as e:
    print(e)
    print("爬取太平保险失败，继续爬其他公司")


# In[11]:


#耗时较长
try:    
    print('爬取生命人寿中')
    driver.get(fude)
    productLinkLst = []
    annualRateLst = []
    dailyRateLst = []
    announcementTimeLst = []
    productNameLst = []

    pageNum = [1, 2, 3, 4]
    xpathStr = "//div[@id='new-pagination-18']//a[text()='1']"
    for page in pageNum:
        if page != 1:
            time.sleep(abs(np.random.normal(4, 1, 1)[0]))
            strTemp = list(xpathStr)
            strTemp[-3] = str(page)
            xpathStr = ''.join(strTemp)
            nextPageBut = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, xpathStr)))
            nextPageBut.click()

        titleLst = driver.find_elements_by_xpath("//div[@id='wnjs-list']//ul//li")
        productIdLst = [re.findall(r'\"(.*)\"', BeautifulSoup(title.get_attribute('innerHTML')).find('a')['onclick'])[0] for title in titleLst if "万能型" in title.text]

        for i in range(len(productIdLst)):
            driver.find_elements_by_xpath("//div[@id='wnjs-list']//ul//li//a")[i].click()
            elem = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, productIdLst[i])))
            productLink = (fude_root + BeautifulSoup(elem.get_attribute('innerHTML')).find('a')['href'])
            productLinkLst.append(productLink)
            driver.find_element_by_xpath("//div[@id='wnjs-title']//em//a").click()

    driver.implicitly_wait(30)
    for productLink in productLinkLst:
        num_try = 5
        #content = driver.find_element_by_id('content')
        while num_try != 0:
            try:
                driver.get(productLink)
                content = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'content')))
                temp = content.get_attribute('innerHTML')
                text = BeautifulSoup(temp).find_all('p')[1].get_text()
                num_try = 0
            except Exception as e:
                #print(e)
                print('retrying')
                time.sleep(abs(np.random.normal(4, 1, 1)[0]))
                num_try -= 1
        #print(BeautifulSoup(temp))
        annualRateLst.append(re.findall('[0-9]{1}\.[0-9]*%', text)[0])
        dailyRateLst.append(re.findall('0{1}\.[0-9]*', text)[0])
        productNameLst.append(''.join(list(re.findall('[富生]{1}.*的', text)[0])[:-1]))
        announcementTimeLst.append(''.join(list(re.findall('2.*月["富生"]', text)[0])[:-1]))
        time.sleep(abs(np.random.normal(4, 1, 1)[0]))
        print(productNameLst[-1])

    temp = {"company" : "富德生命", "product" : productNameLst, "daily rate" : dailyRateLst, 
            "annual rate" : annualRateLst, "time of annoucement" : announcementTimeLst}
    fude_df = pd.DataFrame(data = temp)
    driver.implicitly_wait(0)
    all_df.append(fude_df)
    print("爬取富德保险成功")
except Exception as e:
    print(e)
    print("爬取富德保险失败，继续爬其他公司")


# In[12]:


try:
    print("爬取安邦保险中")
    webpage = requests.get(anbang)
    soup16 = BeautifulSoup(webpage.content, 'html.parser')

    link = 0
    for s in soup16.find_all('a'):
        if targetre2 in s.get_text():
            link = anbang_root + s['href']

    by_month = requests.get(link, timeout = 5)
    soup16_1 = BeautifulSoup(by_month.content, 'html.parser')

    productNameLst = []
    annualRateLst = []
    dailyRateLst = []

    announcementTimeLst = re.findall(r'2.*月', soup16_1.find('title').get_text())[0]

    for s in soup16_1.find_all(text = re.compile('账户$')):
        if s == '万能账户':
            productName = s.find_previous(text = re.compile('万能型'))
        else:
            productName = s.find_previous(text = re.compile('万能型')) + s
        productNameLst.append(productName)
        dailyRate = s.find_next(text = re.compile('^0.*'))
        dailyRate = re.sub('(%)|(\n)|(\t)|(\r)', '', dailyRate)
        dailyRate = str(Decimal(dailyRate)/Decimal(100))
        dailyRateLst.append(dailyRate)
        annualRate = s.find_next(text = re.compile('[0-9]{1}\.[0-9]{1,2}%$'))
        annualRate = re.sub('(%)|(\n)|(\t)|(\r)', '', annualRate)
        annualRate = re.sub('约等于', '', annualRate)
        annualRate = str(Decimal(annualRate)/Decimal(100))
        annualRateLst.append(annualRate)

    temp = {"company" : "安邦人寿", "product" : productNameLst, "daily rate" : dailyRateLst, 
            "annual rate" : annualRateLst, "time of annoucement" : announcementTimeLst}
    anbang_df = pd.DataFrame(data = temp)
    all_df.append(anbang_df)
    print("爬取安邦保险成功")
except Exception as e:
    print(e)
    print("爬取安邦保险失败，继续爬其他公司")


# In[13]:


try:
    print("爬取光大保险中")
    driver.get(guangda)
    el = driver.find_element_by_id('ess_ctr975_Default_grdList')
    el2 = Select(driver.find_element_by_id('ddlUnDate'))
    currentPeriod = el2.options[0].get_attribute('value')
    tablehtml = el.get_attribute('innerHTML')
    soup8 = BeautifulSoup(tablehtml)

    i = 0
    productNameLst = []
    annualRateLst = []
    dailyRateLst = []
    lst = []

    for s in soup8.find_all('td'):
        lst.append(s.get_text())
    lst = lst[3:]
    

    temp = {"company" : "光大永明", "product" : lst[::3], 
            "daily rate" : [Decimal(i)/10000 for i in lst[2::3]], 
            "annual rate" : [Decimal(i)/100 for i in lst[1::3]], 
            "time of annoucement" : currentPeriod}
    guangda_df = pd.DataFrame(data = temp)
    all_df.append(guangda_df)
    print("爬取光大保险成功")
except Exception as e:
    print(e)
    print("爬取光大保险失败，继续爬其他公司")


# In[14]:


#2018年6月后没有更新
try:
    print("爬取中宏保险万能利率")
    soup10 = BeautifulSoup(urllib.request.urlopen(zhonghong))

    productNameLst = []
    annualRateLst = []
    dailyRateLst = []
    announcementTimeLst = '2018年6月'

    for s in soup10.find_all('span'):
        productNameLst.append(s.get_text())
    productNameLst = productNameLst[1:]

    for s in soup10.find_all('tr'):
        soup_temp = re.findall(r'2{1}.*月', s.get_text())
        if len(soup_temp) == 2:
            if soup_temp[0] == announcementTimeLst:
                temp = re.findall(r'[0-9]{1}\.[0-9]*%', s.get_text())
                annualRateLst.append(temp[0])
                dailyRateLst.append(temp[1])

    temp = {"company" : "中宏保险", "product" : productNameLst, "daily rate" : dailyRateLst, 
            "annual rate" : annualRateLst, "time of annoucement" : announcementTimeLst}
    zhonghong_df = pd.DataFrame(data = temp)
    all_df.append(zhonghong_df)
    print("爬取中宏保险成功")
except Exception as e:
    print(e)
    print("爬取中宏保险失败，继续爬其他公司")


# In[15]:


try:
    print("爬取中德安联万能利率中")
    soup11 = BeautifulSoup(urllib.request.urlopen(allianz))
    productNameLst = []
    productLinkLst = []
    annualRateLst = []
    dailyRateLst = []
    announcementTimeLst = []
    i = 0

    for s in soup11.find_all('td', {'class' : 'title_search'}):
        productNameLst.append(re.sub(r'(\n)|(\t)|(\s)', '', s.get_text()))

    for s in soup11.find_all('iframe' , {'name' : 'iframe'}):
        productLinkLst.append(allianz_root + s['src'])

    for productLink in productLinkLst:
        i += 1
        driver.get(productLink)
        time.sleep(1)
        table = driver.find_element_by_id('d1')
        tablehtml = table.get_attribute('innerHTML')
        soup_temp = BeautifulSoup(tablehtml)
        annualRateLst.append(soup_temp.find('td', {'id' : 'r1:0:cellFormat8'}).get_text())
        dailyRateLst.append(soup_temp.find('td', {'id' : 'r1:0:cellFormat9'}).get_text())
        announcementTimeLst.append(soup_temp.find('td', {'class' : "xpn xpq"}).get_text())
    
    productNameLst = [name for name in productNameLst if '查询' not in name]
    
    temp = {"company" : "中德安联", "product" : productNameLst, "daily rate" : dailyRateLst, 
            "annual rate" : annualRateLst, "time of annoucement" : announcementTimeLst}
    allianz_df = pd.DataFrame(data = temp)
    all_df.append(allianz_df)
    print("爬取中德安联保险成功")
except Exception as e:
    print(e)
    print("爬取中德安联保险失败，继续爬其他公司")


# In[16]:


try:
    print("爬取大都会万能利率中")
    soup12 = BeautifulSoup(urllib.request.urlopen(metlife))

    productLinkLst = []
    annualRateLst = []
    dailyRateLst = []
    announcementTimeLst = []
    productNameLst = []

    i = 0

    for s in soup12.find_all('td', {'class', 'innertext'})[:7]:
        productNameLst.append(re.findall(r'中.*险', s.get_text())[0])
        temp = re.findall(r'\([A-Z]*.*\)', s.get_text())
        productID = [re.sub(r'(\()|(\))', '', a) for a in temp]
        productLinkLst.append([metlife_root + a for a in productID])
        #print(productLinkLst[i][0])
        soup_temp = BeautifulSoup(urllib.request.urlopen(productLinkLst[i][0]))
        time.sleep(1)
        textlst = []
        for s in soup_temp.find_all('tr' , {'class' : 'TableItem'}):
            for a in s.find_all('td'):
                textlst.append(re.sub('(\n)|(\t)|(\s)', '', a.get_text()))
        dailyRateLst.append(textlst[2])
        annualRateLst.append(textlst[3])
        announcementTimeLst.append(textlst[5])
        i += 1


    temp = {"company" : "大都会人寿", "product" : productNameLst, "daily rate" : dailyRateLst, 
            "annual rate" : annualRateLst, "time of annoucement" : announcementTimeLst}
    metlife_df = pd.DataFrame(data = temp)
    all_df.append(metlife_df)
    print("爬取大都会保险成功")
except Exception as e:
    print(e)
    print("爬取大都会保险失败，继续爬其他公司")


# In[17]:


try:
    print("爬取中意保险万能利率中")
    productNameLst = []
    annualRateLst = []
    dailyRateLst = []
    announcementTimeLst = []

    soup13 = BeautifulSoup(urllib.request.urlopen(generali))
    time.sleep(1)
    form_html = soup13.find('iframe')['src']
    soup13_1 = BeautifulSoup(urllib.request.urlopen(form_html))

    for s in soup13_1.find_all(text = '结算期间'):
        announcementTimeLst.append(s.find_next('td').get_text())
    for s in soup13_1.find_all(text = '结算利率（年化利率）'):
        annualRateLst.append(s.find_next('td').get_text())
    for s in soup13_1.find_all(text = '日结算利率'):
        dailyRateLst.append(s.find_next('td').get_text())
    for s in soup13_1.find_all('font'):
        if '结算利率' in s.get_text():
            productNameLst.append(s.get_text()[:-4])

    temp = {"company" : "中意保险", "product" : productNameLst, "daily rate" : dailyRateLst, 
            "annual rate" : annualRateLst, "time of annoucement" : announcementTimeLst}
    generali_df = pd.DataFrame(data = temp)
    all_df.append(generali_df)
    print("爬取中意保险成功")
except Exception as e:
    print(e)
    print("爬取中意保险失败，继续爬其他公司")


# In[18]:


try:
    print("爬取人保保险中")
    soup14 = BeautifulSoup(urllib.request.urlopen(renbao))
    productNameLst = []
    annualRateLst = []
    dailyRateLst = []
    announcementTimeLst = []

    for s in soup14.find_all('td', {'class' : 'table_second_1'}):
        productNameLst.append(s.get_text())
    for s in soup14.find_all('td', {'class' : 'table_second_2'}):
        announcementTimeLst.append(s.get_text())
    for s in soup14.find_all('td', {'class' : 'table_second_4'}):
        annualRateLst.append(s.get_text())
    for s in soup14.find_all('td', {'class' : 'table_second_5'}):
        dailyRateLst.append(Decimal(re.sub('(\t)|(\n)|(‰)', '', s.get_text()).replace(' ', ''))/1000)

    temp = {"company" : "人保寿险", "product" : productNameLst, "daily rate" : dailyRateLst, 
            "annual rate" : annualRateLst, "time of annoucement" : announcementTimeLst}
    renbao_df = pd.DataFrame(data = temp)
    all_df.append(renbao_df)
    print("爬取人保保险成功")
except Exception as e:
    print(e)
    print("爬取人保保险失败，继续爬其他公司")


# In[19]:


try:
    print("爬取阳光保险中")
    productNameLst = []
    dailyRateLst = []
    annualRateLst = []
    annualRate = 0
    dailyRate = 0
    productLst = []

    driver.get('https://wecare.sinosig.com/life/cus_zhanghujiazhichaxun_wnx.jsp')
    time.sleep(10)
    driver.find_element_by_id('LminSuAccrate')

    table = driver.find_element_by_id('LminSuAccrate')
    tablehtml = table.get_attribute('innerHTML')

    soup_temp = BeautifulSoup(tablehtml)
    soup_temp

    temp = soup_temp.find_all('td')
    announcementTimeLst = re.findall(r'2.*月', soup_temp.find('li').get_text())[0]

    def iterate_td(soup):
        global annualRate
        global dailyRate
        productName = 0
        if len(soup) == 0:
            return(productLst)
        else:
            if soup[0]['width'] == "20%":
                annualRate = soup[0].get_text()
            elif soup[0]['width'] == "25%":
                dailyRate = soup[0].get_text()
            else:
                productDetail = [soup[0].get_text(), annualRate, dailyRate]
                productLst.append(productDetail)
            return(iterate_td(soup[1:]))

    iterate_td(temp)
    for item in productLst:
        productNameLst.append(item[0])
        annualRateLst.append(item[1])
        dailyRateLst.append(item[2])

    temp = {"company" : "阳光人寿", "product" : productNameLst, "daily rate" : dailyRateLst, 
            "annual rate" : annualRateLst, "time of annoucement" : announcementTimeLst}
    yangguang_df = pd.DataFrame(data = temp)
    yangguang_df[['daily rate', 'annual rate', 'time of annoucement']] = yangguang_df[['daily rate', 'annual rate', 'time of annoucement']].shift(-1)
    yangguang_df.loc[65, ['daily rate', 'annual rate', 'time of annoucement']] = yangguang_df.loc[64, ['daily rate', 'annual rate', 'time of annoucement']]
    all_df.append(yangguang_df)
    print("爬取阳光保险成功")
except Exception as e:
    print(e)
    print("爬取阳光保险失败，继续爬其他公司")


# In[20]:


#有时出错，原因未知
try:
    http.client.HTTPConnection._http_vsn = 10
    http.client.HTTPConnection._http_vsn_str = 'HTTP/1.0'
    print("爬取友邦保险万能利率中")
    req = Request(youbang, headers={'User-Agent': 'Mozilla/5.0'})

    web_byte = urlopen(req).read()

    webpage = web_byte.decode('utf-8')

    productLinkLst = []
    productNameLst = []
    annualRateLst = []
    dailyRateLst = []
    announcementTimeLst = []
    
    soup9 = BeautifulSoup(webpage)
    for s in soup9.find_all('td'):
        for a in s.find_all('a'):
            productLinkLst.append(youbang_root + a['href'])
    
    driver.implicitly_wait(30)
    for i in range(len(productLinkLst)):  
        num_try = 5
        while num_try != 0:
            try:
                driver.get(productLinkLst[i])
                time.sleep(1)
                table = driver.find_element_by_tag_name('tbody')
                tablehtml = table.get_attribute('innerHTML')
                soup9_3 = BeautifulSoup(tablehtml) 
                el = soup9_3.find_all('td')
                announcementTimeLst.append(el[0].get_text())
                annualRateLst.append(el[1].get_text())
                dailyRateLst.append(el[3].get_text())
                productNameLst = soup9.find_all(string = re.compile('(万能型)'))
                print(productNameLst[i])
                num_try  = 0
            except Exception as e:
                print("retrying")
                num_try -= 1

    temp = {"company" : "友邦保险", "product" : productNameLst, "daily rate" : dailyRateLst, 
            "annual rate" : annualRateLst, "time of annoucement" : announcementTimeLst}
    youbang_df = pd.DataFrame(data = temp)
    all_df.append(youbang_df)
    print("爬取友邦保险成功")
    driver.implicitly_wait(0)
except Exception as e:
    print(e)
    print("爬取友邦保险失败，继续爬其他公司")


# In[21]:


try:
    print("爬取中英人寿万能利率中")
    driver.get(zhongying)

    productNameLst = []
    annualRateLst = []
    dailyRateLst = []
    announcementTimeLst = []

    frame = driver.find_element_by_tag_name('iframe')
    driver.switch_to.frame(frame)

    container = driver.find_element_by_id('title')
    select = Select(container)

    index = len(select.options[1:])

    for name in select.options[1:]:
        productNameLst.append(name.get_attribute('text'))

    for i in range(len(select.options)):
        if i > 0:
            container = driver.find_element_by_id('title')
            driver.execute_script("arguments[0].style.display = 'block';", container)
            select = Select(container)
            select.select_by_index(i)
            searchBtn = driver.find_element_by_id('searchBtn')
            searchBtn.click()
            element = driver.find_element_by_tag_name('tbody')
            soup_temp = BeautifulSoup(element.get_attribute('innerHTML'))
            lst = [el.get_text() for el in soup_temp.find_all('td')]
            announcementTimeLst.append(lst[1])
            dailyRateLst.append(lst[3])
            annualRateLst.append(lst[4])
        print(i)

    temp = {"company" : "中英人寿", "product" : productNameLst, "daily rate" : dailyRateLst, 
            "annual rate" : annualRateLst, "time of annoucement" : announcementTimeLst}
    zhongying_df = pd.DataFrame(data = temp)
    all_df.append(zhongying_df)
    print("爬取中英人寿成功")
except Exception as e:
    print(e)
    print("爬取中英人寿失败，继续爬其他公司")


# In[28]:


name_str = "保险公司" + targetre2 + "万能利率汇报(爬虫自动生成).xlsx"
all_data = all_df[0].append(all_df[1:])
all_data.to_excel(name_str)
print("爬取完毕")

