#!/usr/bin/env python
# coding: utf-8

# In[10]:


import requests, bs4, time, datetime, csv
import pandas as pd
from pandas import DataFrame
import xlwings as xw

#The Web page to scrape. DELL Forum
url_1_DELL = 'https://www.dell.com/community/XPS/bd-p/XPS'
#Store the HTML of web page.
htmlfile = requests.get(url_1_DELL)
#Read the HTML via Beautifulsoup tool
objSoup = bs4.BeautifulSoup(htmlfile.text,'lxml')
#Set the location according to the tags of data that we need.
items_title = objSoup.find_all('span','lia-message-unread')
items_date = objSoup.find_all('div','lia-info-area')
items_view = objSoup.find_all('div','lia-stats-area')

ISOTIMEFORMAT = '%Y-%m-%d %H:%M:%S'
theTime = datetime.datetime.now().strftime(ISOTIMEFORMAT)

#The place used to store the data, also called container.
title_1 = []
title_date_1 = []
title_reply_1 = []
title_view_1 = []
title_href_1 = []
title_dell_account = []

#-----Scrape Start-----
print('Dell Forum Crawler  ' + 'TIME:' , theTime + '\n')
print('----- Page:1 -----')
#Start scraping from page 1, there are 11 information of page 1.
for i in range(0,11): #0-11
    
    #for split replies and views
    pos_reply = items_view[i].text.replace(" ", "").replace("\n","").find('R')
    pos_view_1 = items_view[i].text.replace(" ", "").replace("\n","").find('y')
    if pos_view_1 < 0:
        pos_view_1 = items_view[i].text.replace(" ", "").replace("\n","").find('s')
    pos_view_2 = items_view[i].text.replace(" ", "").replace("\n","").find('V')
    pos_date = items_date[i].text.replace(" ", "").replace("\n","").find('Latestposton')
    if pos_date < 0: pos_date = 0
    
    print('Title: ', items_title[i].a.text.strip() + '\n' +
          'Post_Date: ', items_date[i].text.replace(" ", "").replace("\n","")[pos_date-17:pos_date-7] + '\n' + 
          'Replies: ' , items_view[i].text.replace(" ", "").replace("\n","")[0:pos_reply] + '\n' +
          'Views: ' , items_view[i].text.replace(" ", "").replace("\n","")[pos_view_1+1:pos_view_2] + '\n' +
          'Href: dell.com', items_title[i].a['href'])
    #Put the data in the container.
    title_1.append(items_title[i].a.text.strip())
    title_date_1.append(items_date[i].text.replace(" ", "").replace("\n","")[pos_date-17:pos_date-7])
    title_reply_1.append(items_view[i].text.replace(" ", "").replace("\n","")[0:pos_reply])
    title_view_1.append(items_view[i].text.replace(" ", "").replace("\n","")[pos_view_1+1:pos_view_2])
    title_href_1.append('dell.com' + items_title[i].a['href'])
       
    url_1_DELL_Href = 'https://www.dell.com{}'.format(items_title[i].a['href'])
    #Store the HTML of web page.
    htmlfile = requests.get(url_1_DELL_Href)
    #Read the HTML via Beautifulsoup tool
    objSoup = bs4.BeautifulSoup(htmlfile.text,'lxml')
    items_dell_account = objSoup.find_all('div','lia-message-author-rank lia-component-author-rank lia-component-message-view-widget-author-rank')
    
    dell_count = 0
    for i in items_dell_account:
        if 'Moderator' in i.text:
            dell_count = dell_count+1
    
    if dell_count != 0:
        title_dell_account.append('Y')
        dell_reply = 'Y'
    else:
        title_dell_account.append('N')
        dell_reply = 'N'
    
    print('Dell Reply: ' , dell_reply)
    print('\n')
    
print('-----title:', len(title_1))
print('-----Date:', len(title_date_1))
print('-----reply:', len(title_reply_1))
print('-----View:', len(title_view_1))
print('-----href:', len(title_href_1))
print('-----dell:', len(title_dell_account))
time.sleep(3)
print('\n')


#爬蟲第2~945頁
for p in range(2,945): #2-945
    url_2_DELL = 'https://www.dell.com/community/XPS/bd-p/XPS/page/{}'.format(str(p))
    
    htmlfile = requests.get(url_2_DELL)
    objSoup = bs4.BeautifulSoup(htmlfile.text,'lxml')
    items_title = objSoup.find_all('span','lia-message-unread')
    items_date = objSoup.find_all('div','lia-info-area')
    items_view = objSoup.find_all('div','lia-stats-area')
    
    print('----- Page:' ,p , ' -----')
    for i in range(0,10): #0-10
        pos_reply = items_view[i].text.replace(" ", "").replace("\n","").find('R')
        pos_view_1 = items_view[i].text.replace(" ", "").replace("\n","").find('y')
        if pos_view_1 < 0:
            pos_view_1 = items_view[i].text.replace(" ", "").replace("\n","").find('s')
        pos_view_2 = items_view[i].text.replace(" ", "").replace("\n","").find('V')
        pos_date = items_date[i].text.replace(" ", "").replace("\n","").find('Latestposton')
        if pos_date < 0: pos_date = 0
    
        print('Title: ', items_title[i].a.text.strip() + '\n' +
              'Post_Date: ', items_date[i].text.replace(" ", "").replace("\n","")[pos_date-17:pos_date-7] + '\n' + 
              'Replies: ' , items_view[i].text.replace(" ", "").replace("\n","")[0:pos_reply] + '\n' +
              'Views: ' , items_view[i].text.replace(" ", "").replace("\n","")[pos_view_1+1:pos_view_2] + '\n' +
              'Href: dell.com', items_title[i].a['href'])
        #Put the data in the container.
        title_1.append(items_title[i].a.text.strip())
        title_date_1.append(items_date[i].text.replace(" ", "").replace("\n","")[pos_date-17:pos_date-7])
        title_reply_1.append(items_view[i].text.replace(" ", "").replace("\n","")[0:pos_reply])
        title_view_1.append(items_view[i].text.replace(" ", "").replace("\n","")[pos_view_1+1:pos_view_2])
        title_href_1.append('dell.com' + items_title[i].a['href'])
        
        url_2_DELL_Href = 'https://www.dell.com{}'.format(items_title[i].a['href'])
        #Store the HTML of web page.
        htmlfile = requests.get(url_2_DELL_Href)
        #Read the HTML via Beautifulsoup tool
        objSoup = bs4.BeautifulSoup(htmlfile.text,'lxml')
        items_dell_account = objSoup.find_all('div','lia-message-author-rank lia-component-author-rank lia-component-message-view-widget-author-rank')
    
        dell_count = 0
        for i in items_dell_account:
            if 'Moderator' in i.text:
                dell_count = dell_count+1
        
        if dell_count != 0:
            title_dell_account.append('Y')
            dell_reply = 'Y'
        else:
            title_dell_account.append('N')
            dell_reply = 'N'
            
        print('Dell Reply: ' , dell_reply)
        print('\n')
    
    print('-----title:', len(title_1))
    print('-----Date:', len(title_date_1))
    print('-----reply:', len(title_reply_1))
    print('-----View:', len(title_view_1))
    print('-----href:', len(title_href_1))
    print('-----dell:', len(title_dell_account))
    time.sleep(3)
    print('\n')

#The Web page to scrape. NotebookReview
url_1_NotebookReview = 'http://forum.notebookreview.com/forums/dell-xps-and-studio-xps.1049/'
htmlfile = requests.get(url_1_NotebookReview)
objSoup = bs4.BeautifulSoup(htmlfile.text,'lxml')

items_title = objSoup.find_all('div','titleText')
items_date = objSoup.find_all('a','faint')
items_reply = objSoup.find_all('dl','major')
items_view = objSoup.find_all('dl','minor')

ISOTIMEFORMAT = '%Y-%m-%d %H:%M:%S'
theTime_2 = datetime.datetime.now().strftime(ISOTIMEFORMAT)

print('NotebookReview Crawler  ' + 'TIME:' , theTime_2 + '\n')
print('----- Page:1 -----')
for i in range(0,20): #標準參數值0-20
    
    pos_date = items_date[i].text.find('at')
    print(pos_date)
    if pos_date < 0 :
        pos_date = 13    
    
    print('Title：', items_title[i].a.text.strip() + '\n' +
          'Post_Date: ', items_date[i].text[0:pos_date] + '\n' + 
          'Replies:' , items_reply[i].text.replace("Replies:", " ").replace(",","") + '\n' +
          'Views:' , items_view[i].text.replace("Views:", " ").replace(",","") + '\n' +
          'Href: forum.notebookreview.com/', items_title[i].a['href'])
    
    title_1.append(items_title[i].a.text.strip())
    title_date_1.append(items_date[i].text[0:pos_date])
    title_reply_1.append(items_reply[i].text.replace("Replies:", " ").replace(",",""))
    title_view_1.append(items_view[i].text.replace("Views:", " ").replace(",",""))
    title_href_1.append('forum.notebookreview.com/' + items_title[i].a['href'])
    title_dell_account.append('N')
    print('Dell Reply: N')
    print('\n')

print('-----Title:', len(title_1))
print('-----Date:', len(title_date_1))
print('-----Reply:', len(title_reply_1))
print('-----View:', len(title_view_1))
print('-----Href:', len(title_href_1))
print('-----dell:', len(title_dell_account))
time.sleep(3)
print('\n')

for p in range(2,928): #標準參數值2-928
    url_2_NotebookReview = 'http://forum.notebookreview.com/forums/dell-xps-and-studio-xps.1049/page-{}'.format(str(p))
    
    htmlfile = requests.get(url_2_NotebookReview)
    objSoup = bs4.BeautifulSoup(htmlfile.text,'lxml')
    items_title = objSoup.find_all('div','titleText')
    items_date = objSoup.find_all('a','faint')
    items_reply = objSoup.find_all('dl','major','dd')
    items_view = objSoup.find_all('dl','minor')
    
       
    print('----- Page:' ,p , ' -----')
    for i in range(0,20): #標準參數值0-20
        
        pos_date = items_date[i].text.find('at')
        print(pos_date)
        if pos_date < 0 :
            pos_date = 13  
        
        print('Title：', items_title[i].a.text.strip() + '\n' +
              'Post_Date: ', items_date[i].text[0:pos_date] + '\n' + 
              'Replies:' , items_reply[i].text.replace("Replies:", " ").replace(",","") + '\n' +
              'Views:' , items_view[i].text.replace("Views:", " ").replace(",","") + '\n' +
              'Href: forum.notebookreview.com/', items_title[i].a['href'])
    
        title_1.append(items_title[i].a.text.strip())
        title_date_1.append(items_date[i].text[0:pos_date])
        title_reply_1.append(items_reply[i].text.replace("Replies:", " ").replace(",",""))
        title_view_1.append(items_view[i].text.replace("Views:", " ").replace(",",""))
        title_href_1.append('forum.notebookreview.com/' + items_title[i].a['href'])
        title_dell_account.append('N')
        print('Dell Reply: N')
        print('\n')

    print('-----Title:', len(title_1))
    print('-----Date:', len(title_date_1))
    print('-----Reply:', len(title_reply_1))
    print('-----View:', len(title_view_1))
    print('-----Href:', len(title_href_1))
    print('-----dell:', len(title_dell_account))
    time.sleep(3)
    print('\n')

#Write to CSV
wb = xw.Book('SPT Web Crawler Tool_V2.1.xlsm')
sheet = wb.sheets['Database']

sheet.range('a2:a1048576').clear_contents()
sheet.range('b2:b1048576').clear_contents()
sheet.range('c2:c1048576').clear_contents()
sheet.range('d2:d1048576').clear_contents()
sheet.range('e2:e1048576').clear_contents()
sheet.range('f2:f1048576').clear_contents()

sheet.range('a2').options(transpose=True).value = title_1[0:len(title_1)+1]
sheet.range('b2').options(transpose=True).value = title_date_1[0:len(title_date_1)+1]
sheet.range('c2').options(transpose=True).value = title_reply_1[0:len(title_reply_1)+1]
sheet.range('d2').options(transpose=True).value = title_view_1[0:len(title_view_1)+1]
sheet.range('e2').options(transpose=True).value = title_href_1[0:len(title_href_1)+1]
sheet.range('f2').options(transpose=True).value = title_dell_account[0:len(title_dell_account)+1]

wb.save('SPT Web Crawler Tool_V2.1.xlsm')


#-----Scrape End-----
theTime_finish = datetime.datetime.now().strftime(ISOTIMEFORMAT)
print("Crawler Finished  " + 'TIME:' , theTime_finish)
