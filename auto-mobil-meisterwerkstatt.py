# -*- coding: utf-8 -*-
"""
Created on Thu Apr 13 14:55:09 2017

@author: ankurD
"""

import string
import pandas
import time
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.action_chains import ActionChains


from bs4 import BeautifulSoup
import numpy as np
import re
#import clipboard
import glob
import os
#import openpyxl as px
import math


browser = webdriver.Chrome("C:/Users/AnkurDixit/WinPython-64bit-3.6.1.0Qt5/chromedriver.exe")
#browser.get('http://locatorcarrozzeriepeugeot.geosrl.net/')
time.sleep(9)
df = pandas.read_csv('C:/Users/AnkurDixit/Desktop/Python Scripts/Germany_stateCodes.csv')
#df1 = pandas.DataFrame(index = range(0,len(df)), columns = ['name','address','city','contact','email','website'])
browser.get('https://www.automobilmeisterwerkstatt.de/')
time.sleep(12)
for i in range(0,len(df)):
        try:
            actions = ActionChains(browser)
            elem = browser.find_element_by_xpath(".//*[@id='werkstattFinder']/form/div/input")
            actions.move_to_element(elem).click().perform()
            time.sleep(5)
            elem.send_keys(Keys.CONTROL + "a")
            elem.send_keys(Keys.BACKSPACE)
            elem.send_keys(str(df['Codes'][i]))
            elem.send_keys(Keys.ENTER)
            print("Value of i is :",i)
            time.sleep(15)

#elem = browser.find_elements_by_xpath(".//*[@id='sortierung']/li")
            elem1 = browser.find_elements_by_class_name("row")
            for j in range(0,len(elem1)):
                try:
                    time.sleep(10)
            #        browser.find_element_by_xpath(".//*[@id='_alltrucks_workshop_finder_WAR_alltrucksworkshopfinderportlet_wrapper']/workshop/div[1]/ui-view/ui-view/div[2]/div[1]/article/footer/a/span").click()
                    actions = ActionChains(browser)
                    temp = browser.find_element_by_xpath("(.//*[@id='main']/div/div/section/div)["+str(j+3)+"]/div[1]/div")
                    actions.move_to_element(temp).click().perform()
                    df1 = pandas.DataFrame(index = range(0,len(elem1)), columns = ['name','address1','address2','address3','fax','Zip Code','city','contact','email','website'])
#                    new_page = browser.find_element_by_xpath("(.//*[@id='spaltelinks']/p)["+str(k+1)+"]")
                    soup = BeautifulSoup(temp.get_attribute('innerHTML'))
                                #--------------------------------------------------------------------------------------------                                
                                #                     
                    try:
                        match = re.search(u"<body><b>(.*)</b><br/>",str(soup))
                        df1["name"][j] = str(match.group(1))
                    except:
                        match = re.search(u"<h4>(.*)</h4>",str(soup))
                        df1["name"][j] = str("No Value")
 #--------------------------------------------------------------------------------------------                                
#                       
                    try:
                        match = re.search(u"<nobr>(.*)</nobr>",str(soup))
                        df1["address1"][j] = str(match.group(1))
                    except:
                        match = re.search(u"<p>(.*)</p></body></html>",str(soup))
                        df1["address1"][j] = str("No Value")
#                    match = re.search(u"<h4>(.*)</h4>" ,str(soup))
#                    df1["name2"][k] = str(match.group(1)) 
                    try:
                        match = re.search(u"</nobr><br/>(.*)<br/>Tel",str(soup))
                        df1["address2"][j] = str(match.group(1))
                    except:
                        match = re.search(u"<br/>\n			(.*)<br/>",str(soup))
                        df1["address2"][j] = str("No Value")
  #--------------------------------------------------------------------------------------------                                
#                            try:    
#                                match = re.search(u"<br/>\n	          		(.*)\n	          		<br/><br/>",str(soup))
#                                df1["address3"][k] = str(match.group(1))
#                            except:
#                                match = re.search(u"additionalAdress\"(.*)\"street",str(soup))
#                                df1["address3"][k] = str("No Value")
   #--------------------------------------------------------------------------------------------                               
#                        try:
#                            match = re.search(u"<br/>\n			(.*)</p>\n<div class=\"columns clearfix",str(soup))
#                            df1["Zip Code"][k] = str(match.group(1))
#                        except:
#                            match = re.search(u"street\">(.*)</span><br/> </p> <p class=\"address_block_b",str(soup))
#                            df1["Zip Code"][k] = str("No Value")
#               #--------------------------------------------------------------------------------------------                               
                    try:    
                        match = re.search(u"Tel(.*)<br/>Fax" ,str(soup))
                        df1["contact"][j] = str(match.group(1))
                    except:
                        match = re.search(u"tel:(.*)\" rel=" ,str(soup))
                        df1["contact"][j] = str("No Value")
#              #--------------------------------------------------------------------------------------------                                
                    try:    
                        match = re.search(u"Fax:(.*)<br/><a href=\"mailto",str(soup))
                        df1["fax"][j] = str(match.group(1))
                    except:
                        match = re.search(u"Fax: (.*)</span><br/> <span class=\"mail" ,str(soup))
                        df1["fax"][j] = str("No Value")
#            #--------------------------------------------------------------------------------------------                              
                    try:
                        match= re.search(u"mailto:(.*)</a><br/><a href",str(soup))
                        df1['email'][j]=str(match.group(1))
                    except:
                        match= re.search(u"linkTo_UnCryptMailto(.*)</a></span><br/>",str(soup))
                        df1['email'][j]=str("No Value")
  #--------------------------------------------------------------------------------------------                          
                    try:
                        match= re.search(u"</a><br/><a href=(.*) target=\"_blank",str(soup))
                        df1['website'][j]=str(match.group(1))
                    except:
                        match= re.search(u"linkTo_UnCryptMailto(.*)</a></span><br/>",str(soup))
                        df1['website'][j]=str("No Value")

                    time.sleep(6)  
                    print("data frame df1 is : ",df1)
                    print("Loop is ending where the Value of i is :",i)
                except:
                    pass
#                    else:
#                        pass        
        except:
            pass
#else:
#    pass
  
if len(df1)>0:
    writer = pandas.ExcelWriter('C:/Users/AnkurDixit/Desktop/AdvancedPython/Projects_on_python/output_auto-mobil-meisterwerkstatt - ' + str(j) + str(i) +'.xlsx')
    df1.to_excel(writer)
    writer.save()
                    
#                if len(loc2) != 0:
#                    try:
#                        df1 = pandas.DataFrame(index = range(0,len(loc2)), columns = ['name','address1','address2','address3','fax','Zip Code','city','contact','email','website'])
#                        for k in range(0,len(loc2)):
#        #                    if (k%2)==0:
#                                    try:
#                                        
#                                        print("Current Value of i in try block :",i)
#                                        print("Current Value of k is :",k)
#                                        actions = ActionChains(browser)
#                                        elem_new = browser.find_element_by_xpath("(.//*[@id='spalterechts']/p)["+str(k+1)+"]/b")
#                                        actions.move_to_element(elem_new).click().perform()
#                                        time.sleep(12)
#                    #                        soup = BeautifulSoup(temp[k].get_attribute('innerHTML'))
#                                        new_page = browser.find_element_by_xpath("(.//*[@id='spalterechts']/p)["+str(k+1)+"]")
#                                        soup = BeautifulSoup(new_page.get_attribute('innerHTML'))
#                    #--------------------------------------------------------------------------------------------                                
#                    #                     
#                                        try:
#                                            match = re.search(u"<body><b>(.*)</b><br/>",str(soup))
#                                            df1["name"][k] = str(match.group(1))
#                                        except:
#                                            match = re.search(u"<h4>(.*)</h4>",str(soup))
#                                            df1["name"][k] = str("No Value")
#                     #--------------------------------------------------------------------------------------------                                
#                    #                       
#                                        try:
#                                            match = re.search(u"<nobr>(.*)</nobr>",str(soup))
#                                            df1["address1"][k] = str(match.group(1))
#                                        except:
#                                            match = re.search(u"<br/>(.*)<br/>",str(soup))
#                                            df1["address1"][k] = str(match.group(1))
#                    #                    match = re.search(u"<h4>(.*)</h4>" ,str(soup))
#                    #                    df1["name2"][k] = str(match.group(1)) 
#                                        try:
#                                            match = re.search(u"</nobr><br/>(.*)<br/>Tel",str(soup))
#                                            df1["address2"][k] = str(match.group(1))
#                                        except:
#                                            match = re.search(u"<br/>\n			(.*)<br/>",str(soup))
#                                            df1["address2"][k] = str("No Value")
#                      #--------------------------------------------------------------------------------------------                                
#                    #                            try:    
#                    #                                match = re.search(u"<br/>\n	          		(.*)\n	          		<br/><br/>",str(soup))
#                    #                                df1["address3"][k] = str(match.group(1))
#                    #                            except:
#                    #                                match = re.search(u"additionalAdress\"(.*)\"street",str(soup))
#                    #                                df1["address3"][k] = str("No Value")
#                       #--------------------------------------------------------------------------------------------                               
#                    #                        try:
#                    #                            match = re.search(u"<br/>\n			(.*)</p>\n<div class=\"columns clearfix",str(soup))
#                    #                            df1["Zip Code"][k] = str(match.group(1))
#                    #                        except:
#                    #                            match = re.search(u"street\">(.*)</span><br/> </p> <p class=\"address_block_b",str(soup))
#                    #                            df1["Zip Code"][k] = str("No Value")
#        #               #--------------------------------------------------------------------------------------------                               
#                                        try:    
#                                            match = re.search(u"Tel(.*)<br/>Fax" ,str(soup))
#                                            df1["contact"][k] = str(match.group(1))
#                                        except:
#                                            match = re.search(u"tel:(.*)\" rel=" ,str(soup))
#                                            df1["contact"][k] = str("No Value")
#        #              #--------------------------------------------------------------------------------------------                                
#                                        try:    
#                                            match = re.search(u"Fax:(.*)<br/><a href=\"mailto",str(soup))
#                                            df1["fax"][k] = str(match.group(1))
#                                        except:
#                                            match = re.search(u"Fax: (.*)</span><br/> <span class=\"mail" ,str(soup))
#                                            df1["fax"][k] = str("No Value")
#        #            #--------------------------------------------------------------------------------------------                              
#                                        try:
#                                            match= re.search(u"mailto:(.*)</a><br/><a href",str(soup))
#                                            df1['email'][k]=str(match.group(1))
#                                        except:
#                                            match= re.search(u"linkTo_UnCryptMailto(.*)</a></span><br/>",str(soup))
#                                            df1['email'][k]=str("No Value")
#                      #--------------------------------------------------------------------------------------------                          
#                                        try:
#                                            match= re.search(u"</a><br/><a href=(.*) target=\"_blank",str(soup))
#                                            df1['website'][k]=str(match.group(1))
#                                        except:
#                                            match= re.search(u"linkTo_UnCryptMailto(.*)</a></span><br/>",str(soup))
#                                            df1['website'][k]=str("No Value")
#        
#                                        time.sleep(6)  
#                                        print("data frame df2 is : ",df1)
#                                        print("Loop is ending where the Value of i in 2nd if statement is :",i)
#                                    except:
#                                        pass
#        #                    else:
#        #                        pass        
#                    except:
#                        pass
#                else:
#                    pass
#                  
#                if len(df1)>0:
#                    writer = pandas.ExcelWriter('D:/carscrape/macro/auto-mobil-meisterwerkstatt/output_auto-mobil-meisterwerkstatt - ' +str(j) + str(i) + str((i+1)*100) + '.xlsx')
#                    df1.to_excel(writer)
#                    writer.save()
##_______________________________________________________________________________________________________        
##_______________________________________________________________________________________________________
##_______________________________________________________________________________________________________
#        else:
#            loc1 = browser.find_elements_by_xpath(".//*[@id='spaltelinks']/p")
#            loc2 = browser.find_elements_by_xpath(".//*[@id='spalterechts']/p")
#            if len(loc1) != 0:
#                try:
#                    df1 = pandas.DataFrame(index = range(0,len(loc1)), columns = ['name','address1','address2','address3','fax','Zip Code','city','contact','email','website'])
#                    for k in range(0,len(loc1)):
#    #                    if (k%2)==0:
#                                try:
#    #                                df1 = pandas.DataFrame(index = range(0,len(elem*2)), columns = ['name','address1','address2','address3','fax','Zip Code','city','contact','email','website'])
#                                    print("Current Value of i in try block :",i)
#                                    print("Current Value of k is :",k)
#                                    actions = ActionChains(browser)
#                                    elem_new = browser.find_element_by_xpath("(.//*[@id='spaltelinks']/p)["+str(k+1)+"]/b")
#                                    actions.move_to_element(elem_new).click().perform()
#                                    time.sleep(12)
#                #                        soup = BeautifulSoup(temp[k].get_attribute('innerHTML'))
#                                    new_page = browser.find_element_by_xpath("(.//*[@id='spaltelinks']/p)["+str(k+1)+"]")
#                                    soup = BeautifulSoup(new_page.get_attribute('innerHTML'))
#                #--------------------------------------------------------------------------------------------                                
#                #                     
#                                    try:
#                                        match = re.search(u"<body><b>(.*)</b><br/>",str(soup))
#                                        df1["name"][k] = str(match.group(1))
#                                    except:
#                                        match = re.search(u"<h4>(.*)</h4>",str(soup))
#                                        df1["name"][k] = str("No Value")
#                 #--------------------------------------------------------------------------------------------                                
#                #                       
#                                    try:
#                                        match = re.search(u"<nobr>(.*)</nobr>",str(soup))
#                                        df1["address1"][k] = str(match.group(1))
#                                    except:
#                                        match = re.search(u"<p>(.*)</p></body></html>",str(soup))
#                                        df1["address1"][k] = str("No Value")
#                #                    match = re.search(u"<h4>(.*)</h4>" ,str(soup))
#                #                    df1["name2"][k] = str(match.group(1)) 
#                                    try:
#                                        match = re.search(u"</nobr><br/>(.*)<br/>Tel",str(soup))
#                                        df1["address2"][k] = str(match.group(1))
#                                    except:
#                                        match = re.search(u"<br/>\n			(.*)<br/>",str(soup))
#                                        df1["address2"][k] = str("No Value")
#                  #--------------------------------------------------------------------------------------------                                
#                #                            try:    
#                #                                match = re.search(u"<br/>\n	          		(.*)\n	          		<br/><br/>",str(soup))
#                #                                df1["address3"][k] = str(match.group(1))
#                #                            except:
#                #                                match = re.search(u"additionalAdress\"(.*)\"street",str(soup))
#                #                                df1["address3"][k] = str("No Value")
#                   #--------------------------------------------------------------------------------------------                               
#                #                        try:
#                #                            match = re.search(u"<br/>\n			(.*)</p>\n<div class=\"columns clearfix",str(soup))
#                #                            df1["Zip Code"][k] = str(match.group(1))
#                #                        except:
#                #                            match = re.search(u"street\">(.*)</span><br/> </p> <p class=\"address_block_b",str(soup))
#                #                            df1["Zip Code"][k] = str("No Value")
#    #               #--------------------------------------------------------------------------------------------                               
#                                    try:    
#                                        match = re.search(u"Tel(.*)<br/>Fax" ,str(soup))
#                                        df1["contact"][k] = str(match.group(1))
#                                    except:
#                                        match = re.search(u"tel:(.*)\" rel=" ,str(soup))
#                                        df1["contact"][k] = str("No Value")
#    #              #--------------------------------------------------------------------------------------------                                
#                                    try:    
#                                        match = re.search(u"Fax:(.*)<br/><a href=\"mailto",str(soup))
#                                        df1["fax"][k] = str(match.group(1))
#                                    except:
#                                        match = re.search(u"Fax: (.*)</span><br/> <span class=\"mail" ,str(soup))
#                                        df1["fax"][k] = str("No Value")
#    #            #--------------------------------------------------------------------------------------------                              
#                                    try:
#                                        match= re.search(u"mailto:(.*)</a><br/><a href",str(soup))
#                                        df1['email'][k]=str(match.group(1))
#                                    except:
#                                        match= re.search(u"linkTo_UnCryptMailto(.*)</a></span><br/>",str(soup))
#                                        df1['email'][k]=str("No Value")
#                  #--------------------------------------------------------------------------------------------                          
#                                    try:
#                                        match= re.search(u"</a><br/><a href=(.*) target=\"_blank",str(soup))
#                                        df1['website'][k]=str(match.group(1))
#                                    except:
#                                        match= re.search(u"linkTo_UnCryptMailto(.*)</a></span><br/>",str(soup))
#                                        df1['website'][k]=str("No Value")
#    
#                                    time.sleep(6)  
#                                    print("data frame df1 is : ",df1)
#                                    print("Loop is ending where the Value of i is :",i)
#                                except:
#                                    pass
#    #                    else:
#    #                        pass        
#                except:
#                    pass
#            else:
#                pass
#              
#            if len(df1)>0:
#                writer = pandas.ExcelWriter('D:/carscrape/macro/auto-mobil-meisterwerkstatt/output_auto-mobil-meisterwerkstatt - ' + str(j) + str(i)  + '.xlsx')
#                df1.to_excel(writer)
#                writer.save()
#                
#            if len(loc2) != 0:
#                try:
#                    df1 = pandas.DataFrame(index = range(0,len(loc2)), columns = ['name','address1','address2','address3','fax','Zip Code','city','contact','email','website'])
#                    for k in range(0,len(loc2)):
#    #                    if (k%2)==0:
#                                try:
#                                    
#                                    print("Current Value of i in try block :",i)
#                                    print("Current Value of k is :",k)
#                                    actions = ActionChains(browser)
#                                    elem_new = browser.find_element_by_xpath("(.//*[@id='spalterechts']/p)["+str(k+1)+"]/b")
#                                    actions.move_to_element(elem_new).click().perform()
#                                    time.sleep(12)
#                #                        soup = BeautifulSoup(temp[k].get_attribute('innerHTML'))
#                                    new_page = browser.find_element_by_xpath("(.//*[@id='spalterechts']/p)["+str(k+1)+"]")
#                                    soup = BeautifulSoup(new_page.get_attribute('innerHTML'))
#                #--------------------------------------------------------------------------------------------                                
#                #                     
#                                    try:
#                                        match = re.search(u"<body><b>(.*)</b><br/>",str(soup))
#                                        df1["name"][k] = str(match.group(1))
#                                    except:
#                                        match = re.search(u"<h4>(.*)</h4>",str(soup))
#                                        df1["name"][k] = str("No Value")
#                 #--------------------------------------------------------------------------------------------                                
#                #                       
#                                    try:
#                                            match = re.search(u"<nobr>(.*)</nobr>",str(soup))
#                                            df1["address1"][k] = str(match.group(1))
#                                    except:
#                                            match = re.search(u"<br/>(.*)<br/>",str(soup))
#                                            df1["address1"][k] = str(match.group(1))
#                #                    match = re.search(u"<h4>(.*)</h4>" ,str(soup))
#                #                    df1["name2"][k] = str(match.group(1)) 
#                                    try:
#                                        match = re.search(u"</nobr><br/>(.*)<br/>Tel",str(soup))
#                                        df1["address2"][k] = str(match.group(1))
#                                    except:
#                                        match = re.search(u"<br/>\n			(.*)<br/>",str(soup))
#                                        df1["address2"][k] = str("No Value")
#                  #--------------------------------------------------------------------------------------------                                
#                #                            try:    
#                #                                match = re.search(u"<br/>\n	          		(.*)\n	          		<br/><br/>",str(soup))
#                #                                df1["address3"][k] = str(match.group(1))
#                #                            except:
#                #                                match = re.search(u"additionalAdress\"(.*)\"street",str(soup))
#                #                                df1["address3"][k] = str("No Value")
#                   #--------------------------------------------------------------------------------------------                               
#                #                        try:
#                #                            match = re.search(u"<br/>\n			(.*)</p>\n<div class=\"columns clearfix",str(soup))
#                #                            df1["Zip Code"][k] = str(match.group(1))
#                #                        except:
#                #                            match = re.search(u"street\">(.*)</span><br/> </p> <p class=\"address_block_b",str(soup))
#                #                            df1["Zip Code"][k] = str("No Value")
#    #               #--------------------------------------------------------------------------------------------                               
#                                    try:    
#                                        match = re.search(u"Tel(.*)<br/>Fax" ,str(soup))
#                                        df1["contact"][k] = str(match.group(1))
#                                    except:
#                                        match = re.search(u"tel:(.*)\" rel=" ,str(soup))
#                                        df1["contact"][k] = str("No Value")
#    #              #--------------------------------------------------------------------------------------------                                
#                                    try:    
#                                        match = re.search(u"Fax:(.*)<br/><a href=\"mailto",str(soup))
#                                        df1["fax"][k] = str(match.group(1))
#                                    except:
#                                        match = re.search(u"Fax: (.*)</span><br/> <span class=\"mail" ,str(soup))
#                                        df1["fax"][k] = str("No Value")
#    #            #--------------------------------------------------------------------------------------------                              
#                                    try:
#                                        match= re.search(u"mailto:(.*)</a><br/><a href",str(soup))
#                                        df1['email'][k]=str(match.group(1))
#                                    except:
#                                        match= re.search(u"linkTo_UnCryptMailto(.*)</a></span><br/>",str(soup))
#                                        df1['email'][k]=str("No Value")
#                  #--------------------------------------------------------------------------------------------                          
#                                    try:
#                                        match= re.search(u"</a><br/><a href=(.*) target=\"_blank",str(soup))
#                                        df1['website'][k]=str(match.group(1))
#                                    except:
#                                        match= re.search(u"linkTo_UnCryptMailto(.*)</a></span><br/>",str(soup))
#                                        df1['website'][k]=str("No Value")
#    
#                                    time.sleep(6)  
#                                    print("data frame df2 is : ",df1)
#                                    print("Loop is ending where the Value of i in 2nd if statement is :",i)
#                                except:
#                                    pass
#    #                    else:
#    #                        pass        
#                except:
#                    pass
#            else:
#                pass
#              
#            if len(df1)>0:
#                writer = pandas.ExcelWriter('D:/carscrape/macro/auto-mobil-meisterwerkstatt/output_auto-mobil-meisterwerkstatt - ' +str(j) + str(i) + str((i+1)*100) + '.xlsx')
#                df1.to_excel(writer)
#                writer.save()        
#    except:
#        pass        
#            