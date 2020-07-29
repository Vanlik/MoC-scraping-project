from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import StaleElementReferenceException
from selenium.webdriver.support.ui import Select
import pandas as pd
import csv
import time
from collections import OrderedDict
from datetime import datetime
import re
import os


# try:
    #for only window 10
    # opts = webdriver.ChromeOptions()
    # opts.binary_location = "F:\Chromedriver\chromedriver"
    # driver = webdriver.Chrome(options=opts)

    #for other than window 10
my_path = "F:\Chromedriver\chromedriver"
driver = webdriver.Chrome(executable_path=my_path)

driver.get("https://www.businessregistration.moc.gov.kh/cambodia-master/relay.html?url=https%3A%2F%2Fwww.businessregistration.moc.gov.kh%2Fcambodia-master%2Fservice%2Fcreate.html%3FtargetAppCode%3Dcambodia-master%26targetRegisterAppCode%3Dcambodia-br-companies%26service%3DregisterItemSearch&target=cambodia-master")
# driver.get('https://www.businessregistration.moc.gov.kh/cambodia-master/viewInstance/update.html?id=48e104de66a7c46f1912795850b16aab06be7ebf640a909260a3f6a1b077c0aa')
input_number = driver.find_element_by_id('QueryString')
input_number.send_keys("0")
search_button = driver.find_element_by_xpath('/html/body/div/div[1]/div[5]/div/div/div[1]/div/form/div/div/div[1]/div/div[2]/div[2]/div/div/div/div[2]/div/a[3]')
#//*[@id="nodeW208"]
search_button.click()
time.sleep(2) # browser break

next_pg = True
page_cnt = 0
test_mode = False


#empty dataframe
data = pd.DataFrame()
list_firm = []
list_error= []


#select 200 firms per page
select_adap = Select(driver.find_element_by_xpath('/html/body/div/div[1]/div[5]/div/div/div[1]/div/form/div/div/div[1]/div/div[2]/div[4]/div[1]/div/select'))
select_adap.select_by_value('4')
time.sleep(3)

driver.find_element_by_link_text('209').click()
time.sleep(2)
driver.find_element_by_link_text('201').click()
time.sleep(2)
driver.find_element_by_link_text('200').click()
time.sleep(3)
while next_pg:

            #test from x page--------------
    # if page_cnt < 203:
    #     print('page' + str(page_cnt+1))
    #     pass
    # else:

    # ele_frame = driver.find_elements_by_xpath('//*[@id="nodeW212"]/div[4]')
    time.sleep(1)
    ele_frame = driver.find_element_by_xpath(
        '/html/body/div/div[1]/div[5]/div/div/div[1]/div/form/div/div/div[1]/div/div[2]/div[4]/div[4]/div')
    ele_div = ele_frame.find_elements_by_tag_name('a')


    # front page; generate xpath of each firm url

    list_firm_path =[]
    list_date_path =[]


    for i_row in range(1, len(ele_div)+1):
        print('page_' + str(199+ page_cnt + 1) + '_row_' + str(i_row))

        #get registered date
        date_path = '/html/body/div/div[1]/div[5]/div/div/div[1]/div/form/div/div/div[1]/div/div[2]/div[4]/div[4]/div/div[2]/div[' + str(i_row) + ']/div/div/div[2]/div[2]'
        list_date_path.append(date_path)

        # get xpath of each firm per page and click to subpage
        main_url_path = '/html/body/div/div[1]/div[5]/div/div/div[1]/div/form/div/div/div[1]/div/div[2]/div[4]/div[4]/div/div[2]/div[' + str(i_row)+']/div/div/div[2]/div[1]/div[1]/a'
        list_firm_path.append(main_url_path)
        # if int(page_cnt+1) == 1:
        #     main_url_path = '/html/body/div/div[1]/div[5]/div/div/div[1]/div/form/div/div/div[1]/div/div[2]/div[4]/div[4]/div/div[2]/div[' + str(i_row) + ']/div/div/div[2]/div[1]/div[1]/a'
        #     list_firm_path.append(main_url_path)
        # else:
        #     main_url_path= '/html/body/div/div[1]/div[5]/div/div/div[1]/div/form/div[2]/div/div[1]/div/div[2]/div[4]/div[4]/div/div[2]/div['+ str(i_row) + ']/div/div/div[2]/div[1]/div[1]/a'
        #     list_firm_path.append(main_url_path)
    #Details data
    nm_firm = 0
    for enum_row, (x,y) in enumerate(zip(list_date_path, list_firm_path)):
        #-------------------------
        # if page_cnt == 203 and enum_row <100:
        if page_cnt == 0 and enum_row <40:    
            pass
        # if page_cnt >4:
        #     print('finished_page_' + str(page_cnt + 1))
        #     break
        else:
        #-------------------------
            dic_firm = {}
            try:
                ele_date_type = driver.find_element_by_xpath(x)
            except:
                driver.refresh()
                time.sleep(3) #in case bad there exist bad gateway
                ele_date_type = driver.find_element_by_xpath(x)
            dic_firm['fm_type'] = ele_date_type.text.split('\n')[-1]
            dic_firm['Registration Date'] = [i.text for i in ele_date_type.find_elements_by_class_name('appMinimalValue')][0]

            # click
            # try:
            #   driver.find_element_by_xpath(y).click()
            #   time.sleep(5)
            # except:
            driver.refresh()
            element_row = driver.find_element_by_xpath(y)
            driver.execute_script("arguments[0].click();", element_row)
            time.sleep(0.5)
            # actions = ActionChains(driver)
            # actions.move_to_element(element).click().perform()
            # actions.click(element).perform()
            


            # get firm id
            ele_id = driver.find_element_by_xpath('/html/body/div/div[1]/div[5]/div/div/div[1]/div/form/div/div/div[1]/div/div/div[1]/div/div[2]')
            string_list = [i.text for i in ele_id.find_elements_by_class_name('appAttrValue')][0].split()
            dic_firm['firm_id'] = [s for s in string_list if "000" in s][0][1:-1]
            print('page_' + str(199+page_cnt + 1) + '_row_' + str(enum_row + 1) + '_'+ dic_firm['firm_id'])

            # Take details
            ele_top_frame = driver.find_element_by_xpath(
                '/html/body/div/div[1]/div[5]/div/div/div[1]/div/form/div/div/div[1]/div/div/div[5]/div/div/div/div/div/div')
            ele_value = ele_top_frame.find_elements_by_class_name('appAttrValue')
            list_labels = [label.text for label in ele_top_frame.find_elements_by_class_name('appAttrLabelBox')]
            for en, i_value in enumerate(ele_value):
                if en == 0:
                    dic_firm['Company Name (in Khmer)'] = i_value.text
                elif en == 1:
                    dic_firm['Company Name (in English)'] = i_value.text
                elif en == 3:
                    dic_firm['Company Status'] = i_value.text
                elif en == 4:
                    dic_firm['Incorporation Date'] = i_value.text
                elif en == 5:
                    if any('Re-Registration Date' in string for string in list_labels) is True:
                        dic_firm['Re-Registration Date'] = i_value.text
                    else:
                        pass

            # Number of employees
            # # empl_xpath  = ['/html/body/div/div[1]/div[5]/div/div/div[1]/div/form/div/div/div[1]/div/div/div[5]/div/div/div/div/div/div/div[7]/div/div/div/div[3]',
            #                '/html/body/div/div[1]/div[5]/div/div/div[1]/div/form/div/div/div[1]/div/div/div[5]/div/div/div/div/div/div/div[6]/div/div/div/div[2]'
            #                ]
            empl_xpath = [
'/html/body/div/div[1]/div[5]/div/div/div[1]/div/form/div/div/div[1]/div/div/div[5]/div/div/div/div/div/div/div[6]/div/div/div/div[2]/div[2]/div/div',
'/html/body/div/div[1]/div[5]/div/div/div[1]/div/form/div/div/div[1]/div/div/div[5]/div/div/div/div/div/div/div[7]/div/div/div/div[3]/div[2]/div/div',
'/html/body/div/div[1]/div[5]/div/div/div[1]/div/form/div/div/div[1]/div/div/div[5]/div/div/div/div/div/div/div[6]/div/div/div/div[3]/div[2]/div/div',
'/html/body/div/div[1]/div[5]/div/div/div[1]/div/form/div/div/div[1]/div/div/div[5]/div/div/div/div/div/div/div[8]/div/div/div/div[3]/div[2]/div/div',
'/html/body/div/div[1]/div[5]/div/div/div[1]/div/form/div/div/div[1]/div/div/div[5]/div/div/div/div/div/div/div[7]/div/div/div/div[2]/div[2]/div/div'
'/html/body/div/div[1]/div[5]/div/div/div[1]/div/form/div/div/div[1]/div/div/div[5]/div/div/div/div/div/div/div[8]/div/div/div/div[2]/div[2]/div/div',
'/html/body/div[1]/div[1]/div[5]/div/div/div[1]/div/form/div/div/div[1]/div/div/div[5]/div/div/div/div/div/div/div[8]/div/div/div/div[2]/div[2]/div',
'/html/body/div[1]/div[1]/div[5]/div/div/div[1]/div/form/div/div/div[1]/div/div/div[5]/div/div/div/div/div/div/div[7]/div/div/div/div[2]/div[2]/div'

            ]
            # print('getting_number_of_employee_data------------:')
            for i_empl_path in empl_xpath:
                try:
                    ele_empl = driver.find_element_by_xpath(i_empl_path)
                    ele_empl_value = ele_empl.find_elements_by_class_name('appAttrValue')
                    for enu_empl, i_empl in enumerate(ele_empl_value):
                            if enu_empl == 0:
                                dic_firm['Number of male employees'] = int(i_empl.text.strip() or 0) #making integer or empty str to 0
                            elif enu_empl == 1:
                                dic_firm['Number of female employees'] = int(i_empl.text.strip() or 0)
                            else:
                                pass
                except:
                    pass
                # check for error
            try:
                dic_firm['Total number of employees'] = dic_firm['Number of male employees'] + dic_firm['Number of female employees']
                # print('DONE_getting_number_of_employee_data')
            except:
                list_error.append(dic_firm['firm_id']+'_employee_data')
                print('!!!!!!!!!!!!!!!!!!!!!!Error_getting_number_of_employee_data')


            # Business activity
            # biz_xpath= [
            # '/html/body/div/div[1]/div[5]/div/div/div[1]/div/form/div/div/div[1]/div/div/div[5]/div/div/div/div/div/div/div[7]/div/div/div/div[2]/div/div/div[2]/div/div/div/div/div[1]',
            # '/html/body/div/div[1]/div[5]/div/div/div[1]/div/form/div/div/div[1]/div/div/div[5]/div/div/div/div/div/div/div[6]/div/div/div/div[1]/div/div/div[2]/div/div/div/div/div[1]'
            #     ]
            biz_xpath =[
'/html/body/div/div[1]/div[5]/div/div/div[1]/div/form/div/div/div[1]/div/div/div[5]/div/div/div/div/div/div/div[7]/div/div/div/div[2]/div/div/div[2]/div/div/div/div/div[1]',
'/html/body/div/div[1]/div[5]/div/div/div[1]/div/form/div/div/div[1]/div/div/div[5]/div/div/div/div/div/div/div[6]/div/div/div/div[1]/div/div/div[2]/div/div/div/div/div[1]',
'/html/body/div/div[1]/div[5]/div/div/div[1]/div/form/div/div/div[1]/div/div/div[5]/div/div/div/div/div/div/div[6]/div/div/div/div[2]/div/div/div[2]/div/div/div/div/div[1]',
'/html/body/div/div[1]/div[5]/div/div/div[1]/div/form/div/div/div[1]/div/div/div[5]/div/div/div/div/div/div/div[7]/div/div/div/div[1]/div/div/div[2]/div/div/div/div/div[1]',
'/html/body/div/div[1]/div[5]/div/div/div[1]/div/form/div/div/div[1]/div/div/div[5]/div/div/div/div/div/div/div[8]/div/div/div/div[2]/div/div/div[2]/div/div/div/div/div[1]',
'/html/body/div/div[1]/div[5]/div/div/div[1]/div/form/div/div/div[1]/div/div/div[5]/div/div/div/div/div/div/div[8]/div/div/div/div[1]/div/div/div[2]/div/div/div/div/div[1]',
'/html/body/div[1]/div[1]/div[5]/div/div/div[1]/div/form/div/div/div[1]/div/div/div[5]/div/div/div/div/div/div/div[8]/div/div/div/div[1]/div/div/div[2]/div/div/div/div/div[1]'
]
            # print('getting_biz_act_data------------:')

            for i_path in biz_xpath:
                try:
                    ele_act_section = driver.find_element_by_xpath(i_path)
                    biz_act = ele_act_section.find_elements_by_class_name('appAttrValue')
                    list_activity = []
                    for enum, i_biz in enumerate(biz_act):
                        if enum % 2 != 0:
                            list_activity.append(i_biz.text)
                    for enu_act, i_act in enumerate(list_activity):
                        dic_firm['Business Activity'+'_' + str(enu_act+1)] = i_act
                except:
                    pass
                # check for error
            for i_enu in range(len(list_activity)):
                try:
                    dic_firm['Business Activity'+'_' + str(i_enu+1)]
                    # print('DONE_getting_business_activities_data'+str(i_enu+1))
                except:
                    list_error.append(dic_firm['firm_id']+'_business_activities_data')
                    print('!!!!!!!!!!!!!!!!!!!!!Error_getting_business_activities_data'+str(i_enu+1))


            # Address
            ele_head = driver.find_elements_by_xpath('/html/body/div/div[1]/div[5]/div/div/div[1]/div/form/div/div/div[1]/div/div/div[5]/div/ul')
            # time.sleep(0.6)
            if len(ele_head) == 0:
                driver.refresh()
                time.sleep(3)
            # check if Parent company exists
            if 'Parent Company' in [i.text for i in ele_head][0].split('\n'):
                # #clcik on parent section
                # link_parent = driver.find_element_by_xpath('/html/body/div/div[1]/div[5]/div/div/div[1]/div/form/div/div/div[1]/div/div/div[5]/div/ul/li[2]/a')
                # driver.execute_script("arguments[0].click();", link_parent)
                # time.sleep(5)  # for driver to switch

                # #get parent company data


                # #click on addr
                # link_addr = driver.find_element_by_xpath('/html/body/div/div[1]/div[5]/div/div/div[1]/div/form/div/div/div[1]/div/div/div[5]/div/ul/li[3]/a')
                # driver.execute_script("arguments[0].click();", link_addr)
                # time.sleep(5)  # for driver to switch

                #get addr
                link_addr = driver.find_element_by_xpath('/html/body/div/div[1]/div[5]/div/div/div[1]/div/form/div/div/div[1]/div/div/div[5]/div/ul/li[3]/a')
                driver.execute_script("arguments[0].click();", link_addr)
                time.sleep(1.5)  # for driver to switch

                # link__addr_path = ['/html/body/div/div[1]/div[5]/div/div/div[1]/div/form/div/div/div[1]/div/div/div[5]/div/div/div/div/div/div/div[1]/div/div/div/div/div[2]/div/div/div[1]/div/div[2]/div/div']
                # contact_path = '/html/body/div/div[1]/div[5]/div/div/div[1]/div/form/div/div/div[1]/div/div/div[5]/div/div/div/div/div'
            else:
                link_addr = driver.find_element_by_xpath(
        '/html/body/div/div[1]/div[5]/div/div/div[1]/div/form/div/div/div[1]/div/div/div[5]/div/ul/li[2]/a')
                driver.execute_script("arguments[0].click();", link_addr)
                time.sleep(1.5)  # for driver to switch

                # #get addr 
                # link_addr_path = [
                #     '/html/body/div/div[1]/div[5]/div/div/div[1]/div/form/div/div/div[1]/div/div/div[5]/div/div/div/div/div[1]/div[2]/div[1]/div/div/div/div/div/div/div[2]/div/div/div[1]',
                #     '/html/body/div/div[1]/div[5]/div/div/div[1]/div/form/div/div/div[1]/div/div/div[5]/div/div/div/div/div[1]/div[2]/div[1]/div/div/div/div/div/div/div[2]/div/div/div']
                # # contact_path = '/html/body/div/div[1]/div[5]/div/div/div[1]/div/form/div/div/div[1]/div/div/div[5]/div/div/div/div/div[1]/div[2]/div[2]/div'
            # print('getting_addr_contact_data------------:')    
            addr_contact_path = '/html/body/div/div[1]/div[5]/div/div/div[1]/div/form/div/div/div[1]/div/div/div[5]/div/div/div/div'
            try:
                ele_addr = driver.find_element_by_xpath(addr_contact_path)
                ele_value = ele_addr.find_elements_by_class_name('appAttrValue')
            except:
                driver.refresh()
                time.sleep(3)
                ele_addr = driver.find_element_by_xpath(addr_contact_path)
                ele_value = ele_addr.find_elements_by_class_name('appAttrValue')
            try:
                for enu, i_value in enumerate(ele_value):
                    if enu == 0:
                        dic_firm['Physical Registered Office Address'] = i_value.text
                    elif enu == 1:
                        dic_firm['Start Date'] = i_value.text
                    elif enu== len(ele_value)-4:
                        dic_firm['Postal Registered Office Address'] = i_value.text
                    elif enu==len(ele_value)-2:
                        dic_firm['Contact Email'] = i_value.text
                    elif enu ==len(ele_value)-1:
                        dic_firm['Contact Telephone Number'] = i_value.text
                    else:
                        pass
            except:
                driver.refresh()
                time.sleep(5)
                for enu, i_value in enumerate(ele_value):
                    if enu == 0:
                        dic_firm['Physical Registered Office Address'] = i_value.text
                    elif enu == 1:
                        dic_firm['Start Date'] = i_value.text
                    elif enu== len(ele_value)-4:
                        dic_firm['Postal Registered Office Address'] = i_value.text
                    elif enu==len(ele_value)-2:
                        dic_firm['Contact Email'] = i_value.text
                    elif enu ==len(ele_value)-1:
                        dic_firm['Contact Telephone Number'] = i_value.text
                    else:
                        pass

            try:
                dic_firm['Physical Registered Office Address'] or dic_firm['Start Date'] or dic_firm['Postal Registered Office Address'] or dic_firm['Contact Email'] or dic_firm['Contact Telephone Number']
                # print('DONE_getting_addr_contact_data')
            except:
                list_error.append(dic_firm['firm_id']+'_addr_contact_data')
                print('!!!!!!!!!!!!!!!!!!!!!Error_getting_addr_contact_data')
            

            #Add data into list, and dataframe
            list_firm.append(dic_firm)
            df_page = pd.DataFrame.from_dict([OrderedDict(dic_firm)])
            data = pd.concat([data, df_page], axis=0, ignore_index=True, sort=False)


            # click back
            driver.refresh()
            cancel_button = driver.find_element_by_xpath(
                '/html/body/div/div[1]/div[5]/div/div/div[1]/div/form/div/div/div[1]/div/div/div[6]/div/a')
            driver.execute_script("arguments[0].click();", cancel_button)
            time.sleep(0.2)
            
            
            # save every 20 firms
            if (nm_firm+ 1) %20 == 0:
                if os.path.exists('temp_data.csv') is True:
                    os.remove('temp_data.csv')
                    data.to_csv('temp_data.csv',encoding='utf-8' )
                    print(datetime.now().strftime('%Y-%m-%d %H:%M:%S'))

            else:
                pass
            
            nm_firm += 1

            # Test
            if test_mode is True and enum_row == 2:
                print('END_testing')
                break

    # #save for every 200 firms or 20 pages recieved (page_cnt +1) %20 ==0
    # if (page_cnt+ 1) %20 == 0:
    #     if os.path.exists('temp_data.pkl') is True:
    #         os.remove('temp_data.pkl')
    #         data.to_pickle('temp_data.pkl')
    # else:
    #     pass
        
        #Next page
    try:
        # ele_next_pg = driver.find_element_by_xpath('/html/body/div/div[1]/div[5]/div/div/div[1]/div/form/div[2]/div/div[1]/div/div[2]/div[4]/div[5]/div/div/div[3]/a')
        # driver.execute_script("arguments[0].click();", ele_next_pg)
        
        driver.refresh()
        ele_next_pg = driver.find_element_by_xpath(
            #'/html/body/div/div[1]/div[5]/div/div/div[1]/div/form/div[2]/div/div[1]/div/div[2]/div[4]/div[5]/div/div/div[3]/a')
            '/html/body/div/div[1]/div[5]/div/div/div[1]/div/form/div/div/div[1]/div/div[2]/div[4]/div[5]/div/div/div[3]/a')
        ele_next_pg.click()
        time.sleep(2)
        page_cnt += 1
        
    except:
        print('The END_')
        print(datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
        next_pg = False

# driver.quit()

#save data to csv
data.to_csv('data.csv', encoding='utf-8')
# df_error = pd.DataFrame(list_error)
# df_error.to_csv('error_data.csv', encoding='utf-8')


            # Save file
# data = pd. DataFrame.from_dict(dic_firm)
# except:
#   data.to_pickle('temp_data.pkl')



