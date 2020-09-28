from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.ui import Select
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager

import os
#coding:utf-8
from selenium import webdriver
import time
import csv
import sys
import pprint

#登入
def Account_Acess ():
    '''
        selenium using chromedriver setting option and login 
    '''
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument('--headless')
    chrome_options.add_argument('--disable-gpu')
    driver = webdriver.Chrome(options=chrome_options , executable_path = ChromeDriverManager().install())
                              
    driver.get(main_url)

    username = driver.find_element_by_name('User_id').send_keys(account)
    password = driver.find_element_by_name('password').send_keys(pwd)
   
    driver.find_element_by_name('LoginBotton').click()
    print ("登入成功！")

    return driver

def quit_first (dept_id , user_id):
    '''
        quit user or recover user 
    '''

    target_url = 'target'

    driver.get (target_url + dept_id + user_id)
    
    flag = 0
    ##########################帳號離職################## 
    if driver.find_element_by_xpath('//input[@value="離職"]').is_selected():
        flag = 1
    driver.find_element_by_xpath('//input[@value="離職"]').click()
    
    time.sleep(0.5)
    print ("處理中.." , end = '')
    update = driver.find_element_by_id('update')
    driver.execute_script ("arguments[0].click ();" , update)
    try:
        WebDriverWait(driver, 5).until(EC.alert_is_present(), 'Timed out waiting for alerts to appear')
        alert = driver.switch_to.alert
        alert.accept()
    except TimeoutException:
        print ("Timeout and No Alert Appearing")
    print (".")
    try:
        WebDriverWait(driver, 5).until(EC.alert_is_present(), 'Timed out waiting for alerts to appear')
        alert = driver.switch_to.alert
        alert.accept()
    except TimeoutException:
        print ("Timeout and No Alert Appearing")
    if flag == 1:
        print (user_id + " 已恢復。")
    else:
        print (user_id + ' 離職了。')



def main (driver):
    '''
        processing new user data and upload 
    '''

    IsKeeping = 0
    id_list = []
    target_url = 'add_url'
    edit_url = 'edit_url'
    edit_ms_url = edit_url + 'C01'   

    find_user_url = 'userfile_url'

    try:
        while True:
    
            dept = input ("請問是哪個院內機構？(按^D結束程序)\n1.芥菜種會\n2.愛心育幼院\n3.少年之家\n4.高雄新住民家庭服務中心\n")
            if dept == '1':
                dept_id = 'C01'
            elif dept == '2':
                dept_id = 'C03'
            elif dept == '3':
                dept_id = 'C04'
            elif dept == '4':
                dept_id = 'C020'
            process = input ("是否沿用(press enter reply yes)? ")
            if process != 'n':
                IsKeeping = 1
                user_id = input ("請輸入沿用工號：")
                driver.get (edit_url + dept_id + user_id)
                time.sleep (1)
                ########帳號類別##########
                s1 = Select(driver.find_element_by_xpath ("/html/body/form/table/tbody/tr[1]/td/table/tbody/tr[4]/td[4]/select")).first_selected_option.text
                ########部門##############
                s2 = Select(driver.find_element_by_xpath ("/html/body/form/table/tbody/tr[1]/td/table/tbody/tr[5]/td[2]/select")).first_selected_option.text
                ########職稱##############
                s3 = Select(driver.find_element_by_xpath ("/html/body/form/table/tbody/tr[1]/td/table/tbody/tr[5]/td[4]/Select")).first_selected_option.text
                ########直屬主管##########
                s4 = Select(driver.find_element_by_xpath ("/html/body/form/table/tbody/tr[1]/td/table/tbody/tr[5]/td[6]/select")).first_selected_option.text
                ########職務代理人########
                s5 = Select(driver.find_element_by_xpath ("/html/body/form/table/tbody/tr[1]/td/table/tbody/tr[6]/td[4]/select")).first_selected_option.text
                ########主管##############
                s6 = Select(driver.find_element_by_xpath ("/html/body/form/table/tbody/tr[1]/td/table/tbody/tr[6]/td[6]/select")).first_selected_option.text
                ########處級主管##########
                s7 = Select(driver.find_element_by_xpath ("/html/body/form/table/tbody/tr[1]/td/table/tbody/tr[7]/td[6]/select")).first_selected_option.text
       
            ###輸入新帳號資訊###
            print ("請輸入新帳號資訊：")
            data_keys = ['dept_id', 'userid' , 'user_name' , 'usergroup','department' , 'job_title' , 'supervisor' , 'email' , 'agent' , 'supervisor2' , 'join_date' , 'supervisor3' , 'SSOAccount' , 'staff_no' , 'password']
            data_values = ['芥菜種會' , 'account' , '姓名' , '一般使用者' , '' , '專員' , '' , 'aaa@mail.com.tw' , '' , '' , '2020/01/01' , '' , 'aaa@mail.com.tw' , 'account' , 'pwd']
            data_keys_trans = ['服務據點' , '帳號' , '姓名' , '帳號類別' , '部門' , '職稱' , '直屬主管' , 'Email' , '職務代理人' , '主管' , '到職日期' , '處級主管']

            ########服務據點################# 
            #data_values[0] = dept_id
            ########帳號#####################
            data_values[1] = input ("1." + data_keys_trans[1] + "：")
            ########姓名#####################
            data_values[2] = input ("2." + data_keys_trans[2] + "：")
            ########Email####################
            data_values[7] = input ("3." + data_keys_trans[7] + "：") + '@mustard.org.tw'
            ########到職日期#################
            data_values[10] = input ("4." + data_keys_trans[10] + "：")

            if IsKeeping == 1:

                data_values [3] = s1 
                data_values [4] = s2
                data_values [5] = s3
                data_values [6] = s4
                data_values [8] = s5
                data_values [9] = s6
                data_values [11] = s7

            else:
                ########帳號類別#############s1
                data_values[3] = input ("1." + data_keys_trans[3] + "：(ex,項目社工、一般使用者) ")
                ########部門#################s2
                data_values[4] = input ("2." + data_keys_trans[4] + "：(ex,北區服務中心、財會部) ")
                ########職稱#################s3
                data_values[5] = input ("3." + data_keys_trans[5] + "：(ex,社工、專員) ")
                ########直屬主管#############s4
                data_values[6] = input ("4." + data_keys_trans[6] + "：")
                ########職務代理人###########s5
                data_values[8] = input ("5." + data_keys_trans[8] + "：")
                ########主管#################s6
                data_values[9] = input ("6." + data_keys_trans[9] + "：")
                ########處級主管#############s7
                data_values[11] = input ("7." + data_keys_trans[11] + "：")

       
            while True:
                data_values[12] = data_values[7]
                data_values[13] = data_values[1]
#                data_view = dict (zip (data_keys_trans , data_values))
                print ("請確認您所輸入的內容")
                for i in range (len (data_keys_trans)):
                    print ("%d. %s : %s" % (i+1 , data_keys_trans[i] , data_values [i]))
#                pp = pprint.PrettyPrinter (indent=4)
#                pp.pprint (data_view)
                ans = input ("以上資訊是否正確？ (press enter and next) ")
                if ans == 'n':
                    print ("請輸入要修正的項目編號：(1-12) ")
                    n = input ()
                    if n == 8:
                        data_values[int (n) - 1] = input ("請輸入修改 " + data_keys_trans[int (n) - 1] + " 後的資料：") + '@mustard.org.tw'
                    else:
                        data_values[int (n) - 1] = input ("請輸入修改 " + data_keys_trans[int (n) - 1] + " 後的資料：")
                else:
                    break 
        
            data = dict (zip (data_keys , data_values))
        
            ###開始處理###
            print ("開始上傳資料")
            driver.get (target_url)
            driver.find_element_by_name ('userid').clear ()
            driver.find_element_by_name ('userid').send_keys (data['userid'])
            driver.find_element_by_name ('password').send_keys (data['password'])
            driver.find_element_by_name ('staff_no').send_keys (data['staff_no'])
            driver.find_element_by_name ('user_name').send_keys (data['user_name'])
            Select(driver.find_element_by_name ('usergroup')).select_by_visible_text (data['usergroup'])
            Select(driver.find_element_by_name ('department')).select_by_visible_text (data['department'])
            Select(driver.find_element_by_name ('job_title')).select_by_visible_text (data['job_title'])
            driver.find_element_by_name ('email').send_keys (data['email'])
            Select(driver.find_element_by_name ('agent')).select_by_visible_text (data['agent'])
            driver.find_element_by_name ('join_date').send_keys (data['join_date'])
            driver.find_element_by_name ('SSOAccount').send_keys (data['SSOAccount'])


            if dept != '1':
                
                save = driver.find_element_by_id ('save')
                driver.execute_script ("arguments[0].click ();" , save)
                
                url = edit_ms_url + ''.join(data['userid'])
                print (url)
                driver.get (url)
                try:
                    WebDriverWait(driver, 3).until(EC.alert_is_present(), 'Timed out waiting for alerts to appear')
                    alert = driver.switch_to.alert
                    alert.accept()
                except TimeoutException:
                    print ("Timeout and No Alert Appearing")
    
                if dept == '2':
                    data ['dept_id'] = '愛心育幼院'
                elif dept == '3':
                    data ['dept_id'] = '少年之家'

                Select(driver.find_element_by_name ('dept_id')).select_by_visible_text (data['dept_id'])
                print (Select(driver.find_element_by_name ('dept_id')).first_selected_option.text)
                update = driver.find_element_by_id ('update')
                driver.execute_script ("arguments[0].click ();" , update)
            
                try:
                    WebDriverWait(driver, 3).until(EC.alert_is_present(), 'Timed out waiting for alerts to appear')
                    alert = driver.switch_to.alert
                    alert.accept()
                except TimeoutException:
                    print ("Timeout and No Alert Appearing")
#                driver.get (find_user_url)
#                driver.find_element_by_name ("user_id").send_keys (i)
#                query = driver.find_element_by_id ('query')
#                driver.execute_script ("arguments[0].click ();" , query)
#                print ("查詢帳號網址")
#                time.sleep (0.5)
        
#                driver.switch_to_frame(driver.find_element_by_tag_name("iframe"))
#                user_href = driver.find_element_by_link_text (i).get_attribute ("href")
#                print ("網址： " + user_href)
                try:
                    WebDriverWait(driver, 3).until(EC.alert_is_present(), 'Timed out waiting for alerts to appear')
                    alert = driver.switch_to.alert
                    alert.accept()
                except TimeoutException:
                    print ("Timeout and No Alert Appearing")
               
                print ("已修改機構！")
                driver.get (edit_url + dept_id + ''.join(data['userid']))



                Select(driver.find_element_by_name ('supervisor')).select_by_visible_text (data['supervisor'])
                Select(driver.find_element_by_name ('supervisor2')).select_by_visible_text (data['supervisor2'])
                Select(driver.find_element_by_name ('supervisor3')).select_by_visible_text (data['supervisor3'])
            
                update = driver.find_element_by_id ('update')
                driver.execute_script ("arguments[0].click ();" , update)
            
                try:
                    WebDriverWait(driver, 3).until(EC.alert_is_present(), 'Timed out waiting for alerts to appear')
                    alert = driver.switch_to.alert
                    alert.accept()
                except TimeoutException:
                    print ("Timeout and No Alert Appearing")
                try:
                    WebDriverWait(driver, 3).until(EC.alert_is_present(), 'Timed out waiting for alerts to appear')
                    alert = driver.switch_to.alert
                    alert.accept()
                except TimeoutException:
                    print ("Timeout and No Alert Appearing")

                print ("新增完成！")

                quit_first (dept_id , data['userid'])

            else:

                Select(driver.find_element_by_name ('dept_id')).select_by_visible_text (data ['dept_id'])
                Select(driver.find_element_by_name ('supervisor')).select_by_visible_text (data['supervisor'])
                Select(driver.find_element_by_name ('supervisor2')).select_by_visible_text (data['supervisor2'])
                Select(driver.find_element_by_name ('supervisor3')).select_by_visible_text (data['supervisor3'])
                save = driver.find_element_by_id ('save')
                driver.execute_script ("arguments[0].click ();" , save)
        
                try:
                    WebDriverWait(driver, 3).until(EC.alert_is_present(), 'Timed out waiting for alerts to appear')
                    alert = driver.switch_to.alert
                    alert.accept()
                except TimeoutException:
                    print ("Timeout and No Alert Appearing")
                print ("新增完成！")

                quit_first (dept_id , data['userid'])

    except EOFError:
        print ("End Function")
    
    driver.quit()

if __name__ == '__main__':
    driver = Account_Acess ()
    main (driver)
