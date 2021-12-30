
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from datetime import date, timedelta, datetime
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import collections



#main program
#get services
s = Service('C:/Users/z0033n5z/Downloads/chromedriver/chromedriver.exe')
driver = webdriver.Chrome(service=s)

#open the web page
driver.get('https://intranet.siemens.co.id/ptsi/wbit/ts2c10-viewedittimesheet.asp')

login_button = '/html/body/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/table/tbody/tr[2]/td[2]/table/tbody/tr/td/table/tbody/tr[9]/td/form/table/tbody/tr/td/table/tbody/tr[2]/td/div/input'
login = driver.find_element(By.XPATH,login_button)
login.click()

df= pd.read_excel('TimeSheet_RioRS.xlsx', sheet_name='Time Sheet')
mytimesheet = df.values.tolist()

d = collections.defaultdict(list)
for sub in mytimesheet:
    d[sub[0]].append(sub)

date_index = 0
start_time_index = 1
hours_spent_index = 2
project_index = 3
work_code_index = 4
description_index = 5
location_index = 6

check_xpath = '//*[@id="f111"]/table/tbody/tr/td/table[1]/tbody/tr[1]/td[4]/span[1]/select'
WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.XPATH, check_xpath)))

root_path = '//*[@id="f111"]/table/tbody/tr'
tbody_xpath = '//*[@id="f111"]/table/tbody/tr/td/table[2]/tbody'

#now loop through each rows
for rowss in d.keys():
    WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.XPATH, check_xpath)))
    root = driver.find_element(By.XPATH, root_path)
    # find the select
    month_dropdown = root.find_element(By.XPATH, ".//select[@name='m']")
    year_dropdown = root.find_element(By.XPATH, ".//select[@name='y']")
    refresh_button = root.find_element(By.XPATH, ".//input[@value='Refresh']")
    currently_selected_month = Select(month_dropdown).first_selected_option.text
    currently_selected_year = Select(year_dropdown).first_selected_option.text
    tbody = driver.find_element(By.XPATH, tbody_xpath)
    if str(d[rowss][0][hours_spent_index]) != 'nan':
        # select year and month
        if (currently_selected_year != str(d[rowss][0][date_index].year)) or \
                (currently_selected_month != d[rowss][0][date_index].strftime("%B")):
            try:
                Select(year_dropdown).select_by_value(str(d[rowss][0][date_index].year))
                currently_selected_year = str(d[rowss][0][date_index].year)
                Select(month_dropdown).select_by_visible_text(d[rowss][0][date_index].strftime("%B"))
                currently_selected_month = d[rowss][0][date_index].strftime("%B")
                refresh_button.click()
                # refind tbody
                tbody = driver.find_element(By.XPATH, tbody_xpath)
            except:
                print("something wrong")
                continue
        # now find table with date
        try:
            date_td = tbody.find_element(By.XPATH, "//a[text()='" + str(d[rowss][0][date_index].day) + "']")
            date_td.find_element(By.XPATH, "..").find_element(By.XPATH, "./following-sibling::td") \
                .find_element(By.XPATH, ".//a").click()
            # now we switched windows
            save_button_xpath = '//*[@id="f121"]/table[1]/tbody/tr[3]/td[3]/input[1]'
            return_button_xpath = '//*[@id="f121"]/table[1]/tbody/tr[3]/td[3]/input[2]'
            dtm_start_xpath = '//*[@id="f121"]/table[2]/tbody/tr[2]/td[2]/input'
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, dtm_start_xpath)))
            i = 1
            tab_body = driver.find_element(By.XPATH,'//*[@id="f121"]/table[2]/tbody')
            for rows in d[rowss]:
                print(rows)
                #now fill each row
                #find start date
                start_date = tab_body.find_element(By.XPATH,'.//td[text()="' + str(i) + '."]').find_element(By.XPATH,'..')\
                             .find_element(By.XPATH, ".//input[@name='start" + str(i) + "']")
                start_date.clear()
                start_date.send_keys(rows[start_time_index].strftime("%H:%M"))
                #end date
                end_date = tab_body.find_element(By.XPATH, './/td[text()="' + str(i) + '."]').find_element(By.XPATH,'..') \
                    .find_element(By.XPATH, ".//input[@name='end" + str(i) + "']")
                end_date.clear()
                enddate = (datetime.combine(rows[date_index], rows[start_time_index]) + timedelta(minutes= rows[hours_spent_index].minute, hours = rows[hours_spent_index].hour)).time()
                end_date.send_keys(enddate.strftime("%H:%M"))
                #work code
                work_code = tab_body.find_element(By.XPATH, './/td[text()="' + str(i) + '."]').find_element(By.XPATH,'..') \
                    .find_element(By.XPATH, ".//input[@name='prodx" + str(i) + "']").find_element(By.XPATH,'./following-sibling::a')
                work_code.click()
                MainWindow = driver.window_handles[0]
                #switch to pop up window
                WorkCode_Window = driver.window_handles[1]
                driver.switch_to.window(WorkCode_Window)
                work_code_table = driver.find_element(By.XPATH,'/html/body/form/div/table/tbody/tr[3]/td/table/tbody')
                work_code_table.find_element(By.XPATH,".//input[@value='"+ rows[work_code_index] + "']").click()
                driver.find_element(By.XPATH,'/html/body/form/div/input[2]').click()
                driver.switch_to.window(MainWindow)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, dtm_start_xpath)))
                #location
                location = tab_body.find_element(By.XPATH, './/td[text()="' + str(i) + '."]').find_element(By.XPATH,'..') \
                    .find_element(By.XPATH, ".//input[@name='loc" + str(i) + "']").find_element(By.XPATH,'./following-sibling::a')
                location.click()
                Location_Window = driver.window_handles[1]
                driver.switch_to.window(Location_Window)
                location_table = driver.find_element(By.XPATH,'/html/body/form/div/table/tbody/tr[3]/td/table/tbody')
                location_table.find_element(By.XPATH,".//input[@value='"+ rows[location_index] + "']").click()
                driver.find_element(By.XPATH,'/html/body/form/div/input[2]').click()
                driver.switch_to.window(MainWindow)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, dtm_start_xpath)))
                i += 1
            #save and return
            driver.find_element(By.XPATH,save_button_xpath).click()
            obj = driver.switch_to.alert
            obj.accept()
            driver.find_element(By.XPATH,return_button_xpath).click()
        except:
            print("something wrong")
            continue








