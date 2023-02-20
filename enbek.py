import ctypes
import glob
import json
import os
import pathlib
import shutil
import urllib3
import re
import win32clipboard
from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.remote.switch_to import SwitchTo
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.select import Select
from selenium.webdriver.support.ui import WebDriverWait
from datetime import datetime, timedelta
import pandas as pd
import time
import pyautogui
from pywinauto import Application, keyboard

df = pd.read_excel("файл с !с")
df1 = pd.read_excel("мапинг должности для енбека")

chrome_options = webdriver.ChromeOptions()
prefs = {"profile.default_content_setting_values.notifications" : 2}
chrome_options.add_experimental_option("prefs",prefs)
driver = webdriver.Chrome(executable_path=r'path', chrome_options=chrome_options)

driver.get('https://enbek.kz/ru/cabinet/vacint/mypou')
time.sleep(1)
inputElement1 = driver.find_element(By.NAME, "email")
inputElement1.send_keys('email')

inputElement2 = driver.find_element(By.NAME, "password")
inputElement2.send_keys('password')


inputElement2.send_keys(Keys.ENTER)
driver.maximize_window()

driver.get('https://hr.enbek.kz/')
time.sleep(1)

driver.find_element(By.CSS_SELECTOR, '[class="greenBtn mob"]').click()
time.sleep(1)


driver.find_element(By.CSS_SELECTOR, '[class="smallButton_btn__1IAPG"]').click()
time.sleep(1)

driver.get('https://hr.enbek.kz/contracts')

df_recruitment = df[(df['Unnamed: 6'] == 'Прием на работу') & ((df['Unnamed: 24'] == '(ГО) АО «Home Credit Bank»') | (df['Unnamed: 24'] == 'управление по г.Алматы'))]

b1 = []
b2 = []
b3 = []
b4 = []
b5 = []
b6 = []
b7 = []
b8 = []
b9 = []

# 75
for i in range(len(list(df_recruitment['Unnamed: 4']))):
    
# for m in range(len(list_iin)):
#     ind = list(df['ИИН']).index(list_iin[m])
#     i = ind
# #     try:
    iin = str(list(df_recruitment['Unnamed: 4'])[i])
    for n in range(3):
        if len(str(iin)) != 12:
            iin = str('0'+ str(iin))
    print(iin)   
    fullName = list(df_recruitment['Unnamed: 15'])[i]
    FIO = list(df_recruitment['Unnamed: 16'])[i]
    registerDate = list(df_recruitment['Unnamed: 20'])[i]
    dateFrom = list(df_recruitment['Unnamed: 21'])[i]
    dateTo = list(df_recruitment['Unnamed: 22'])[i]
    job2 = list(df_recruitment['Unnamed: 11'])[i]
    loc = list(df_recruitment['Unnamed: 13'])[i]
    branch = list(df_recruitment['Unnamed: 24'])[i]

    inputElement1 = driver.find_element(By.CSS_SELECTOR, '[class="MuiInputBase-input css-mnn31"]')
    inputElement1.send_keys(Keys.CONTROL + "a")
    inputElement1.send_keys(Keys.DELETE)
    inputElement1.send_keys(iin)
    time.sleep(1) 

    try:
        driver.find_element(By.CSS_SELECTOR, '[class="contraxtsSearch_searchField__button__2DHET"]').click()
    except:
        driver.find_elements(By.XPATH, "//*[text()='Назад']")[0].click()
        time.sleep(1)
        driver.find_element(By.CSS_SELECTOR, '[class="contraxtsSearch_searchField__button__2DHET"]').click()        

    contractNumber0 = ''
    try:
        f1 = WebDriverWait(driver,60).until(
    lambda driver: driver.find_elements(By.CSS_SELECTOR, '[class="contractsTable_tableLoading__1wlw8"]') or 
               driver.find_elements(By.CSS_SELECTOR, '[class="contractsTable_pixelGamingContractNumber__226ny"]')) 
        contractNumber0 = driver.find_element(By.CSS_SELECTOR, '[class="contractsTable_pixelGamingContractNumber__226ny"]').text
        time.sleep(1)

        if f1[0].text == 'Произошла ошибка, повторите попытку позже!':
            b1.append(datetime.now())
            b2.append(iin)
            b3.append('Создание трудового договора')
            b4.append('Не успешно')
            b5.append('Ошибка: Произошла ошибка, повторите попытку позже!')
            b6.append(contractNum)
            b7.append(FIO)
            b8.append(registerDate)
            b9.append(branch)
            continue

    except:
            print('Трудовой договор не найден')


    contractNumber = str(list(df_recruitment['Unnamed: 23'])[i])
    cN = contractNumber[:-2]
    if cN == 'n':
        b1.append(datetime.now())
        b2.append(iin)
        b3.append('Создание трудового договора')
        b4.append('Не успешно')
        b5.append('Ошибка: Номер договора пустой')
        b6.append(contractNumber)
        b7.append(FIO)
        b8.append(registerDate)
        b9.append(branch)
        continue

    try:
        index = list(df1['Штатная должность']).index(job2)
        job1 = df1['Соответствующая должность из Енбек кз'][index]
    except:
        #Ошибка: не найден маппинг должности. должность из 1c: 
        b1.append(datetime.now())
        b2.append(iin)
        b3.append('Создание трудового договора')
        b4.append('Не успешно')
        b5.append('Ошибка: не найден маппинг должности. должность из 1c: ' +job2)
        b6.append(contractNumber)
        b7.append(FIO)
        b8.append(registerDate)
        b9.append(branch)
        continue

    try:
        index = list(df2['место работы']).index(loc)
#         region = df2['Регион места работы'][index]
#         area = df2['Район места работы'][index]
#         Locality = df2['Населённый пункт'][index]
        
        region = 'Г.АЛМАТЫ'
        area = 'МЕДЕУСКИЙ РАЙОН'
        Locality = '0'
    except:
        #Ошибка: не найден маппинг место работы. место работы из 1c: 
        b1.append(datetime.now())
        b2.append(iin)
        b3.append('Создание трудового договора')
        b4.append('Не успешно')
        b5.append('Ошибка: Не найден маппинг место работы. место работы из 1c: '+str(loc))
        b6.append(contractNumber)
        b7.append(FIO)
        b8.append(registerDate)
        b9.append(branch)
        continue
        
    x1_s = 0
    if len(f1) >= 1:
        for x1 in f1:
            if x1.text == str(contractNumber):
                x1_s = 1
                b1.append(datetime.now())
                b2.append(iin)
                b3.append('Создание трудового договора')
                b4.append('Успешно')
                b5.append('Договор с таким номером существует')
                b6.append(contractNumber)
                b7.append(FIO)
                b8.append(registerDate)
                b9.append(branch)
        if x1_s == 1:
            continue

    z = 0
    for n in range(5):

        try:
            driver.find_element(By.CSS_SELECTOR, '[class="Button_icon__1FAE1"]').click()
        except:
            print('button error')
        time.sleep(2)

        d1 = datetime.strptime(dateFrom, '%d.%m.%Y')
        d2 = datetime.strptime(dateTo, '%d.%m.%Y')
        if d2-d1 != timedelta(days=365):
            dateTo = ''
            b1.append(datetime.now())
            b2.append(iin)
            b3.append('Создание трудового договора')
            b4.append('Не успешно')
            b5.append('Ошибка: период указан некорректно, проверьте информацию в 1С')
            b6.append(contractNumber)
            b7.append(FIO)
            b8.append(registerDate)
            b9.append(branch)
            break

        input_iin = driver.find_element(By.NAME, "iin")
        input_iin.send_keys(iin)
        time.sleep(1)


        k1 = driver.find_element(By.CSS_SELECTOR, '[class="ContractForm_fieldAndButton__2PPKt"]')
        k1.find_element(By.CSS_SELECTOR, '[type="button"]').click()
        time.sleep(4)

        # Номер договора *
        input_contractNumber = driver.find_element(By.NAME, "contractNumber")
        input_contractNumber.send_keys(contractNumber)

        # Срок действия трудового договора *
        if dateTo == '':
            lable_options = driver.find_elements(By.CSS_SELECTOR, '[class="Input_input__3bchA    style_input__1NppQ"]')
            lable_options[0].click()
            time.sleep(0.5)
            driver.find_elements(By.CSS_SELECTOR, '[class="style_option__i6QQC "]')[0].click()
        else:
            lable_options = driver.find_elements(By.CSS_SELECTOR, '[class="Input_input__3bchA    style_input__1NppQ"]')
            lable_options[0].click()
            time.sleep(0.5)
            driver.find_elements(By.CSS_SELECTOR, '[class="style_option__i6QQC "]')[1].click()

        # Дата подписания договора *
        input_registerDate = driver.find_element(By.NAME, "registerDate")
        input_registerDate.click()
        input_registerDate.send_keys(registerDate)
        time.sleep(0.5)

        # Дата начала работы *
        input_dateFrom = driver.find_element(By.NAME, "dateFrom")
        input_dateFrom.click()
        input_dateFrom.send_keys(dateFrom)
        time.sleep(0.5)

        # Дата окончания действия договора
        if dateTo != '':
            input_dateTo = driver.find_element(By.NAME, "dateTo")
            input_dateTo.click()
            input_dateTo.send_keys(dateTo)
            time.sleep(0.5)

        # Должность по НКЗ *
        input_job1 = lable_options[1]
        input_job1.send_keys(job1)

        # Должность *
        input_job2 = lable_options[2]
        input_job2.send_keys(job2)

        # Вид работы *
        lable_options[3].click()
        time.sleep(0.5)
        driver.find_elements(By.CSS_SELECTOR, '[class="style_option__i6QQC "]')[1].click()

        # Режим рабочего времени *
        lable_options[4].click()
        time.sleep(0.5)
        driver.find_elements(By.CSS_SELECTOR, '[class="style_option__i6QQC "]')[0].click()

        # Форма занятости *
        lable_options[5].click()
        time.sleep(0.5)
        driver.find_elements(By.CSS_SELECTOR, '[class="style_option__i6QQC "]')[2].click()


        # Регион места работы *
        lable_options[7].click()
        time.sleep(0.5)
        name = region
        while True:
            try:
                driver.find_element(By.XPATH, "//*[text()='%s']" %name)
                break
            except:
                element = driver.find_elements(By.CSS_SELECTOR, '[class="style_option__i6QQC "]')[-1:]
                driver.execute_script("arguments[0].scrollIntoView();", element[0])
                time.sleep(0.5)
        driver.find_element(By.XPATH, "//*[text()='%s']" %name).click()
        time.sleep(3)

        # Район места работы *
        lable_options[8].click()
        time.sleep(1)
        name = area
        while True:
            try:
                driver.find_element(By.XPATH, "//*[text()='%s']" %name)
                break
            except:
                element = driver.find_elements(By.CSS_SELECTOR, '[class="style_option__i6QQC "]')[-1:]
                driver.execute_script("arguments[0].scrollIntoView();", element[0])
                time.sleep(0.5)
        driver.find_element(By.XPATH, "//*[text()='%s']" %name).click()
        time.sleep(3)

        # Населённый пункт *
        if Locality != '0':
            lable_options[9].click()
            time.sleep(1)
            name = Locality
            while True:
                try:
                    driver.find_element(By.XPATH, "//*[text()='%s']" %name)
                    break
                except:
                    element = driver.find_elements(By.CSS_SELECTOR, '[class="style_option__i6QQC "]')[-1:]
                    driver.execute_script("arguments[0].scrollIntoView();", element[0])
                    time.sleep(0.5)
            driver.find_element(By.XPATH, "//*[text()='%s']" %name).click()
            time.sleep(1)

        # Адрес места работы *
        input_loc = driver.find_element(By.NAME, "workingPlace")
        input_loc.send_keys(loc)
        time.sleep(3)


        driver.find_element(By.XPATH, "//*[text()='Подписать ЭЦП и отправить']").click()
        time.sleep(0.5)
        lable_options[1].click()
        time.sleep(0.2)
        if driver.find_elements(By.CSS_SELECTOR, '[class="style_option__i6QQC "]')[0].text == job1:
            driver.find_element(By.CSS_SELECTOR, '[class="style_option__i6QQC "]').click()
        else:
            driver.find_elements(By.CSS_SELECTOR, '[class="style_option__i6QQC "]')[1].click()
            
        time.sleep(0.2)
        driver.find_element(By.XPATH, "//*[text()='Подписать ЭЦП и отправить']").click()
        time.sleep(3)

        ecp = 'ECP'
        pas = ''
        if list(df_recruitment['Unnamed: 24'])[i] == 'филиал АО «Home Credit Bank» в г.Кызылорда':
            ecp = 'ECP2'
            pas = ''
            
        pyautogui.write(ecp, interval=0.1)
        pyautogui.press('enter')
        time.sleep(1)
        pyautogui.keyDown('shift')
        pyautogui.press('tab')
        pyautogui.keyUp('shift')
        time.sleep(0.2)
        pyautogui.press('down')
        time.sleep(0.2)
        pyautogui.press('enter')
        time.sleep(0.2)
        pyautogui.write(pas, interval=0.1)
        pyautogui.press('enter')
        time.sleep(0.2)
        pyautogui.press('enter')
        time.sleep(3)

        status = driver.find_element(By.CSS_SELECTOR, '[class="ant-message-notice-content"]').text
        if status == 'Данные отправлены успешно':  

            inputElement1 = driver.find_element(By.CSS_SELECTOR, '[class="MuiInputBase-input css-mnn31"]')
            inputElement1.send_keys(Keys.CONTROL + "a")
            inputElement1.send_keys(Keys.DELETE)
            inputElement1.send_keys(iin)
            time.sleep(1) 

            try:
                driver.find_element(By.CSS_SELECTOR, '[class="contraxtsSearch_searchField__button__2DHET"]').click()
            except:
                driver.find_elements(By.XPATH, "//*[text()='Назад']")[0].click()
                time.sleep(1)
                driver.find_element(By.CSS_SELECTOR, '[class="contraxtsSearch_searchField__button__2DHET"]').click()


            f1 = WebDriverWait(driver,60).until(
            lambda driver: driver.find_elements(By.CSS_SELECTOR, '[class="contractsTable_tableLoading__1wlw8"]') or 
                       driver.find_elements(By.CSS_SELECTOR, '[class="contractsTable_pixelGamingContractNumber__226ny"]')) 

            if len(f1) >= 1:
                for x1 in f1:
                    if x1.text == str(contractNumber):
                        print(x1.text)
                        b1.append(datetime.now())
                        b2.append(iin)
                        b3.append('Создание трудового договора')
                        b4.append('Успешно')
                        b5.append('')
                        b6.append(contractNumber)
                        b7.append(FIO)
                        b8.append(registerDate)
                        b9.append(branch)
                        z = 1
                        break
        if z == 1:
            break
        time.sleep(720)
        if z == 0:                
            b1.append(datetime.now())
            b2.append(iin)
            b3.append('Создание трудового договора')
            b4.append('Не успешно:')
            b5.append('Ошибка: Енбек не доступен')
            b6.append(contractNumber)
            b7.append(FIO)
            b8.append(registerDate)
            driver.find_element(By.CSS_SELECTOR, '[class="style_back__2XOdr"]').click()
            b9.append(branch)
    if dateTo == '':
        continue
        

df_result1 = pd.DataFrame({ 'Дата':b1,
                            'ИИН':b2,
                            'ФИО':b7,
                            'Действие':b3,
                            'Дата начала ТД':b8,
                            'Статус':b4,
                            'Примечание':b5,
                            'Номер договора':b6,
                            'подразделения':b9 },
                            index=pd.RangeIndex(start=1,stop = len(b1)+1, name='index'))




df_job = df[((df['Unnamed: 6'] == 'Изменение должности') | (df['Unnamed: 6'] == 'Изменение подразделения, должности')) & ((df['Unnamed: 24'] == '(ГО) АО «Home Credit Bank»') | (df['Unnamed: 24'] == 'управление по г.Алматы'))] #| (df['Unnamed: 24'] == 'филиал АО «Home Credit Bank» в г.Кызылорда'))]

b1 = []
b2 = []
b3 = []
b4 = []
b5 = []
b6 = []
b7 = []
b8 = []
b9 = []


for i in range(len(list(df_job['Unnamed: 4']))):
    print(i)
    iin = str(list(df_job['Unnamed: 4'])[i])
    for n in range(3):
        if len(str(iin)) != 12:
            iin = str('0'+ str(iin))
    print(iin)
    FIO = list(df_job['Unnamed: 16'])[i]
    contractNum = list(df_job['Unnamed: 23'])[i]

    dateFrom_text = list(df_job['Unnamed: 8'])[i]
    registerDate = list(df_job['Unnamed: 9'])[i]
    dateFrom = list(df_job['Unnamed: 9'])[i]
    dateTo = list(df_job['Unnamed: 10'])[i]
    dpositionCode = list(df_job['Unnamed: 11'])[i]
    branch = list(df_job['Unnamed: 24'])[i]

    inputElement1 = driver.find_element(By.CSS_SELECTOR, '[class="MuiInputBase-input css-mnn31"]')
    inputElement1.send_keys(Keys.CONTROL + "a")
    inputElement1.send_keys(Keys.DELETE)
    inputElement1.send_keys(iin)
    time.sleep(1)

    try:
        driver.find_element(By.CSS_SELECTOR, '[class="contraxtsSearch_searchField__button__2DHET"]').click()
    except:
        driver.find_elements(By.XPATH, "//*[text()='Назад']")[0].click()
        time.sleep(1)
        driver.find_element(By.CSS_SELECTOR, '[class="contraxtsSearch_searchField__button__2DHET"]').click()


    f1 = WebDriverWait(driver,60).until(
    lambda driver: driver.find_elements(By.CSS_SELECTOR, '[class="contractsTable_tableLoading__1wlw8"]') or 
               driver.find_elements(By.CSS_SELECTOR, '[class="contractsTable_pixelGamingContractNumber__226ny"]')) 

    time.sleep(1)

    cc = 0
    ccc = 0
    if len(f1) >= 1:
        for x1 in f1:
            if x1.text == 'Данных согласно запросу не найдено':
                cc = 1
                ccc = 1
                b1.append(datetime.now())
                b2.append(iin)
                b3.append('Создание допсоглашение')
                b4.append('Не успешно')
                b5.append('Ошибка: Трудовой договор не найден')
                b6.append(contractNum)
                b7.append(FIO)
                b8.append(registerDate)
                b9.append(branch)


            if x1.text == str(contractNum):
                ccc = 1
                x1.click()
                time.sleep(1)

            elif x1.text == 'Произошла ошибка, повторите попытку позже!':
                cc = 1
                ccc = 1
                b1.append(datetime.now())
                b2.append(iin)
                b3.append('Создание допсоглашение')
                b4.append('Не успешно')
                b5.append('Ошибка: Произошла ошибка, повторите попытку позже!')
                b6.append(contractNum)
                b7.append(FIO)
                b8.append(registerDate)
                b9.append(branch)
        if ccc == 0: 
            b1.append(datetime.now())
            b2.append(iin)
            b3.append('Создание допсоглашение')
            b4.append('Не успешно')
            b5.append('Ошибка: Трудовой договор не найден')
            b6.append(contractNum)
            b7.append(FIO)
            b8.append(registerDate)
            b9.append(branch)
            continue
        if cc == 1:
            continue

    s_text0 = WebDriverWait(driver,30).until(lambda driver: driver.find_elements(By.CSS_SELECTOR, '[class="ContractCard_detailInformation__uiD3I"]'))[0]
    s_text = s_text0.find_element(By.CSS_SELECTOR, 'span').text
    if s_text == 'Расторгнутый':

        b1.append(datetime.now())
        b2.append(iin)
        b3.append('Создание допсоглашение')
        b4.append('Не успешно')
        b5.append('Ошибка: договор расторгнутый')
        b6.append(contractNum)
        b7.append(FIO)
        b8.append(registerDate)
        b9.append(branch)
        print('no button')

        WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "//*[text()='Назад']"))).click()
        continue

    v22 = 0
    cc2 = 0
    v = driver.find_element(By.CSS_SELECTOR, '[class="AdditionalContracts_table__2DiFr"]').find_elements(By.CSS_SELECTOR, 'tr')[1:]

    actions = ActionChains(driver)
    print('test')
    time.sleep(1) 
    v_i = []
    for v2 in range(len(v)):
        if v[v2].find_elements(By.CSS_SELECTOR, 'td')[1].text == registerDate and v[v2].find_elements(By.CSS_SELECTOR, 'td')[5].text =='Подписан':
            if v[v2].find_elements(By.CSS_SELECTOR, 'td')[0].text != 'Б/Н' or v[v2].find_elements(By.CSS_SELECTOR, 'td')[3].text != '-':
                v22 = 1
                v_i.append(v2)
    print(v_i)
    if len(v_i) == 1:
        if v[v_i[0]].find_elements(By.CSS_SELECTOR, 'td')[1].text == v[v_i[0]].find_elements(By.CSS_SELECTOR, 'td')[3].text:
            actions.move_to_element(v[v_i[0]]).perform() 
            v[v_i[0]].find_element(By.CSS_SELECTOR, '[class="DropdownMenu_dropdownButton__2xihO undefined"]').click()
            v[v_i[0]].find_element(By.CSS_SELECTOR, '[class="DropdownMenu_ul__3zssB DropdownMenu_shown__2q_x1 "]').find_elements(By.CSS_SELECTOR, 'li')[1].click()

            time.sleep(1)
            inputElement1 = driver.find_element(By.CSS_SELECTOR, '[class="Input_input__3bchA   "]')
            inputElement1.send_keys(Keys.CONTROL + "a")
            inputElement1.send_keys(Keys.DELETE)
            inputElement1.send_keys('Б/Н')

            time.sleep(1)
            inputElement2 = driver.find_elements(By.CSS_SELECTOR, '[class="DatePicker_input__3cTs1 "]')[2]
            inputElement2.send_keys(Keys.CONTROL + "a")
            inputElement2.send_keys(Keys.DELETE)

            driver.find_element(By.XPATH, "//*[text()='Продолжить']").click()
            time.sleep(0.5)
            inputElement1.click()
            time.sleep(1)
            driver.find_element(By.XPATH, "//*[text()='Сохранить и подписать ЭЦП']").click()

        else:
            b1.append(datetime.now())
            b2.append(iin)
            b3.append('Создание допсоглашение')
            b4.append('Успешно')
            b5.append('Номер договора с указонной датой существует')
            b6.append(contractNum)
            b7.append(FIO)
            b8.append(registerDate)
            b9.append(branch)
            continue

    v24 = 0
    if len(v_i) > 1:
        for v_i1 in v_i:
            print(v[v_i1].find_elements(By.CSS_SELECTOR, 'td')[1].text, v[v_i1].find_elements(By.CSS_SELECTOR, 'td')[3].text)
            if v[v_i1].find_elements(By.CSS_SELECTOR, 'td')[1].text == v[v_i1].find_elements(By.CSS_SELECTOR, 'td')[3].text:
                v24 = 1
                actions.move_to_element(v[v_i1]).perform() 
                v[v_i1].find_element(By.CSS_SELECTOR, '[class="DropdownMenu_dropdownButton__2xihO undefined"]').click()
                v[v_i1].find_element(By.CSS_SELECTOR, '[class="DropdownMenu_ul__3zssB DropdownMenu_shown__2q_x1 "]').find_elements(By.CSS_SELECTOR, 'li')[2].click()
                time.sleep(1)
                driver.find_element(By.XPATH, "//*[text()='Да, подписать ЭЦП и удалить']").click()


        if v24 == 0:
            b1.append(datetime.now())
            b2.append(iin)
            b3.append('Создание допсоглашение')
            b4.append('Не успешно')
            b5.append('Ошибка: Другая ЭЦП')
            b6.append(contractNum)
            b7.append(FIO)
            b8.append(registerDate)
            b9.append(branch)
            continue

    v23 = 0            
    if v22 == 0:
        for v2 in range(len(v)):
            if v[v2].find_elements(By.CSS_SELECTOR, 'td')[1].text == registerDate and v[v2].find_elements(By.CSS_SELECTOR, 'td')[5].text =='Подписан':
                if v[v2].find_elements(By.CSS_SELECTOR, 'td')[0].text == 'Б/Н' or v[v2].find_elements(By.CSS_SELECTOR, 'td')[3].text == '-':
                    v23 = 1
        if v23 == 1:
            b1.append(datetime.now())
            b2.append(iin)
            b3.append('Создание допсоглашение')
            b4.append('Успешно')
            b5.append('Номер договора с указонной датой существует')
            b6.append(contractNum)
            b7.append(FIO)
            b8.append(registerDate)
            b9.append(branch)
            continue

        try:
            WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "//*[text()='Создать допсоглашение']"))).click()


        except:
            b1.append(datetime.now())
            b2.append(iin)
            b3.append('Создание допсоглашение')
            b4.append('Не успешно')
            b5.append('Ошибка: договор расторгнутый')
            b6.append(contractNum)
            b7.append(FIO)
            b8.append(registerDate)
            b9.append(branch)
            print('no button')

            WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "//*[text()='Назад']"))).click()
            continue

        time.sleep(1)
        # Информация о дополнительном соглашении
        #


        try:
            index = list(df1['Штатная должность']).index(dpositionCode)
            job = df1['Соответствующая должность из Енбек кз'][index]
        except:
            #Ошибка: не найден маппинг должности. должность из 1c: 
            b1.append(datetime.now())
            b2.append(iin)
            b3.append('Создание допсоглашение')
            b4.append('Не успешно')
            b5.append('Ошибка: не найден маппинг должности. должность из 1c: ' +dpositionCode)
            b6.append(contractNum)
            b7.append(FIO)
            b8.append(registerDate)
            b9.append(branch)

            driver.find_elements(By.XPATH, "//*[text()='Назад']")[1].click()
            driver.find_elements(By.XPATH, "//*[text()='Назад']")[0].click()
            continue

        #Номер дополнительного соглашения *
        input_contractNum = driver.find_element(By.NAME, "contractNum")
        input_contractNum.send_keys('Б/Н')
        time.sleep(1)

        #Дата заключения доп.соглашения *
        input_registerDate = driver.find_element(By.NAME, "registerDate")
        input_registerDate.click()
        input_registerDate.send_keys(registerDate)
        time.sleep(1)

        #Дата начала действия доп.соглашения *
        input_dateFrom = driver.find_element(By.NAME, "dateFrom")
        input_dateFrom.click()
        input_dateFrom.send_keys(dateFrom)
        time.sleep(1)
        input_contractNum.click()
        time.sleep(1)


        driver.find_element(By.XPATH, "//*[text()='Выбрать']").click()
        WebDriverWait(driver,5).until(lambda driver: driver.find_element(By.XPATH, "//*[text()='Продолжить']")).click()
        driver.find_element(By.XPATH, "//*[text()='Изменение должности']").click()
        time.sleep(1)
        try:
            driver.find_element(By.XPATH, "//*[text()='Продолжить']").click()
        except:
            print('error Продолжить')
        time.sleep(1)

        # Изменение должности
        #Должность по НКЗ *
        input_dpositionCode = driver.find_elements(By.CSS_SELECTOR, '[class="Input_input__3bchA    style_input__1NppQ"]')[0]
        input_dpositionCode.send_keys(job)
        time.sleep(1)
        if len(driver.find_elements(By.CSS_SELECTOR, '[class="style_option__i6QQC "]')) > 1:
            for l in range(10):
                element = driver.find_elements(By.CSS_SELECTOR, '[class="style_option__i6QQC "]')[-1:]
                driver.execute_script("arguments[0].scrollIntoView();", element[0])
                time.sleep(1)

            element[0].click()
        else:      
            time.sleep(1)
            driver.find_element(By.CSS_SELECTOR, '[class="style_option__i6QQC "]').click()
            time.sleep(1) 

        #Должность *
        input_job = driver.find_elements(By.CSS_SELECTOR, '[class="Input_input__3bchA    style_input__1NppQ"]')[1]
        input_job.send_keys(dpositionCode)
        time.sleep(1)

        #Сохранить и подписать ЭЦП
        driver.find_element(By.XPATH, "//*[text()='Сохранить и подписать ЭЦП']").click()
    time.sleep(1)

    # Running the aforementioned command and saving its output
    output = os.popen('wmic process get description, processid').read()
    data_list = ''.join(output).split('\n')
    searh = 'javaw.exe'
    text = ''
    for k in data_list:
        if searh in k:
            text = k

    process_nums = int(re.findall(r'\d+', text)[0])
    print(process_nums)

    Dialog = 0
    for z in range(5):
        try:
            app = Application(backend='uia').connect(process = process_nums)
            app.top_window().set_focus()
            Dialog = 1
        except:
            print('error Dialog')
        if Dialog == 1:
            break
        
    ecp = 'ECP'
    pas = ''
    if list(df_job['Unnamed: 24'])[i] == 'филиал АО «Home Credit Bank» в г.Кызылорда':
        ecp = 'ECP2'
        pas = ''
        
    pyautogui.write(ecp, interval=0.1)
    pyautogui.press('enter')
    time.sleep(0.2)
    pyautogui.keyDown('shift')
    pyautogui.press('tab')
    pyautogui.keyUp('shift')
    time.sleep(0.2)
    pyautogui.press('down')
    time.sleep(0.2)
    pyautogui.press('enter')
    time.sleep(0.5)
    pyautogui.write(pas, interval=0.1)
    time.sleep(0.5)
    pyautogui.press('enter')
    time.sleep(0.5)
    pyautogui.press('enter')
    time.sleep(1)

    try:
        WebDriverWait(driver,5).until(lambda driver: driver.find_elements(By.XPATH, "//*[text()='Операция успешно выполнена']"))
        b1.append(datetime.now())
        b2.append(iin)
        b3.append('Создание допсоглашение')
        b4.append('Успешно')
        b5.append('')
        b6.append(contractNum)
        b7.append(FIO)
        b8.append(registerDate)
        b9.append(branch)

        time.sleep(3)
        WebDriverWait(driver,5).until(lambda driver: driver.find_elements(By.XPATH, "//*[text()='Назад']"))[0].click()

    except:
        WebDriverWait(driver,5).until(lambda driver: driver.find_elements(By.XPATH, "//*[text()='Назад']"))[1].click()
#         driver.find_elements(By.XPATH, "//*[text()='Назад']")[1].click()
        time.sleep(1)
        driver.find_elements(By.XPATH, "//*[text()='Назад']")[0].click()
        b1.append(datetime.now())
        b2.append(iin)
        b3.append('Создание допсоглашение')
        b4.append('Успешно')
        b5.append('Номер договора с указонной датой существует')
        b6.append(contractNum)
        b7.append(FIO)
        b8.append(registerDate)
        b9.append(branch)


    time.sleep(5)
    
df_vacation = df[((df['Unnamed: 6'] == 'Отпуск по беременности и родам') | (df['Unnamed: 6'] == 'Отпуск по уходу за ребенком') | (df['Unnamed: 6'] == 'Возврат на работу')) & ((df['Unnamed: 24'] == '(ГО) АО «Home Credit Bank»') | (df['Unnamed: 24'] == 'управление по г.Алматы'))]#| (df['Unnamed: 24'] == 'филиал АО «Home Credit Bank» в г.Кызылорда'))]

d = list(df_vacation['Unnamed: 9'])
a = []
for i in range(len(list(df_vacation['Unnamed: 9']))):
    datetime_object = datetime.strptime(d[i], '%d.%m.%Y')
    if datetime.now()-timedelta(days = 2) <= datetime_object:
        a.append(i)
        print(datetime_object , i)

b1 = []
b2 = []
b3 = []
b4 = []
b5 = []
b6 = []
b7 = []
b9 = []
b8 = []
# for i in range(47,len(df_vacation)):
for i in a:
# Добавить соцотпуск
#     try:
    iin = str(list(df_vacation['Unnamed: 4'])[i])
    for n in range(3):
        if len(str(iin)) != 12:
            iin = '0'+ str(iin)
            
    FIO = list(df_vacation['Unnamed: 16'])[i]
    action = list(df_vacation['Unnamed: 6'])[i]
    branch = list(df_vacation['Unnamed: 24'])[i]
    beginDate = (list(df_vacation['Unnamed: 9'])[i])
    
    inputElement1 = driver.find_element(By.CSS_SELECTOR, '[class="MuiInputBase-input css-mnn31"]')
    inputElement1.send_keys(Keys.CONTROL + "a")
    inputElement1.send_keys(Keys.DELETE)
    inputElement1.send_keys(iin)
    time.sleep(1)

    print(iin)
    contractNum = list(df_vacation['Unnamed: 23'])[i]
    
    try:
        driver.find_element(By.CSS_SELECTOR, '[class="contraxtsSearch_searchField__button__2DHET"]').click()
    except:
        driver.find_elements(By.XPATH, "//*[text()='Назад']")[0].click()
        time.sleep(1)
        driver.find_element(By.CSS_SELECTOR, '[class="contraxtsSearch_searchField__button__2DHET"]').click()   
#     time.sleep(5)
    
    print('1')

    try:
#         driver.find_element(By.CSS_SELECTOR, '[class="contractsTable_pixelGamingContractNumber__3OFq9"]').click()
        WebDriverWait(driver,20).until(lambda driver: driver.find_element(By.CSS_SELECTOR, '[class="contractsTable_pixelGamingContractNumber__226ny"]')).click()
        time.sleep(2)
    except:
#         e = driver.find_element(By.CSS_SELECTOR, '[class="contractsTable_pixelGamingContractNumber__226ny"]').text
        b1.append(datetime.now())
        b2.append(iin)
        b3.append(action)
        b4.append('Не успешно')
        b5.append('Ошибка: Трудовой договор не найден')
        b6.append(contractNum)
        b7.append(FIO)
        b9.append(branch)
        b8.append(beginDate)
        continue
    
    if list(df_vacation['Unnamed: 6'])[i] == 'Возврат на работу':
        print('3')
        beginDate = (datetime.strptime(beginDate, '%d.%m.%Y')-timedelta(days = 1)).strftime("%d.%m.%Y")
        v1 = []
        v = driver.find_element(By.CSS_SELECTOR, '[class="Table_table__2OuB7"]').find_elements(By.CSS_SELECTOR, 'tr')[1:]
        for v2 in v:
            if v2.find_element(By.CSS_SELECTOR, '[class="SocialLeavesTable_typeCol__36wuR"]').text == 'Без сохранения заработной платы по уходу за ребенком до достижения им возраста 3 лет':
                v1.append(v2)
        iv = 0
        for v3 in v1:
            date_v1 = v3.find_elements(By.CSS_SELECTOR, 'td')[3].text
            date_v2 = v3.find_elements(By.CSS_SELECTOR, 'td')[2].text
            if beginDate != date_v1:
                if datetime.strptime(date_v1, '%d.%m.%Y') >= datetime.strptime(beginDate, '%d.%m.%Y') >= datetime.strptime(date_v2, '%d.%m.%Y'):
                    v3.find_element(By.CSS_SELECTOR, '[class="DropdownMenu_dropdownButton__2xihO undefined"]').click()
                    time.sleep(1)
                    v3.find_element(By.CSS_SELECTOR, '[class="DropdownMenu_ul__3zssB DropdownMenu_shown__2q_x1 "]').find_elements(By.CSS_SELECTOR, 'li')[1].click()
                    
                    inputElement3 = driver.find_elements(By.CSS_SELECTOR, '[class="DatePicker_input__3cTs1 "]')[1]
                    inputElement3.send_keys(Keys.CONTROL + "a")
                    inputElement3.send_keys(Keys.DELETE)
                    inputElement3.send_keys(beginDate)
            else:
                iv = 1
                b1.append(datetime.now())
                b2.append(iin)
                b3.append(action)
                b4.append('Успешно')
                b5.append('дата возврат на работу уже отредактирован')
                b6.append(contractNum)
                b7.append(FIO)
                b9.append(branch)
                b8.append(beginDate)
        if iv == 1:
            continue                
#         driver.find_element(By.XPATH, "//*[text()='Сохранить и подписать ЭЦП']").click()
    else:
        try:
            time.sleep(1)
            WebDriverWait(driver,10).until(lambda driver: driver.find_element(By.XPATH, "//*[text()='Добавить соцотпуск']")).click()
            time.sleep(1)
            print('2')
        except:
            e = driver.find_element(By.XPATH, "//*[text()='Назад']").click()
            b1.append(datetime.now())
            b2.append(iin)
            b3.append(action)
            b4.append('Не успешно')
            b5.append('Ошибка: Договор расторгнутый')
            b6.append(contractNum)
            b7.append(FIO)
            b9.append(branch)
            b8.append(beginDate)
            continue

        driver.find_element(By.CSS_SELECTOR, '[placeholder="Выберите из списка"]').click()
        if list(df_vacation['Unnamed: 6'])[i] == 'Отпуск по уходу за ребенком':
        #     Без сохранения заработной платы по уходу за ребенком до достижения им возраста 3 лет
            driver.find_elements(By.CSS_SELECTOR, '[class="style_option__i6QQC "]')[1].click()


            endDate = (list(df_vacation['Unnamed: 10'])[i])

            # Не работал(а) с*
            input_beginDate = driver.find_element(By.NAME, "beginDate")
            input_beginDate.click()
            input_beginDate.send_keys(beginDate)
            time.sleep(1)

            # Не работал(а) по*
            input_endDate = driver.find_element(By.NAME, "endDate")
            input_endDate.click()
            input_endDate.send_keys(endDate)
            time.sleep(1)


        elif list(df_vacation['Unnamed: 6'])[i] == 'Отпуск по беременности и родам':

        # В связи с беременностью и рождением ребенка
            driver.find_elements(By.CSS_SELECTOR, '[class="style_option__i6QQC "]')[0].click()

            endDate = (list(df_vacation['Unnamed: 10'])[i])
            firstDayDate = (datetime.strptime(endDate, '%d.%m.%Y')+timedelta(days = 1)).strftime("%d.%m.%Y")
            daysOff = str(datetime.strptime(endDate, '%d.%m.%Y')-datetime.strptime(beginDate, '%d.%m.%Y')).split(' ')[0]
            timeSheetNum = 'б/н'

            # Не работал(а) с*
            input_beginDate = driver.find_element(By.NAME, "beginDate")
            input_beginDate.click()
            input_beginDate.send_keys(beginDate)
            time.sleep(1)

            # Не работал(а) по*
            input_endDate = driver.find_element(By.NAME, "endDate")
            input_endDate.click()
            input_endDate.send_keys(endDate)
            time.sleep(1)

            # Дата первого рабочего дня*
            input_firstDayDate = driver.find_element(By.NAME, "firstDayDate")
            input_firstDayDate.click()
            input_firstDayDate.send_keys(firstDayDate)
            time.sleep(1)

            # Выходные дни за период нетрудоспособности
    #         input_daysOff = driver.find_element(By.NAME, "daysOff")
    #         input_daysOff.click()
    #         input_daysOff.send_keys(daysOff)
    #         time.sleep(1)

            # Номер табеля*
            input_timeSheetNum = driver.find_element(By.NAME, "timeSheetNum")
            input_timeSheetNum.send_keys(timeSheetNum)
            time.sleep(1)
            input_timeSheetNum.click()
            time.sleep(1)



        
        
        
    #     Сохранить и подписать ЭЦП
    driver.find_element(By.XPATH, "//*[text()='Сохранить и подписать ЭЦП']").click()
    time.sleep(1)
    
    # Running the aforementioned command and saving its output
    output = os.popen('wmic process get description, processid').read()
    # Displaying the output
    # print(output)
    data_list = ''.join(output).split('\n')
    searh = 'javaw.exe'
    text = ''
    for k in data_list:
        if searh in k:
            text = k

    process_nums = int(re.findall(r'\d+', text)[0])
    print(process_nums)

    Dialog = 0
    for z in range(5):
        try:
            app = Application(backend='uia').connect(process = process_nums)
            app.top_window().set_focus()
            Dialog = 1
        except:
            print('error Dialog')
        if Dialog == 1:
            break
        
    ecp = 'ECP'
    pas = ''
    if list(df_vacation['Unnamed: 24'])[i] == 'филиал АО «Home Credit Bank» в г.Кызылорда':
        ecp = 'ECP2'
        pas = ''
        
    pyautogui.write(ecp, interval=0.1)
    pyautogui.press('enter')
    time.sleep(0.2)
    pyautogui.keyDown('shift')
    pyautogui.press('tab')
    pyautogui.keyUp('shift')
    time.sleep(0.2)
    pyautogui.press('down')
    time.sleep(0.2)
    pyautogui.press('enter')
    time.sleep(0.2)
    pyautogui.write(pas, interval=0.1)
    pyautogui.press('enter')
    time.sleep(0.2)
    pyautogui.press('enter')
    time.sleep(1)

    try:
        WebDriverWait(driver,5).until(lambda driver: driver.find_elements(By.XPATH, "//*[text()='Операция выполнена успешно']"))
        b1.append(datetime.now())
        b2.append(iin)
        b3.append(action)
        b4.append('Успешно')
        b5.append('')
        b6.append(contractNum)
        b7.append(FIO)
        b9.append(branch)
        b8.append(beginDate)
        time.sleep(3)
        WebDriverWait(driver,15).until(lambda driver: driver.find_elements(By.XPATH, "//*[text()='Назад']"))[0].click()

    except:
        WebDriverWait(driver,5).until(lambda driver: driver.find_elements(By.XPATH, "//*[text()='Назад']"))[0].click()
#         driver.find_elements(By.XPATH, "//*[text()='Назад']")[1].click()
        time.sleep(1)
#             driver.find_elements(By.XPATH, "//*[text()='Назад']")[0].click()
        b1.append(datetime.now())
        b2.append(iin)
        b3.append(action)
        b4.append('Успешно')
        b5.append('Социальный отпуск с указаной датой "Не работал с" к данному трудовому договору уже существует')
        b6.append(contractNum)
        b7.append(FIO)
        b9.append(branch)
        b8.append(beginDate)



    
df_result3 = pd.DataFrame({ 'Дата':b1,
                            'ИИН':b2,
                            'ФИО':b7,
                            'Действие':b3,
                            'Статус':b4,
                            'Примечание':b5,
                            'Номер договора':b6,
                            'подразделения':b9,
                            'дата начала события':b8},
                             index=pd.RangeIndex(start=1,stop = len(b1)+1, name='index')) 

  
    
df_dismissal = df[(df['Unnamed: 6'] == 'Увольнение') & ((df['Unnamed: 24'] == '(ГО) АО «Home Credit Bank»') | (df['Unnamed: 24'] == 'управление по г.Алматы'))] #| (df['Unnamed: 24'] == 'филиал АО «Home Credit Bank» в г.Кызылорда'))]

b1 = []
b2 = []
b3 = []
b4 = []
b5 = []
b6 = []
b7 = []
b8 = []
b9 = []


for i in range(len(list(df_dismissal['Unnamed: 4']))):
    iin = str(list(df_dismissal['Unnamed: 4'])[i])
    for n in range(3):
        if len(str(iin)) != 12:
            iin = '0'+ str(iin)
    
    FIO = list(df_dismissal['Unnamed: 16'])[i]
    date = list(df_dismissal['Unnamed: 19'])[i]  
    contractNum = list(df_dismissal['Unnamed: 23'])[i]
    branch = list(df_dismissal['Unnamed: 24'])[i]
    
    inputElement1 = driver.find_element(By.CSS_SELECTOR, '[class="MuiInputBase-input css-mnn31"]')
    inputElement1.send_keys(Keys.CONTROL + "a")
    inputElement1.send_keys(Keys.DELETE)
    inputElement1.send_keys(iin)
    time.sleep(1)

    print(iin)
    driver.find_element(By.CSS_SELECTOR, '[class="contraxtsSearch_searchField__button__2DHET"]').click()
    
    time.sleep(1)
    
    try:
        x = WebDriverWait(driver,10).until(lambda driver: driver.find_elements(By.CSS_SELECTOR, '[class="contractsTable_pixelGamingContractNumber__226ny"]'))
    except:
        b1.append(datetime.now())
        b2.append(iin)
        b3.append('расторжение')
        b4.append('Не успешно')
        b5.append('Ошибка: Трудовой договор не найден')
        b6.append(contractNum)
        b7.append(FIO)
        b8.append(date)
        b9.append(branch)
        continue
        
    time.sleep(1)
    try:
        if len(x) > 1:
            for x1 in x:
                if x1.text == str(contractNum):
                    x1.click()
        elif len(x) == 1:
            x[0].click()
    except:
        print('err')


    try:    
#         driver.find_element(By.XPATH, "//*[text()='Расторгнуть']").click()
        WebDriverWait(driver,10).until(lambda driver: driver.find_elements(By.XPATH, "//*[text()='Расторгнуть']"))[0].click()
    except:
        WebDriverWait(driver,5).until(lambda driver: driver.find_elements(By.XPATH, "//*[text()='Назад']"))[0].click()
#         driver.find_element(By.XPATH, "//*[text()='Назад']").click()
        b1.append(datetime.now())
        b2.append(iin)
        b3.append('расторжение')
        b4.append('Успешно')
        b5.append('Договор был ранее расторгнутый')
        b6.append(contractNum)
        b7.append(FIO)
        b8.append(date)
        b9.append(branch)
        continue
    

    
    driver.find_element(By.XPATH, "//*[text()='Понятно, все равно расторгнуть']").click()
    time.sleep(1)
    
    
    input_date = driver.find_element(By.NAME, "terminationDate")
    input_date.click()
    input_date.send_keys(date)
    time.sleep(1)
    driver.find_element(By.CSS_SELECTOR, '[class="Input_input__3bchA    style_input__1NppQ"]').click()
    time.sleep(1)
    driver.find_element(By.CSS_SELECTOR, '[class="style_option__i6QQC "]').click()
    
    
    driver.find_element(By.XPATH, "//*[text()='Подписать ЭЦП']").click()
    time.sleep(2)
    
    # Running the aforementioned command and saving its output
    output = os.popen('wmic process get description, processid').read()
    # Displaying the output
    # print(output)
    data_list = ''.join(output).split('\n')
    searh = 'javaw.exe'
    text = ''
    for n in data_list:
        if searh in n:
            text = n

    process_nums = int(re.findall(r'\d+', text)[0])
    print(process_nums)

    Dialog = 0
    for z in range(5):
        try:
            app = Application(backend='uia').connect(process = process_nums)
            app.top_window().set_focus()
            Dialog = 1
        except:
            print('error Dialog')
        if Dialog == 1:
            break
    ecp = 'ECP'
    pas = ''
    if list(df_dismissal['Unnamed: 24'])[i] == 'филиал АО «Home Credit Bank» в г.Кызылорда':
        ecp = 'ECP2'
        pas = ''
        
    pyautogui.write(ecp, interval=0.1)
    pyautogui.press('enter')
    time.sleep(0.2)
    pyautogui.keyDown('shift')
    pyautogui.press('tab')
    pyautogui.keyUp('shift')
    time.sleep(0.2)
    pyautogui.press('down')
    time.sleep(0.2)
    pyautogui.press('enter')
    time.sleep(0.2)
    pyautogui.write(pas, interval=0.1)
    pyautogui.press('enter')
    time.sleep(0.2)
    pyautogui.press('enter')
    time.sleep(1)
    pyautogui.press('enter')
    
    
    b1.append(datetime.now())
    b2.append(iin)
    b3.append('расторжение')
    b4.append('Успешно')
    b5.append('')
    b6.append(contractNum)
    b7.append(FIO)
    b8.append(date)
    b9.append(branch)
    time.sleep(2)
    
df_result4 = pd.DataFrame({ 'Дата':b1,
                            'ИИН':b2,
                            'ФИО':b7,
                            'Действие':b3,
                            'Дата увольнения по приказу':b8,
                            'Статус':b4,
                            'Примечание':b5,
                            'Номер договора':b6,
                            'подразделения':b9},
                             index=pd.RangeIndex(start=1,stop = len(b1)+1, name='index')) 


writer = pd.ExcelWriter('file name')

# Write each dataframe to a different worksheet.
df_result1.to_excel(writer, sheet_name='Прием на работу')
df_result2.to_excel(writer, sheet_name='допсоглашение')
df_result3.to_excel(writer, sheet_name='соцотпуск')
df_result4.to_excel(writer, sheet_name='расторжение')

# Close the Pandas Excel writer and output the Excel file.
writer.save() 


        