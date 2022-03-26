import PySimpleGUI as sg
import time
from selenium import webdriver
from subprocess import CREATE_NO_WINDOW
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import NoSuchElementException
      
sg.theme('BluePurple')
   
layout = [[sg.Text(key='-OUTPUT-')],[sg.Text('USERNAME '),sg.Input(size=(35,2),key='-IN-')],
[sg.Text('PASSWORD'),sg.Input(size=(35,2),key='-PASS-')],[sg.Button('SUBMIT'), sg.Button('Exit')]]
  
window = sg.Window('Introduction', layout)
  
while True:
    event, values = window.read()
    print(event, values)
      
    if event in  (None, 'Exit'):
        break
      
    if event == 'SUBMIT':
        ser = Service("chromedriver.exe")
        ser.creationflags = CREATE_NO_WINDOW

        driver=webdriver.Chrome(service=ser)
        driver.maximize_window()


        driver.get("https://mastersolutions.online/Home")
        driver.find_element(By.XPATH,"/html/body/header/div/div/div[2]/div[2]/a[1]").click()

        #USERNAME
        #"MSTU6399"
        driver.find_element(By.XPATH,"//*[@id=\"UserName\"]").send_keys(values['-IN-'])

        #PASSWORD
        #"Rs@2000"
        driver.find_element(By.XPATH,"//*[@id=\"Password\"]").send_keys(values['-PASS-'])

        #LOGIN
        driver.find_element(By.XPATH,"/html/body/div[2]/div/div/div/div/form/div[3]/input[1]").click()

        #WAIT
        #time.sleep(3)

        #Cancel_Notification
        #driver.find_element(By.XPATH,"//*[@id=\"NotificationModel\"]/div/div/div[1]/button/span").click()

        #Todays Work
        #print(driver.find_element(By.XPATH,"/html/body/div[2]/div/div/div/div/form/div[2]/div"))
        try:
            driver.find_element(By.XPATH,"//*[@id=\"sectionsNav\"]/div/div[2]/ul/li[3]/a").click()
            #Job Work
            driver.find_element(By.XPATH,"/html/body/div/div[2]/section/section/div[2]/div/div/div/a").click()
            print(int(driver.find_element(By.XPATH,"/html/body/div/div[2]/section/section/div[2]/div/div[3]/div[2]/div/div[2]/h3").get_attribute('innerHTML')))

            #fetch Imageand Enter the Code
            while int(driver.find_element(By.XPATH,"/html/body/div/div[2]/section/section/div[2]/div/div[3]/div[2]/div/div[2]/h3").get_attribute('innerHTML')) !=0 :
                src= driver.find_element(By.CLASS_NAME,"noselect").get_attribute('innerHTML')
                print(src)
                driver.find_element(By.XPATH,"//*[@id=\"EnteredCaptcha\"]").send_keys(src,Keys.ENTER)
            else:
               window['-OUTPUT-'].update("You are Done with your Today's Job")
        except NoSuchElementException:
             window['-OUTPUT-'].update("Wrong USERNAME or PASSWORD")

        
window.close()
