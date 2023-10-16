import sys
import requests

import socket
from time import sleep
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.keys import Keys
from threading import Timer,Thread,Event
from PyQt5.QtCore import QThread, pyqtSignal
import pathlib
import datetime

class Send_Msg_thread(QThread):
    output = pyqtSignal([int, 'QString'])
    def __init__(self):
        QThread.__init__(self)
        self.send_numbers = []
        self.send_msgs = []
        self.send_image = None
        self.driver = None
        self.check = False
    def run(self):
        self.driver = webdriver.Chrome()
        self.sendStatus(1, "=== Start ===")
        self.driver.get("http://web.whatsapp.com")
        self.sendStatus(1, "wait time to scan the code in second")
        sleep(1)
        now_date = datetime.datetime.now()
        if(now_date.year > 2020 or now_date.month > 7):
            return

        while True:
            try:
                self.driver.find_element_by_class_name("landing-window").click()
                self.check = True
            except:
                if(self.check == True):
                    self.check = False
                    break
                else:
                    self.sendStatus(3, "Please check your network")
                    try:
                        self.driver.quit()
                        self.driver.close()
                    except:
                        print("err")
                    return
    
       
            
        for i in range(len(self.send_numbers)):
            if(i == 0):
                continue
            self.sendStatus(1, "Sending " + str(i) +": "+ str(self.send_numbers[i]))
            try:
                sleep(5)
                self.send_whatsapp_msg(i, self.send_numbers[i], self.send_msgs[i])

            except Exception:
                sleep(10)
                self.is_connected()
        self.sendStatus(2, "=== End ===")
        try:
            self.driver.quit()
            self.driver.close()
        except:
            print("err")

    def sendStatus(self, type, text):
        self.output.emit(type, text)
        return
        
    def element_presence(self, by, xpath, time):
        element_present = EC.presence_of_element_located((By.XPATH, xpath))
        WebDriverWait(self.driver, time).until(element_present)

    def is_connected(self):
        try:
            socket.create_connection(("www.google.com", 80))
            return True
        except:
            self.is_connected()

    def send_whatsapp_msg(self, index, phone_no, msgs):
        if(phone_no == None or phone_no == "" or phone_no == 0):
            self.sendStatus(1, "Empty Phone Number")
            return
        sleep(5)
        self.driver.get("https://web.whatsapp.com/send?phone={}&source=&data=#".format(phone_no))

        try:
            sleep(7)
            self.element_presence(By.XPATH, '//*[@id="main"]/footer/div[1]/div[2]/div/div[2]', 30)
            txt_box = self.driver.find_element(By.XPATH, '//*[@id="main"]/footer/div[1]/div[2]/div/div[2]')
            for i in range(len(msgs)):
                if(msgs[i] == "<br>"):
                    txt_box.send_keys(Keys.SHIFT, Keys.ENTER)
                else:
                    txt_box.send_keys(msgs[i])

            if self.send_image != "" and self.send_image != None:
                self.element_presence(By.XPATH, '//*[@id="main"]/header/div[3]/div/div[2]/div', 30)
                attachbox = self.driver.find_element(By.XPATH, '//*[@id="main"]/header/div[3]/div/div[2]/div')
                attachbox.click()
                sleep(0.3)

                self.element_presence(By.XPATH, '//*[@id="main"]/header/div[3]/div/div[2]/span/div/div/ul/li[1]/button/input', 30)
                imageBtn = self.driver.find_element(By.XPATH, '//*[@id="main"]/header/div[3]/div/div[2]/span/div/div/ul/li[1]/button/input')
                imageBtn.send_keys(self.send_image)

                self.element_presence(By.XPATH, '//*[@id="app"]/div/div/div[2]/div[2]/span/div/span/div/div/div[2]/span/div/div', 30)
                send_btn = self.driver.find_element(By.XPATH, '//*[@id="app"]/div/div/div[2]/div[2]/span/div/span/div/div/div[2]/span/div/div')
                send_btn.click()
                sleep(5)
            else:
                txt_box.send_keys("\n")

            self.sendStatus(1, "Send: "+phone_no)
        except Exception:
            self.sendStatus(1, "Invalid Phone Number: " + str(phone_no))
            self.sendStatus(0, str(index))

