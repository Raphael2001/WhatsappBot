from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from time import sleep
import pandas as pd
import sys
from client import *

# The Phone number at the file File.xlsx should start witout the 0


class WhatsAppBot:

    def __init__(self, filepath, root):
        self.root = root
        self.driver = webdriver.Chrome()
        self.worksheet = self.openxlsx(filepath)
        self.total_rows = len(self.worksheet)
        self.Clients = Clients()
        self.client = Client("", "")
        self.msg = []

    def main(self):

        print("-------------------starting--------------------")
        while self.total_rows != 0:
            # for row_cursor in range(0, self.total_rows):
            phone_number = '972' + self.getphonenumber()
            name = self.worksheet.iat[0, 0]
            self.client = Client(name, phone_number)

            print(self.client.phone_number + " " + self.client.name)

            try:
                self.driver.get("https://web.whatsapp.com/send?phone=" + self.client.phone_number + " .")
                self.driver.execute_script("window.onbeforeunload = function() {};")
                self.clear_list()
                self.append_to_list()
                # msg = "Hello " + self.client.name + ", how are you today? Tell me about your self!"

                flag = False
                while flag is False:
                    try:
                        sleep(20)
                        self.sendmsg()
                        print("Message sent")
                        self.Clients.add_to_sent_message_list(self.client)
                        self.remove_from_worksheet()
                        flag = True
                    except Exception:
                        try:
                            self.okbuttonclick()
                        except Exception:
                            break

            except Exception:
                print("Message was not sent")
                self.okbuttonclick()

        self.root.destroy()
        self.Clients.sent_message_to_excel()
        self.Clients.no_whatsapp_to_excel()
        print("Files Modified")

    #
    # def sendmsg(self, msg):
    #     msg_box = self.driver.find_element_by_xpath('//*[@id="main"]/footer/div[1]/div[2]/div/div[2]')
    #     msg_box.send_keys(msg + Keys.ENTER)
    #
#     def sendmsg(self):
#         msg_box = self.driver.find_element_by_xpath('//*[@id="main"]/footer/div[1]/div[2]/div/div[2]')
#         msg_box.send_keys(" ")

    def sendmsg(self):
        msg_box = self.driver.find_element_by_xpath('//*[@id="main"]/footer/div[1]/div[2]/div/div[2]')
        for i in range(0, len(self.msg)):
            msg_box.send_keys(self.msg[i])
            self.line()
        msg_box.send_keys(Keys.ENTER)

    def print_worksheet(self):
        for row_number in range(0, self.total_rows):
            name = self.worksheet.iat[row_number, 0]
            print(str(row_number) + ". " + name)

    def remove_from_worksheet(self):
        self.worksheet = self.worksheet.drop(self.worksheet.index[0])
        self.total_rows = len(self.worksheet)
        print("Total Rows: " + str(self.total_rows))

    def openxlsx(self, filepath):
        df = pd.read_excel(filepath)
        return df

    def getphonenumber(self):
        phone_number = self.worksheet.iat[0, 1]
        phone_number = str(phone_number)
        # phone_number = phone_number[:-2]
        return phone_number

    def okbuttonclick(self):
        ok_button_path = '//*[@id="app"]/div/span[2]/div/span/div/div/div/div/div/div[2]/div'
        ok_button = self.driver.find_element_by_xpath(ok_button_path)
        ok_button.click()
        self.client.phone_number = '0' + self.getphonenumber()
        print("No WhatsApp")
        self.Clients.add_to_no_whatsapp_list(self.client)
        self.remove_from_worksheet()

    def line(self):
        combine_keys = ActionChains(self.driver)
        combine_keys.key_down(Keys.SHIFT).send_keys(Keys.ENTER).key_up(Keys.SHIFT).perform()

    def clear_list(self):
        self.msg = []

    def append_to_list(self):
        #write you massage here.
        #each line is new line at the massage
        self.msg.append("שלום " + self.client.name + ",")
        self.msg.append("מה שלומך?")
        self.msg.append("אני ז'רמי מחברת הומאוטריט-לאב.")
        self.msg.append("אנחנו שמחים שהצטרפת לאחרונה לקהל הלקוחות שלנו.")
        self.msg.append("אנחנו זמינים לכל שאלה ועניין דרך הווצאפ או בטל' 03-5560838")
        self.msg.append("בברכת בריאות רבה")
