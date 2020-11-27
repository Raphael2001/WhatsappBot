from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from time import sleep
import xlrd
import NoWhatsApp
import pandas as pd
import sys

## The Phone number at the file File.xlsx should start witout the 0

class WhatsAppBot:
    sys.setrecursionlimit(500)

    def __init__(self, filepath, root):
        self.root = root
        self.driver = webdriver.Chrome()
        self.nowhatsapplist = []
        self.worksheet = self.openxlsx(filepath)
        self.total_rows = len(self.worksheet)
        self.name = ''
        self.phone_number = ''
        self.row_cursor = 0
        self.column_names = ['Name', 'PhoneNumber']
        self.df = pd.DataFrame(columns=self.column_names)

    def main(self):
        while self.total_rows != 0:
            print("-------------------starting--------------------")
            self.print_worksheet()
            self.row_cursor = 0
            for self.row_cursor in range(0, self.total_rows):
                self.phone_number = '972' + self.getphonenumber()
                self.name = self.worksheet.iat[self.row_cursor, 0]

                print(self.phone_number + " " + self.name)

                try:
                    self.driver.get("https://web.whatsapp.com/send?phone=" + self.phone_number + " .")
                    self.driver.execute_script("window.onbeforeunload = function() {};")

                    # msg = "Hello " + self.name + ", how are you today? Tell me about your self!"
                    msg = "בלאק פריידי הגיע גם אלינו! כל הגולשים באתר נהנים מהנחות שוות לכבוד חג השופינג הגדול אבל בגלל שהצטרפת לרשימת הדיוור שלנו מגיע לך משהו מיוחד - 30% הנחה על כל המוצרים באתר! קוד הקופון: BLACK30 לרכישה: https://bit.ly/blackfriday-homeot להסרה יש להשיב את המילה ״הסר״ "

                    flag = False
                    while flag is False:
                        try:
                            sleep(20)
                            self.sendmsg()
                            flag = True
                            print("Message sent")
                            self.remove_from_worksheet()

                        except Exception:
                            try:
                                self.okbuttonclick()
                            except Exception:
                                break

                except Exception:
                    print("Message was not sent")
                    self.okbuttonclick()

        self.root.destroy()
        self.df.to_excel("MessageSent.xlsx")
        NoWhatsApp.Create_CSV(self.nowhatsapplist)
        print("Files Modified")

    #
    # def sendmsg(self, msg):
    #     msg_box = self.driver.find_element_by_xpath('//*[@id="main"]/footer/div[1]/div[2]/div/div[2]')
    #     msg_box.send_keys(msg + Keys.ENTER)

    def sendmsg(self):
        msg_box = self.driver.find_element_by_xpath('//*[@id="main"]/footer/div[1]/div[2]/div/div[2]')
        msg_box.send_keys(" ")

    def print_worksheet(self):
        for row_number in range(0, self.total_rows):
            self.name = self.worksheet.iat[row_number, 0]
            print(self.name)

    def remove_from_worksheet(self):
        self.worksheet.drop(self.worksheet.index[self.row_cursor])
        new_row = {'Name': self.name, 'PhoneNumber': self.phone_number}
        self.df = self.df.append(new_row, ignore_index=True)
        self.row_cursor = self.row_cursor - 1
        self.total_rows = self.total_rows - 1
        print("Total Rows: " + str(self.total_rows))

    def openxlsx(self, filepath):
        df = pd.read_excel(filepath)
        return df

    def getphonenumber(self):
        phone_number = self.worksheet.iat[self.row_cursor, 1]
        phone_number = str(phone_number)
        # phone_number = phone_number[:-2]
        return phone_number

    def okbuttonclick(self):
        ok_button_path = '//*[@id="app"]/div/span[2]/div/span/div/div/div/div/div/div[2]/div'
        ok_button = self.driver.find_element_by_xpath(ok_button_path)
        ok_button.click()
        self.phone_number = '0' + self.getphonenumber()
        print("No WhatsApp")
        self.remove_from_worksheet
        self.nowhatsapplist.append([self.name, self.phone_number])

    def loopfuntcion(self, msg):
        try:
            self.sendmsg(msg)
        except Exception:
            try:
                self.okbuttonclick()
            except Exception:
                raise Exception("ERROR")
