from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from time import sleep
import xlrd
import NoWhatsApp

## The Phone number at the file File.xlsx should start witout the 0


class WhatsAppBot:
    def __init__(self, filepath):
        self.driver = webdriver.Chrome()
        self.nowhatsapplist = []
        self.worksheet = self.openxlsx(filepath)
        self.total_rows = self.worksheet.nrows
        self.number = ''
        self.phone_number = ''

    def main(self):

        for self.row_curser in range(1, self.total_rows):
            self.phone_number = '972' + self.getphonenumber(self.row_curser, self.worksheet)
            self.name = self.worksheet.cell(self.row_curser, 0).value

            print(self.phone_number + " " + self.name)

            try:
                self.driver.get("https://web.whatsapp.com/send?phone=" + self.phone_number + " .")
                self.driver.execute_script("window.onbeforeunload = function() {};")

                msg = "Hello " + self.name + ", how are you today? Tell me about your self!"

                flag = False
                while flag is False:
                    try:
                        sleep(15)
                        self.sendmsg(msg)
                        flag = True
                    except Exception:
                        try:
                            self.okbuttonclick()
                        except Exception:
                            break

            except Exception:
                self.okbuttonclick()

        NoWhatsApp.Create_CSV(self.nowhatsapplist)

    def sendmsg(self, msg):
        msg_box = self.driver.find_element_by_xpath('//*[@id="main"]/footer/div[1]/div[2]/div/div[2]')
        msg_box.send_keys(msg + Keys.ENTER)

    def openxlsx(self, filepath):
        workbook = xlrd.open_workbook(filepath)
        worksheet = workbook.sheet_by_name("Sheet1")
        return worksheet

    def getphonenumber(self, row_curser, worksheet):
        phone_number = worksheet.cell(row_curser, 1).value
        phone_number = str(phone_number)
        phone_number = phone_number[:-2]
        return phone_number

    def okbuttonclick(self):
        ok_button_path = '//*[@id="app"]/div/span[2]/div/span/div/div/div/div/div/div[2]/div'
        ok_button = self.driver.find_element_by_xpath(ok_button_path)
        ok_button.click()
        self.phone_number = '0' + self.getphonenumber(self.row_curser, self.worksheet)
        print("No WhatsApp")
        self.nowhatsapplist.append([self.name, self.phone_number])

    def loopfuntcion(self, msg):
        try:
            self.sendmsg(msg)
        except Exception:
            try:
                self.okbuttonclick()
            except Exception:
                raise Exception("ERROR")
