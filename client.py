import pandas as pd


class Client:
    def __init__(self, name,  phonenumber):
        self.phone_number = phonenumber
        self.name = name


class Clients:
    def __init__(self):
        self.column_names = ['Name', 'PhoneNumber']
        self.sent_message = pd.DataFrame(columns=self.column_names)
        self.no_whatsapp = pd.DataFrame(columns=self.column_names)

    def add_to_sent_message_list(self, client):
        new_row = {'Name': client.name, 'PhoneNumber': client.phone_number}
        self.sent_message = self.sent_message.append(new_row, ignore_index=True)

    def add_to_no_whatsapp_list(self, client):
        new_row = {'Name': client.name, 'PhoneNumber': client.phone_number}
        self.no_whatsapp = self.no_whatsapp.append(new_row, ignore_index=True)

    def sent_message_to_excel(self):
        self.sent_message.to_excel("MessageSent.xlsx")

    def no_whatsapp_to_excel(self):
        self.no_whatsapp.to_excel("NoWhatsApp.xlsx")

