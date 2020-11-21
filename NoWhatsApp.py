import csv

fields = ['Name', 'Phone Number']
filename = "NoWhatsApp.csv"


class Create_CSV():

    def __init__(self, list):
        rows = list


        with open(filename, 'w') as csvfile:
            csvwriter = csv.writer(csvfile)

            csvwriter.writerow(fields)

            csvwriter.writerows(rows)
