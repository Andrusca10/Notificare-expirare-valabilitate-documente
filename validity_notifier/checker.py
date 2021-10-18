# Standard Library
import csv
import datetime
import time

# Third Party Libraries
import pandas
import yagmail


def read_and_check():
    df = pandas.read_excel('masini.xlsx')
    df.to_csv('masini.csv', index=None)
    list = []
    with open('masini.csv') as f:
        reader = csv.DictReader(f)
        day = datetime.date.today()
        for row in reader:
            for k, v in row.items():
                if k == 'Autorizatie_Austria':
                    due_date = datetime.datetime.strptime(v, '%Y-%m-%d').date()
                    if datetime.timedelta(days=21) >= due_date - day > datetime.timedelta(days=0):
                        list.append(
                            "Autorizatia pentru Austria a masinii {} expira la data de {}".format(row['Masina'], v))
                if k == 'Autorizatie_Olanda':
                    due_date = datetime.datetime.strptime(v, '%Y-%m-%d').date()
                    if datetime.timedelta(days=21) >= due_date - day > datetime.timedelta(days=0):
                        list.append(
                            "Autorizatia pentru Olanda a masinii {} expira la data de {}".format(row['Masina'], v))
                if k == 'Autorizatie_Franta':
                    due_date = datetime.datetime.strptime(v, '%Y-%m-%d').date()
                    if datetime.timedelta(days=21) >= due_date - day > datetime.timedelta(days=0):
                        list.append(
                            "Autorizatia pentru Franta a masinii {} expira la data de {}".format(row['Masina'], v))
                if k == 'Asigurare_RCA':
                    due_date = datetime.datetime.strptime(v, '%Y-%m-%d').date()
                    if datetime.timedelta(days=14) >= due_date - day > datetime.timedelta(days=0):
                        list.append(
                            "Asigurarea RCA pentru masina {} expira la data de {}".format(row['Masina'], v))
                if k == 'Asigurare_CASCO':
                    due_date = datetime.datetime.strptime(v, '%Y-%m-%d').date()
                    if datetime.timedelta(days=14) >= due_date - day > datetime.timedelta(days=0):
                        list.append(
                            "Asigurarea CASCO pentru masina {} expira la data de {}".format(row['Masina'], v))
                if k == 'Asigurare_CMR':
                    due_date = datetime.datetime.strptime(v, '%Y-%m-%d').date()
                    if datetime.timedelta(days=14) >= due_date - day > datetime.timedelta(days=0):
                        list.append(
                            "Asigurarea CMR pentru masina {} expira la data de {}".format(row['Masina'], v))
                if k == 'ITP':
                    due_date = datetime.datetime.strptime(v, '%Y-%m-%d').date()
                    if datetime.timedelta(days=14) >= due_date - day > datetime.timedelta(days=0):
                        list.append(
                            "ITP-ul pentru masina {} expira la data de {}".format(row['Masina'], v))
                if k == 'Inspectie_tahograf':
                    due_date = datetime.datetime.strptime(v, '%Y-%m-%d').date()
                    if datetime.timedelta(days=14) >= due_date - day > datetime.timedelta(days=0):
                        list.append(
                            "Inspectia tahograf pentru masina {} expira la data de {}".format(row['Masina'], v))
    with open("file.txt", "w") as output:
        for row in list:
            output.write(str(row) + '\n')


def send_email(receiver, subject, content, attach):
    try:
        # replace user and pass with your credentials
        yag = yagmail.SMTP(user='sender@gmail.com', password='***')
        yag.send(receiver, subject, content, attach)
        print("Email sent successfully")
    except:
        print("Error, email was not sent")


while True:
    read_and_check()
    send_email("receiver@yahoo.com", "Notificare expirare valabilitate documente",
              "Salut\n\nTe rog sa verifici atasamentul.\n\nO zi buna!", "file.txt")  # replace receiver@yahoo.com with your receiver
    time.sleep(86400)  # the verification is made one time per day 24h = 86400s
