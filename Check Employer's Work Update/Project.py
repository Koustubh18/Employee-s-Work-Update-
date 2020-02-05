import pandas as pd
import os
import time
from datetime import datetime, timedelta
import smtplib
import config


def read_data(file):
    base = ".xlsx"
    df = pd.read_excel(file)
    file_list = df['Files'].tolist()
    file_list = [i + base for i in file_list]
    return file_list


def search_files(file_list):
    directory = "C:/Users/Pravin/Desktop/Intern/Second/Search_folder"  # Directory Where we have to apply search operation
    dir_files = os.listdir(directory)
    search_result = {}
    time_result = {}

    for file in file_list:
        if file in dir_files:
            search_result[file] = 1
            time_result[file] = datetime.fromtimestamp(os.path.getmtime(f"{directory}/{file}"))
        else:
            search_result[file] = 0
            time_result[file] = datetime.fromtimestamp(0)

    return search_result, time_result


def send_email(Sender_email, Sender_email_pass, Receiver_email, message):

    try:
        s = smtplib.SMTP('smtp.gmail.com', 587)
        s.ehlo()
        s.starttls()
        s.login(Sender_email, Sender_email_pass)
        body = f"Subject:Task\n\n{message}"
        s.sendmail(Sender_email, Receiver_email, body)
        s.quit()
        print("Successfully Sent")

    except:
        print("Email failed to sent")


def main(file_list):
    present_search_result, present_time_result = search_files(file_list)
    past_search_result, past_time_result = present_search_result, present_time_result
    check_up = datetime.now()

    while True:

        print(f"Check Up Date:{check_up.date()}")

        boundry_for = datetime(datetime.now().year, datetime.now().month, 1)
        boundry_mid = datetime(datetime.now().year, datetime.now().month, 8)
        boundry_back = datetime(datetime.now().year, datetime.now().month, 15)

        if check_up.date() == datetime.now().date() and check_up.date() <= boundry_back.date():
            missing = [k for k, v in present_search_result.items() if v == 0]
            updating = [k for k, v in present_time_result.items() if v != datetime.fromtimestamp(0) if v.date() < boundry_for.date()]
            # print(missing, updating)

            if len(missing) == 0 and len(updating) == 0:
                if check_up.month + 1 <= 12:
                    check_up = datetime(datetime.now().year, (datetime.now().month) + 1, 1)
                else:
                    check_up = datetime((datetime.now().year) + 1, 1, 1)
            else:
                if check_up.date() <= boundry_mid.date():
                    message = f"Hey I need you to update files\nMissing:{missing}\nNot Updated{updating}"
                    print(f"Hey I need you to update files\nMissing:{missing}\nNot Updated{updating}")
                else:
                    message = f"Urgent need to update\nMissing:{missing}\nNot Updated{updating}"
                    print(f"Urgent need to update\nMissing:{missing}\nNot Updated{updating}")

                send_email(config.email, config.passward, "koustubhlearning@gmail.com", message)
                check_up = check_up + timedelta(hours=24)

        c = 0
        for i in file_list:
            if present_search_result[i] != past_search_result[i]:
                c += 1
                if present_search_result[i] == 1:
                    message = f"File Added:{i} Time:{present_time_result[i]}"
                    print(f"File Added:{i} Time:{present_time_result[i]}")

                else:
                    message = f'Removed File:{i} Time:{datetime.now()}'
                    print(f'Removed File:{i} Time:{datetime.now()}')

                send_email(config.email, config.passward, "koustubhlearning@gmail.com", message)

            else:
                if present_time_result[i] > past_time_result[i]:
                    c += 1
                    message = f"File Updated:{i} Time:{present_time_result[i]}"
                    print(f"File Updated:{i} Time:{present_time_result[i]}")
                    send_email(config.email, config.passward, "koustubhlearning@gmail.com", message)

        if c == 0:
            print("No change")

        past_search_result, past_time_result = present_search_result, present_time_result
        present_search_result, present_time_result = search_files(file_list)

        time.sleep(2)


file_list = read_data('data.xlsx')  # File Name To extract file names
# main(file_list)
