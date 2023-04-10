import requests
import json
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import openpyxl
from datetime import date
import math
import os
from dotenv import load_dotenv, find_dotenv

load_dotenv(find_dotenv())

'''В файл Task_from_PVB.xlsx проставить нужные артикулы для поиска в первый столбец. Шапку с заголовками оставить. 
Всю остальную информацию удалить(если она есть) 

Артикулы вставлять как специальная вставка

'''
user_name = os.getenv('user_name')
password_token = os.getenv('password_token')

user_agent_value = 'Mozilla/5.0 (Windows NT 6.1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/107.0.0.0 Safari/537.36'
headers = {'User-Agent': user_agent_value}
data = {
    'UserName': user_name,
    'Password': password_token,
    'RememberMe': 'False'
}


def main():
    DataToExcel(GetData(autorization(user_agent_value, headers, data), ProcessedString(GetTaskArticles()), headers))
    SendMesageToMail()


def autorization(user_agent_value, headers, data):
    session = requests.Session()
    autorization = session.post(url="https://www.bilopt.ru/Account/LogOn", headers=headers, data=data)
    return session


def GetTaskArticles():
    list_of_task_articles = []
    book = openpyxl.open("Task_from_PVB.xlsx", read_only=False, data_only=True)
    sheet = book.active
    for row in range(2, (sheet.max_row + 1)):
        list_of_task_articles.append(sheet[row][0].value)
    book.close()
    return list_of_task_articles


def ProcessedString(non_processed_list):
    new_string = ''
    processed_list = []
    for article in non_processed_list:
        new_string = ''
        for letters in str(article):
            if letters.isalpha() or letters.isdigit():
                new_string = new_string + letters
            else:
                pass
        processed_list.append(new_string)
    return processed_list


def GetData(session, processed_list, headers):
    brand = str(input('Введите название бренда,например Denso (Регистр имеет значение, вводить как на сайте BilOpt): '))
    total_list = []
    qty_list_total = []
    for articles in processed_list:
        info_list = []
        url = f"https://www.bilopt.ru/Search/GetFindHeaders?productId=&number={articles}"
        response = session.get(url=url, headers=headers)
        user_friendly_json = json.loads(response.text)
        if len(user_friendly_json["ProductLists"]) == 0 or user_friendly_json["ProductLists"][0]['Groups'][0][
            'Manufacturers'] == None:
            info_list.append('Нет данных на сайте')
            info_list.append('Нет данных на сайте')
            info_list.append(0)
            info_list.append(0)
            info_list.append(0)
            info_list.append(0)
            info_list.append(0)
            info_list.append(0)
            total_list.append(info_list)
        for x in range(len(user_friendly_json['ProductLists'])):
            if user_friendly_json["ProductLists"][x]['Groups'][0]['Manufacturers'] != None:
                for j in range(len(user_friendly_json["ProductLists"][x]['Groups'][0]['Manufacturers'])):
                    if user_friendly_json["ProductLists"][x]['Groups'][0]['Manufacturers'] == None:
                        continue
                    current_brand = user_friendly_json["ProductLists"][x]['Groups'][0]['Manufacturers'][j]
                    if current_brand != brand:
                        pass
                    else:
                        product_id = user_friendly_json["ProductLists"][x]['Groups'][0]['Products'][0]['ProductId']
                        info_list.append(user_friendly_json["ProductLists"][x]['Groups'][0]['Manufacturers'][j])
                        info_list.append(
                            user_friendly_json["ProductLists"][x]['Groups'][0]['Products'][0]["ProductNumber"])
                        info_list.append(math.ceil(
                            user_friendly_json["ProductLists"][x]['Groups'][0]['Products'][0]["MinimalPrice"]))
                        info_list.append(
                            user_friendly_json["ProductLists"][x]['Groups'][0]['Products'][0]["MaximumPrice"])
                        url_qty_inf = f'https://www.bilopt.ru/Search/GetFindOffers?productId=&number={articles}&city=&selectedProductId={product_id}'
                        response_qty = session.get(url=url_qty_inf, headers=headers)
                        user_friendly_json_qty = json.loads(response_qty.text)
                        for x in range(len(user_friendly_json_qty['Items'])):
                            if user_friendly_json_qty['Items'][x]['Quantity'] == '' or \
                                    user_friendly_json_qty['Items'][x]['Quantity'] == None:
                                pass
                            else:
                                qty_list_total.append(int(user_friendly_json_qty['Items'][x]['Quantity']))
                        if len(qty_list_total) == 0:
                            max_qty = 0
                            min_qty = 0
                            average_qty = 0
                            total_qty = 0
                        else:
                            max_qty = max(qty_list_total)
                            min_qty = min(qty_list_total)
                            average_qty = round(sum(qty_list_total) / len(qty_list_total))
                            total_qty = sum(qty_list_total)
                        info_list.append(max_qty)
                        info_list.append(min_qty)
                        info_list.append(average_qty)
                        info_list.append(total_qty)
                        total_list.append(info_list)
                        qty_list_total = []
    return total_list


def DataToExcel(total_list):
    book = openpyxl.open("Task_from_PVB.xlsx", read_only=False, data_only=True)
    sheet = book.active
    for i in range(len(total_list)):
        for info in range(len(total_list[i])):
            sheet.cell(row=i + 2, column=info + 1)
            sheet[i + 2][info + 1].value = total_list[i][info]
    book.save("Task_from_PVB.xlsx")
    book.close()


def SendMesageToMail():
    email_pass = os.getenv('email_pass')
    current_date = date.today()
    server = smtplib.SMTP('smtp.gmail.com', 587)
    sender = 'tableopposite@gmail.com'
    send_to = 'purchase2@bilight.biz'
    password = email_pass
    server.starttls()
    message = f'Цены БИЛОПТ на {current_date}'
    msg = MIMEMultipart()
    msg['From'] = sender
    msg['To'] = send_to
    msg['Subject'] = f'Цены БИЛОПТ на {current_date}'
    msg.attach(MIMEText(message))
    try:
        file = open('Task_from_PVB.xlsx', 'rb')
        part = MIMEBase('application', 'Task_from_PVB.xlsx')
        part.set_payload(file.read())
        file.close()
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', 'attachment', filename='Task_from_PVB.xlsx')
        msg.attach(part)
        server.login(sender, password)
        server.sendmail(sender, send_to, msg.as_string())
        return print('Письмо отправленно успешно')
    except Exception as _ex:
        return f'{_ex}\n Проверьте ваш логин или пароль!'
    server.quit()


if __name__ == "__main__":
    main()
