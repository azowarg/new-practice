import requests, smtplib, yagmail
import pandas as pd

ses = requests.session()
report = pd.DataFrame()
urls = ["https://www.moex.com/ru/derivatives/currency-rate.aspx?currency=USD_RUB", \
        "https://www.moex.com/ru/derivatives/currency-rate.aspx?currency=EUR_RUB"]
for url in urls:
    col = ["Дата", "Курс", "Изменения"]
    spreadsheet = pd.read_html(ses.get(url).text, decimal=',', thousands="*")[2]
    spreadsheet.columns = spreadsheet.columns.droplevel(0)
    col[-2:] = [" ".join([str, url[-7:]]) for str in col[-2:]]
    del spreadsheet['Время']
    spreadsheet.columns = col
    if 'Дата' not in report.columns:
        report = pd.concat([report, spreadsheet['Дата']], axis = 1)
    report = pd.concat([report, spreadsheet['Курс ' + url[-7:]], spreadsheet['Изменения ' + url[-7:]]], axis = 1)
ses.close()
report["EUR/USD"] = report['Курс EUR_RUB'] / report['Курс USD_RUB']
report.to_excel('report.xlsx', index = 0)
yag = yagmail.SMTP(user=input("Gmail account: "), password=input("Password: "), host='smtp.gmail.com')
yag.send(to=input("Report receiver: "),subject="Currency_parsing", contents='excel file', attachments='report.xlsx' )