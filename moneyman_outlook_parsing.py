import win32com.client
import win32com
import pandas as pd
from outlook import *

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
accounts = win32com.client.Dispatch("Outlook.Application").Session.Accounts

f = outlook_get_folder_from_name(outlook, 'МаниМен платежи')

date_start = pd.to_datetime('2019-04-01', format='%Y-%m-%d')
date_end = pd.to_datetime('2019-05-15', format='%Y-%m-%d')

msg_list = extract_msg_by_dates(outlook, 'МаниМен платежи', date_start, date_end)

df_payments = pd.DataFrame()
start = msg_list[0][0]
end = start + len(msg_list)
errors = []

for msg in msg_list:
    if msg[0] % 50 == 0:
        print('Обрабатываю {} из {}'.format(msg[0] - start, end - start))
    try:
        df_pay = None
        dfs = pd.read_html(msg[2])
        for df in dfs:
            if df.iloc[0, 0] == 'ID клиента':
                df_pay = df
        if df_pay is not None:
            df_pay.set_axis(labels=df_pay.iloc[0], axis='columns', inplace=True)
            df_pay.drop(0, inplace=True)
            df_pay.set_axis(labels=[msg[0]], axis='index', inplace=True)
            df_pay.loc[msg[0], 'inbox_date'] = msg[1]
            df_payments = df_payments.append(df_pay, sort=False)
    except Exception as e:
        errors.append([msg[0], msg[1], e])

df_errors = pd.DataFrame(errors)

writer = pd.ExcelWriter('moneyman_parse.xlsx', engine='xlsxwriter')
df_errors.to_excel(writer, sheet_name='errors')
df_payments.to_excel(writer, sheet_name='payments')

writer.save()
