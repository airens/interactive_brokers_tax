#!/usr/bin/env python
# coding: utf-8

# In[1]:


import pandas as pd
import numpy as np
import platform
import os
import sys
from glob import glob
from datetime import datetime
import requests
import yfinance as yf
from docxtpl import DocxTemplate
from collections import namedtuple

dirname = "ibdata"

sections = [
    "Deposits & Withdrawals",
    "Trades",
    "Fees",
    "Dividends",
    "Withholding Tax",
    "Change in Dividend Accruals",
    "Interest",
]
currencies = {
    "RUB": [],
    "USD": ["01235", "RUB=X", None],
    "EUR": ["01239", "EURRUB=X", None],
    "CNH": ["01375", "CNHRUB=X", None]
}

StartDate = "20.03.2019"

Year = int(input("Введите год отчета: "))


# In[2]:


div_res = None
div_accurals_res = None
cashflow_res = None
interest_res = None
trades_res = None
interest_res = None
fees_res = None
div_sum = 0
div_tax_paid_rub_sum = 0
div_tax_full_rub_sum = 0
div_tax_rest_sum = 0
div_tax_need_pay_final_sum = 0
div_accurals_sum = 0
div_accurals_tax_paid_rub_sum = 0
div_accurals_tax_full_rub_sum = 0
div_accurals_tax_rest_sum = 0
div_final_tax_rest_sum = 0
div_final_sum = 0
div_tax_paid_final_sum = 0
income_rub_sum = 0
deduction_sum = 0
fees_rub_sum = 0
interest_rub_sum = 0
interest_rest_sum = 0
cashflow_rub_sum = 0
cashflow_usd_sum = 0
cashflow_eur_sum = 0


# In[3]:


def get_crs_tables():
    for currency,data in currencies.items():
        if not len(data):  # rub
            continue
        print(f"Получение таблицы курса {currency}...")
        Format1 = "%d.%m.%Y"
        Format2 = "%m.%d.%Y"
        From = datetime.strptime(StartDate,Format1)
        To = datetime.today()
        url = f"https://cbr.ru/Queries/UniDbQuery/DownloadExcel/98956?Posted=True&mode=1&VAL_NM_RQ=R{data[0]}&From="
        url += f"{From.strftime(Format1)}&To={To.strftime(Format1)}"
        url += f"&FromDate={From.strftime(Format2).replace('.', '%2F')}&ToDate={To.strftime(Format2).replace('.', '%2F')}"
        response = requests.get(url)
        df = pd.read_excel(response.content).rename(columns={"data": "date", "curs": "val"})
        assert df.shape[0] > 0, f"Не удалось загрузить таблицу курсов {currency}!"
        currencies[currency][2] = df
        with open(f"{currency}.xlsx", "wb") as file:
            file.write(response.content)

get_crs_tables()


# In[4]:


def split_report():
    print("Разделение отчета на разделы...")
    fname = f"{Year}.csv"
    if not os.path.exists(fname):
        input(f"Не найден файл отчета за {Year}г. ({fname})")
        sys.exit()
    with open(fname, encoding="utf8") as file:
        if not os.path.exists(dirname):
            os.mkdir(dirname)
        for old_file in glob(os.path.join(dirname, f"{Year}*.csv")):
            os.remove(old_file)
            #print("replaced ", old_file)
        out_file = None
        line = file.readline()
        while line:
            section, header, column, *_ = line.split(',')
            section = section
            if section == "Trades" and (column =="Account" or "Forex" in line):  # Skip strategies and forex
                line = file.readline()
                continue            
            if header == "Header":
                if out_file:
                    out_file.close()
                    out_file = None
                out_fname = ""
                if section in sections:
                    out_fname = os.path.join(dirname, f"{Year}_{section}.csv")
                    if os.path.exists(out_fname):  # if second header in the same section - skip header
                        out_fname = out_fname.replace(".csv", f"{file.tell()}.csv")
                    out_file = open(out_fname, 'w')
                    assert out_file, f"Can't open file {out_fname}!"
            if out_file and section in sections:
                out_file.write(line)
            line = file.readline()
    if out_file:
        out_file.close()

split_report()


# In[5]:


def get_ticker_price(ticker: str):
    return float(yf.Ticker(ticker).history(period="1d").Close.median())


# In[ ]:





# In[6]:


def load_data():
    print("Чтение разделов отчета...")
    data = {}
    for fname in glob(os.path.join(dirname, f"*.csv")):
        if (int(fname.replace(dirname + os.path.sep, '').split('_')[0]) > Year):
            continue
        print(f"--{fname}")
        df = pd.read_csv(fname, thousands=',')
        if (df.shape[0] == 0):
            print("----пропуск (пустой файл)")
            continue
        section = df.iloc[0, 0]
        if section not in data:
            data[section] = df
        else:
            df.columns = data[section].columns
            data[section] = pd.concat([data[section], df], ignore_index=True)
    if "Deposits & Withdrawals" in data:
        cashflow = data["Deposits & Withdrawals"]
        cashflow.columns = [col.lower().strip() for col in cashflow]
        cashflow = cashflow.rename(columns={"settle date": "date"})
        cashflow = cashflow[cashflow.header == "Data"]
        cashflow = pd.DataFrame(cashflow[cashflow.currency.isin(currencies)])
        cashflow.date = pd.to_datetime(cashflow.date)
    else:
        cashflow = None
    if "Trades" in data:
        trades = data["Trades"]
        trades.columns = [col.lower().strip() for col in trades]
        trades = trades.rename(columns={"comm/fee": "fee", "date/time": "date", "t. price": "price", "comm in usd":"fee"})
        trades = trades[trades.header == "Data"]
        trades = trades[trades.fee < 0]
        trades.date = pd.to_datetime(trades.date)
    else:
        trades = None
    if "Fees" in data:
        comissions = data["Fees"]
        comissions.columns = [col.lower().strip() for col in comissions]
        comissions = comissions[comissions.header == "Data"]
        comissions = comissions[comissions.subtitle != "Total"]
        comissions.date = pd.to_datetime(comissions.date)
        comissions = comissions[comissions.date.dt.year == Year]
    else:
        comissions = None
    if "Interest" in data:
        interests =  data["Interest"]
        interests.columns = [col.lower().strip() for col in interests]
        interests = interests[interests.header == "Data"]
        interests = interests[interests.currency != "Total"]
        interests.date = pd.to_datetime(interests.date)
        interests = interests[interests.date.dt.year == Year]
    else:
        interests = None
    if "Dividends" in data:
        div = data["Dividends"]
        div.columns = [col.lower().strip() for col in div]
        div = pd.DataFrame(div[div.currency.isin(currencies)])
        div.date = pd.to_datetime(div.date)
        div = pd.DataFrame(div[div.date.dt.year == Year])
    else:
        div = None
    if div is not None and "Withholding Tax" in data:
        div_tax = data["Withholding Tax"]
        div_tax.columns = [col.lower().strip() for col in div_tax]
        div_tax = pd.DataFrame(div_tax[div_tax.currency.isin(currencies)])
        div_tax.date = pd.to_datetime(div_tax.date)
        div_tax = pd.DataFrame(div_tax[div_tax.date.dt.year == Year])
        div.description = [desc.split(" Cash Dividend")[0].replace(" (", "(") for desc in div.description]
        div_tax.description = [desc.split(" Cash Dividend")[0].replace(" (", "(") for desc in div_tax.description]
        div.rename(columns={"description": "ticker"}, inplace=True)
        div_tax.rename(columns={"description": "ticker"}, inplace=True)
        if div.shape[0] != div_tax.shape[0]:
            print("Размеры таблиц дивидендов и налогов по ним не совпадают. Попробуем исправить...")
            df = pd.DataFrame(columns=div_tax.columns)
            for index, row in div.iterrows():
                tax_row = div_tax[
                    (div_tax["date"] == row["date"])
                    & (div_tax["ticker"] == row["ticker"])
                    & (div_tax["currency"] == row["currency"])
                ]
                if not tax_row.empty:
                    if len(tax_row) == 1:
                        amount = tax_row["amount"].item()
                    else:
                        print("Данных за день для тикера больше чем предполагалось, они будут объединены")
                        print(tax_row.head())
                        amount = tax_row["amount"].sum()
                else:
                    amount = np.float64(0)

                df.loc[index] = pd.Series({
                    "withholding tax": "Withholding Tax",
                    "header": "Data",
                    "currency": row["currency"],
                    "date": row["date"],
                    "ticker": row["ticker"],
                    "amount": amount,
                    "code": np.nan,
                })

            div_tax = df
    else:
        div_tax = None
    if "Change in Dividend Accruals" in data:
        div_accurals = data["Change in Dividend Accruals"]
        div_accurals.columns = [col.lower().strip() for col in div_accurals]
        div_accurals = pd.DataFrame(div_accurals[div_accurals.currency.isin(currencies)])
        div_accurals.date = pd.to_datetime(div_accurals.date)
        div_accurals = pd.DataFrame(div_accurals[div_accurals.date.dt.year == Year])
    else:
        div_accurals = None
    return cashflow, trades, comissions, div, div_tax, div_accurals, interests

cashflow, trades, comissions, div, div_tax, div_accurals, interests = load_data()


# In[ ]:





# In[7]:


def trades_add():
    res = []
    if datetime.today().year == Year and input("Хотите добавить сделки купли\\продажи? (y\\n) ") == "y":
        print("Введите сделки в формате{тикер} {шт.} {валюта}. (шт.<0 - продажа, пустая строка для завершения, res - сначала)")
        i = 0
        while(1):
            s = input(f"{i+1}: ")
            spl = s.split(" ")
            if s == "res":
                res.clear()
                i = 0
                print("Сброс...")
                continue
            elif (len(spl) != 3):
                break
            else:
                ticker, cnt, cur = spl
                price = get_ticker_price(ticker)
                res.append({
                    #"trades": "Trades",
                    #"header": "Data",
                    #"datadiscriminator": "Order",
                    #"asset category": "Stocks",
                    "currency": cur.upper(),
                    "symbol": ticker.upper(),
                    "fee": -1.0,
                    "date": datetime.today(),
                    "quantity": float(cnt),
                    "price": float(price)
                })
                print("Куплено" if float(cnt)>0 else "Продано", abs(float(cnt)), ticker.upper(), "по цене", price, cur)
                i += 1
        print(f"Добавлено {len(res)} сделок")
    return res
if trades is not None:
    trades_extra = trades_add()
    for i, trade in enumerate(trades_extra):
        cur = trade["currency"]
        if cur in currencies:
            trades = trades.append(trade, ignore_index=True)
        else:
            print(f"Неизвестная валюта {cur}, сделка № {i+1} не используется")


# In[8]:


def get_currency(date, cur):
    assert cur in currencies, f"Неизвестная валюта {cur}!"
    if not len(currencies[cur]):
        return 1  # rub
    data = currencies[cur][2]
    diff = (data.date - date)
    indexmax = (diff[(diff <= pd.to_timedelta(0))].idxmax())
    assert float(data.iloc[[indexmax]].nominal) != 0.0, f"Номинал валюты не может быть 0"
    return float(data.iloc[[indexmax]].val) / float(data.iloc[[indexmax]].nominal)


# In[9]:


def cashflow_calc():
    print(f"Расчет таблицы переводов...")
    res = cashflow[["date", "currency", "amount"]].copy()
    res["type"] = ["Перевод на счет" if amount > 0 else "Снятие со счета" for amount in cashflow.amount]
    cashflow_rub_sum = res[res.currency == "RUB"].amount.sum().round(2)
    cashflow_usd_sum = res[res.currency == "USD"].amount.sum().round(2)
    cashflow_eur_sum = res[res.currency == "EUR"].amount.sum().round(2)
    return res, cashflow_rub_sum, cashflow_usd_sum, cashflow_eur_sum

if cashflow is not None:
    cashflow_res, cashflow_rub_sum, cashflow_usd_sum, cashflow_eur_sum = cashflow_calc()
    print("\ncashflow_res:")
    print(cashflow_res.head(2))
    print(cashflow_rub_sum, cashflow_usd_sum, cashflow_eur_sum)
    print("\n")
else:
    print("Нет данных по переводам")


# In[10]:


def div_calc():
    print(f"Расчет таблицы дивидендов...")
    res = pd.DataFrame()
    res["ticker"] = div["ticker"].values
    res["date"] = div["date"].values
    res["amount"] = div["amount"].values.round(2)
    res["currency"] = div["currency"].values
    if div_tax is None:
        print("Не найдена таблица удержанного налога с дивидендов. Налог на дивиденды будет 13%")
    res["tax_paid"] = -div_tax["amount"].values.astype(np.double).round(2) if div_tax is not None else 0
    res["cur_price"] = [get_currency(row.date, row.currency) for _, row in div.iterrows()]
    res["amount_rub"] = (res.amount*res.cur_price).round(2)
    res["tax_paid_rub"] = (res.tax_paid*res.cur_price).round(2)
    res["tax_full_rub"] = (res.amount_rub*13/100).round(2)
    res["tax_rest_rub"] = (res.tax_full_rub - res.tax_paid_rub).round(2).clip(lower=0)
    return res

if div is not None:
    div_res = div_calc()
    div_sum = round(div_res.amount_rub.sum(), 2)
    div_tax_paid_rub_sum = round(div_res.tax_paid_rub.sum(), 2)
    div_tax_full_rub_sum = round(div_res.tax_full_rub.sum(), 2)
    div_tax_rest_sum = round(div_res.tax_rest_rub.sum(), 2)
    print("\ndiv_res:")
    print(div_res.head(2))
    print("\n")
else:
    print("Нет данных по начисленным дивидендам")


# In[11]:


def div_accurals_calc():
    print(f"Расчет таблицы корректировки дивидендов...")
    res = pd.DataFrame()
    res["ticker"] = div_accurals['symbol']
    res["date"] = div_accurals["date"]
    res["amount"] = div_accurals["gross amount"].round(2)
    res["currency"] = div_accurals["currency"].values
    res["tax_paid"] = div_accurals["tax"].round(2)
    res["cur_price"] = [get_currency(row.date, row.currency) for _, row in div_accurals.iterrows()]
    res["amount_rub"] = (res.amount*res.cur_price).round(2)
    res["tax_paid_rub"] = (res.tax_paid*res.cur_price).round(2)
    res["tax_full_rub"] = (res.amount_rub*13/100).round(2)
    res["tax_rest_rub"] = (res.tax_full_rub - res.tax_paid_rub).round(2)
    return res

if div_accurals is not None:
    div_accurals_res = div_accurals_calc()
    div_accurals_sum = div_accurals_res.amount_rub.sum().round(2)
    div_accurals_tax_paid_rub_sum = div_accurals_res.tax_paid_rub.sum().round(2)
    div_accurals_tax_full_rub_sum = div_accurals_res.tax_full_rub.sum().round(2)
    div_accurals_tax_rest_sum = div_accurals_res.tax_rest_rub.sum().round(2)
    print("\ndiv_accurals_res:")
    print(div_accurals_res.head(2))
    print("\n")
else:
    print("Нет данных по изменениям в начисленнии дивидендов")


# In[12]:


div_final_tax_rest_sum = (div_tax_rest_sum + div_accurals_tax_rest_sum).round(2)
div_final_sum = (div_sum + div_accurals_sum).round(2)
div_tax_paid_final_sum = (div_tax_paid_rub_sum + div_accurals_tax_paid_rub_sum).round(2)
div_tax_need_pay_final_sum = (div_tax_rest_sum + div_accurals_tax_rest_sum).round(2)


# In[13]:


def fees_calc():
    print("Расчет таблицы комиссий...")
    fees = pd.DataFrame()
    fees["date"] = comissions.date
    fees["fee"] = comissions.amount*-1
    fees["currency"] = comissions["currency"].values
    fees["cur_price"] = [get_currency(row.date, row.currency) for _, row in comissions.iterrows()]
    fees["fee_rub"] = (fees.fee*fees.cur_price).round(2)
    return fees

if comissions is not None:
    fees_res = fees_calc()
    fees_rub_sum = fees_res.fee_rub.sum().round(2)
    print("\nfees_res:")
    print(fees_res.head(2))
    print("\n")
else:
    print("Нет данных по комиссиям")


# In[14]:


def trades_calc():
    print("Расчет таблицы сделок...")
    Asset = namedtuple("Asset", "date price fee currency")
    assets = {}
    rows = []
    for key, val in trades.sort_values(['date'], ascending=True).groupby("symbol"):
        fail = False
        if not key in assets:
            assets[key] = []
        for date, price, fee, quantity, currency in zip(val.date, val.price, val.fee, val.quantity, val.currency):
            if fail:
                break
            for _ in range(int(abs(quantity))):
                if fail:
                    break
                if quantity > 0:
                    assets[key].append(Asset(date, price, fee/abs(quantity), currency))
                elif quantity < 0:
                    if assets[key]:
                        buy_date, buy_price, buy_fee, buy_currency = assets[key].pop(0)
                        if date.year == Year:
                            #buy
                            rows.append(
                                {
                                    'ticker': key,
                                    'date': buy_date,
                                    'price': buy_price,
                                    'fee': buy_fee,
                                    'cnt': 1,
                                    'currency': buy_currency
                                }
                            )
                            #sell
                            rows.append(
                                {
                                    'ticker': key,
                                    'date': date,
                                    'price': price,
                                    'fee': fee/abs(quantity),
                                    'cnt': -1,
                                    'currency': currency
                                }
                            )
                    else:
                        print(f"Актив ({key}) продан в большем количестве, чем покупался. Операции SHORT не поддерживаются.")
                        rows = [row for row in rows if row["ticker"] != key]
                        fail = True
    if datetime.today().year == Year:
        print("Рассчет налоговых оптимизаций...")
        for key, val in assets.items():
            if key in currencies:
                continue
            price_today = get_ticker_price(key)
            res = 0
            cnt = 0
            for buy_date, buy_price, _, currency in val:
                result = -buy_price*get_currency(buy_date, currency) + price_today*get_currency(datetime.today(), currency)
                if result<0:
                    cnt += 1
                else:
                    break
                res += result
            if res < 0:
                print("--Можно продать", cnt, key, "и получить", abs(round(res, 2)), "р. бумажного убытка")
        print("\n")
    return pd.DataFrame(rows, columns=['ticker', 'date', 'price', 'fee', 'cnt', 'currency'])


if trades is not None:
    trades_res = trades_calc()
    if len(trades_res):
        trades_res = trades_res.groupby(['ticker', 'date', 'price', 'currency'], as_index=False)[['cnt', 'fee']].sum()
    trades_res["type"] = ["Покупка" if cnt > 0 else "Продажа" for cnt in trades_res.cnt]
    trades_res["price"] = trades_res.price.round(2)
    trades_res["fee"] = trades_res.fee.round(2)*-1
    trades_res["amount"] = (trades_res.price*trades_res.cnt*-1).round(2)
    trades_res["amount_buy"] = trades_res["amount"]
    trades_res["cur_price"] = [get_currency(row.date, row.currency) for _, row in trades_res.iterrows()]
    trades_res["amount_rub"] = (trades_res.amount*trades_res.cur_price).round(2)
    trades_res.loc[trades_res['amount_rub'] < 0, 'amount_rub'] = 0
    trades_res.loc[trades_res['amount_buy'] > 0, 'amount_buy'] = 0    
    trades_res["deduction"] = ((trades_res.amount_buy*-1+trades_res.fee)*trades_res.cur_price).round(2)
    trades_res["cnt"] = trades_res.cnt.abs()
    trades_res = trades_res.sort_values(["ticker", "type", "date"])
    trades_res.loc[trades_res.duplicated(subset="ticker"), "ticker"] = ""
    income_rub_sum = round(trades_res.amount_rub.sum(), 2)
    deduction_sum = round(trades_res.deduction.sum(), 2)
    print("\ntrades_res:")
    print(trades_res.head(2))
    print("\n")
else:
    print("Нет данных по сделкам")


# In[15]:


def interest_calc():
    print("Расчет таблицы по программе повышения доходности")
    interest_calc = pd.DataFrame()
    interest_calc["date"] = interests.date
    interest_calc["description"] = interests.description
    interest_calc["currency"] = interests.currency
    interest_calc["amount"] = interests.amount
    interest_calc["cur_price"] = [get_currency(row.date, row.currency) for _, row in interests.iterrows()]
    interest_calc["amount_rub"] = (interest_calc.amount * interest_calc.cur_price).round(2)
    interest_calc["rest"] = (interest_calc.amount * interest_calc.cur_price * 0.13).round(2)
    interest_calc = interest_calc.sort_values(['date'])
    return interest_calc

if interests is not None:
    interest_res = interest_calc()
    interest_rub_sum = interest_res.amount_rub.sum().round(2)
    interest_rest_sum = interest_res.rest.sum().round(2)
    print("\ninterest_res:")
    print(interest_res.head(2))
    print("\n")
else:
    print("Нет данных по начисленным на наличные процентам")


# In[ ]:





# In[17]:


Fname = f"Пояснительная записка {Year}.docx"
def create_doc():
    print("Формирование отчета...")
    doc = DocxTemplate("template.docx")
    context = {
        'start_date': StartDate,
        'year': Year,
        'tbl_div': div_res.to_dict(orient='records') if div_res is not None else {},
        'div_sum': div_sum,
        'div_tax_paid_rub_sum': div_tax_paid_rub_sum,
        'div_tax_full_rub_sum': div_tax_full_rub_sum,
        'div_tax_rest_sum': div_tax_rest_sum,
        'tbl_div_accurals': div_accurals_res.to_dict(orient='records') if div_accurals_res is not None else {},
        'div_accurals_sum': div_accurals_sum,
        'div_accurals_tax_paid_rub_sum': div_accurals_tax_paid_rub_sum,
        'div_accurals_tax_full_rub_sum': div_accurals_tax_full_rub_sum,
        'div_accurals_tax_rest_sum': div_accurals_tax_rest_sum,
        'div_final_tax_rest_sum': div_final_tax_rest_sum,
        'div_final_sum': div_final_sum,
        'div_tax_paid_final_sum': div_tax_paid_final_sum,
        'div_tax_need_pay_final_sum': div_tax_need_pay_final_sum,
        'tbl_cashflow': cashflow_res.to_dict(orient='records') if cashflow_res is not None else {},
        'tbl_trades': trades_res.to_dict(orient='records') if trades_res is not None else {},
        'tbl_interest': interest_res.to_dict(orient='records') if interest_res is not None else {},
        'interest_rub_sum': interest_rub_sum,
        'interest_rest_sum': interest_rest_sum,
        'income_rub_sum': income_rub_sum,
        'deduction_sum': deduction_sum,
        'tbl_fees': fees_res.to_dict(orient='records') if fees_res is not None else {},
        'fees_rub_sum': fees_rub_sum,
        'deduction_total_sum': (fees_rub_sum + deduction_sum).round(2)
    }
    doc.render(context)
    doc.save(Fname)
create_doc()


# In[18]:


if platform.system() == 'Windows':
    input("Готово. Нажмите Enter, чтобы продолжить")
    os.startfile(Fname, "open")
else:
    print("Готово")


# In[ ]:




