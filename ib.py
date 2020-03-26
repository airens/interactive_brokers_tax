#!/usr/bin/env python
# coding: utf-8

# In[13]:


import pandas as pd
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
    "Trades",
    "Fees",
    "Dividends",
    "Withholding Tax",
    "Change in Dividend Accruals",
]

StartDate = "20.03.2019"
Year = int(input("Введите год отчета: "))   


# In[14]:


def get_usd_table():
    print("Получение таблицы курса $...")
    Format1 = "%d.%m.%Y"
    Format2 = "%m.%d.%Y"
    From = datetime.strptime(StartDate,Format1)
    To = datetime.today()
    url = "https://cbr.ru/Queries/UniDbQuery/DownloadExcel/98956?Posted=True&mode=1&VAL_NM_RQ=R01235&From="
    url += f"{From.strftime(Format1)}&To={To.strftime(Format1)}"
    url += f"&FromDate={From.strftime(Format2).replace('.', '%2F')}&ToDate={To.strftime(Format2).replace('.', '%2F')}"
    response = requests.get(url)
    with open("usd.xlsx", "wb") as file:
        file.write(response.content)

get_usd_table()


# In[15]:


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
            section, header, *_ = line.split(',')
            section = section
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


# In[16]:


def get_ticker_price(ticker: str):
    return float(yf.Ticker(ticker).history(period="1d").Close.median())


# In[17]:


def trades_add():
    res = []
    if datetime.today().year == Year and input("Хотите добавить сделки купли\\продажи? (y\\n) ") == "y":
        print("Введите сделки в формате{тикер} {шт.}. (шт.<0 - продажа, пустая строка для завершения, res - сначала)")
        i = 0
        while(1):
            s = input(f"{i+1}: ")
            if s == "res":
                res.clear()
                i = 0
                print("Сброс...")
                continue
            elif (len(s) < 0 or not ' ' in s):
                break
            else:
                ticker, cnt = s.split(" ")
                price = get_ticker_price(ticker)
                res.append({
                    #"trades": "Trades",
                    #"header": "Data",
                    #"datadiscriminator": "Order",
                    #"asset category": "Stocks",
                    #"currency": "USD",
                    "symbol": ticker.upper(),
                    "fee": -1.0,
                    "date": datetime.today(),
                    "quantity": float(cnt),
                    "price": float(price)                    
                })
                print("Куплено" if float(cnt)>0 else "Продано", abs(float(cnt)), ticker.upper(), "по цене", price)
                i += 1
        print(f"Добавлено {len(res)} сделок")
    return res
trades_extra = trades_add()


# In[18]:


def load_data():
    print("Чтение разделов отчета...")
    data = {}
    for fname in glob(os.path.join(dirname, f"*.csv")):
        if (int(fname.replace(dirname + os.path.sep, '').split('_')[0]) > Year):
            continue
        print(f"--{fname}")
        df = pd.read_csv(fname, thousands=',')
        section = df.iloc[0, 0]
        if section not in data:
            data[section] = df
        else:
            df.columns = data[section].columns
            data[section] = data[section].append(df, ignore_index=True)
    trades = data["Trades"]
    trades.columns = [col.lower() for col in trades]
    trades = trades.rename(columns={"comm/fee": "fee", "date/time": "date", "t. price": "price"})    
    trades = trades[trades.header == "Data"]
    trades = trades[trades.fee < 0]
    trades.date = pd.to_datetime(trades.date)
    comissions = data["Fees"]
    comissions.columns = [col.lower() for col in comissions]
    comissions = comissions[comissions.header == "Data"]
    comissions = comissions[comissions.subtitle != "Total"]
    comissions.date = pd.to_datetime(comissions.date)
    comissions = comissions[comissions.date.dt.year == Year]
    div = data["Dividends"]
    div.columns = [col.lower() for col in div]
    div = pd.DataFrame(div[div.currency == "USD"])
    div.date = pd.to_datetime(div.date)
    div = pd.DataFrame(div[div.date.dt.year == Year])
    div_tax = data["Withholding Tax"]
    div_tax.columns = [col.lower() for col in div_tax]
    div_tax = pd.DataFrame(div_tax[div_tax.currency == "USD"])
    div_tax.date = pd.to_datetime(div_tax.date)    
    div_tax = pd.DataFrame(div_tax[div_tax.date.dt.year == Year])
    assert div.shape[0] == div_tax.shape[0], "Tax table must be the same size as dividents table!"
    div_accurals = data["Change in Dividend Accruals"]
    div_accurals.columns = [col.lower() for col in div_accurals]
    div_accurals = pd.DataFrame(div_accurals[div_accurals.currency == "USD"])
    div_accurals.date = pd.to_datetime(div_accurals.date)
    div_accurals = pd.DataFrame(div_accurals[div_accurals.date.dt.year == Year])
    usd = pd.read_excel("usd.xlsx").drop(columns=["nominal", "cdx"])
    usd = usd.rename(columns={"data": "date", "curs": "val"})
    if (len(trades_extra)):
        trades = trades.append(trades_extra)
        usd = usd.append([{"date": datetime.today().replace(hour=0, minute=0), "val": get_ticker_price("RUB=X")}], ignore_index=True)    
    return trades, comissions, div, div_tax, div_accurals, usd

trades, comissions, div, div_tax, div_accurals, usd = load_data()


# In[19]:


def get_usd(date):
    diff = (usd.date - date)
    indexmax = (diff[(diff <= pd.to_timedelta(0))].idxmax())
    return float(usd.iloc[[indexmax]].val)    


# In[20]:


def div_calc():
    print(f"Расчет таблицы дивидендов...")
    res = pd.DataFrame()
    res["ticker"] = [desc.split(" Cash Dividend")[0] for desc in div.description]
    res["date"] = div["date"].values
    res["amount"] = div["amount"].values.round(2)
    res["tax_paid"] = -div_tax["amount"].values.round(2)
    res["usd_price"] = [get_usd(d) for d in div.date]
    res["amount_rub"] = (res.amount*res.usd_price).round(2)
    res["tax_paid_rub"] = (res.tax_paid*res.usd_price).round(2)
    res["tax_full_rub"] = (res.amount_rub*13/100).round(2)
    res["tax_rest_rub"] = (res.tax_full_rub - res.tax_paid_rub).round(2)
    return res

div_res = div_calc()
div_sum = div_res.amount_rub.sum().round(2)
div_tax_paid_rub_sum = div_res.tax_paid_rub.sum().round(2)
div_tax_full_rub_sum = div_res.tax_full_rub.sum().round(2)
div_tax_rest_sum = div_res.tax_rest_rub.sum().round(2)
print("\ndiv_res:")
print(div_res.head(20))
print("\n")


# In[21]:


def div_accurals_calc():
    print(f"Расчет таблицы корректировки дивидендов...")
    res = pd.DataFrame()
    res["ticker"] = div_accurals['symbol']
    res["date"] = div_accurals["date"]
    res["amount"] = div_accurals["gross amount"].round(2)
    res["tax_paid"] = div_accurals["tax"].round(2)
    res["usd_price"] = [get_usd(d) for d in div_accurals.date]
    res["amount_rub"] = (res.amount*res.usd_price).round(2)
    res["tax_paid_rub"] = (res.tax_paid*res.usd_price).round(2)
    res["tax_full_rub"] = (res.amount_rub*13/100).round(2)
    res["tax_rest_rub"] = (res.tax_full_rub - res.tax_paid_rub).round(2)
    return res

div_accurals_res = div_accurals_calc()
div_accurals_sum = div_accurals_res.amount_rub.sum().round(2)
div_accurals_tax_paid_rub_sum = div_accurals_res.tax_paid_rub.sum().round(2)
div_accurals_tax_full_rub_sum = div_accurals_res.tax_full_rub.sum().round(2)
div_accurals_tax_rest_sum = div_accurals_res.tax_rest_rub.sum().round(2)
print("\ndiv_accurals_res:")
print(div_accurals_res.head(2))
print("\n")


# In[10]:


div_final_tax_rest_sum = div_tax_rest_sum.round(2) + div_accurals_tax_rest_sum.round(2)
div_final_sum = div_sum.round(2) + div_accurals_sum.round(2)
div_tax_paid_final_sum = div_tax_paid_rub_sum.round(2) + div_accurals_tax_paid_rub_sum.round(2)


# In[11]:


def fees_calc():
    print("Расчет таблицы комиссий...")
    fees = pd.DataFrame()
    fees["date"] = comissions.date
    fees["fee"] = comissions.amount*-1
    fees["usd_price"] = [get_usd(d) for d in comissions.date]
    fees["fee_rub"] = (fees.fee*fees.usd_price).round(2)
    return fees
        
fees_res = fees_calc()
fees_rub_sum = fees_res.fee_rub.sum().round(2)
print("\nfees_res:")
print(fees_res.head(2))
print("\n")


# In[31]:


def trades_calc():
    print("Расчет таблицы сделок...")
    Asset = namedtuple("Asset", "date price fee")
    assets = {}
    rows = []
    for key, val in trades.groupby("symbol"):
        if not key in assets:
            assets[key] = []
        for date, price, fee, quantity in zip(val.date, val.price, val.fee, val.quantity):
            for _ in range(int(abs(quantity))):
                if quantity > 0:
                    assets[key].append(Asset(date, price, fee))
                elif quantity < 0:
                    buy_date, buy_price, buy_fee = assets[key].pop(0)
                    if date.year == Year:
                        #buy
                        rows.append(
                            {
                                'ticker': key,
                                'date': buy_date,
                                'price': buy_price,
                                'fee': buy_fee,
                                'cnt': 1,
                            }
                        )
                        #sell
                        rows.append(
                            {
                                'ticker': key,
                                'date': date,
                                'price': price,
                                'fee': fee,
                                'cnt': -1,
                            }
                        )
    if datetime.today().year == Year:
        print("Рассчет налоговых оптимизаций...")
        usd_today = get_usd(datetime.today())
        for key, val in assets.items():
            if "USD" in key:
                continue
            price_today = get_ticker_price(key)
            res = 0
            cnt = 0
            for buy_date, buy_price, _ in val:
                result = -buy_price*get_usd(buy_date) + price_today*usd_today
                if result<0:
                    cnt += 1
                else:
                    break
                res += result
            if res < 0:
                print("--Можно продать", cnt, key, "и получить", abs(round(res, 2)), "р. бумажного убытка")
        print("\n")
    return pd.DataFrame(rows, columns=['ticker', 'date', 'price', 'fee', 'cnt'])
trades_res = trades_calc()
if len(trades_res):
    trades_res = trades_res.groupby(['ticker', 'date', 'price', 'fee'], as_index=False)['cnt'].sum()
trades_res["type"] = ["Покупка" if cnt > 0 else "Продажа" for cnt in trades_res.cnt]
trades_res["price"] = trades_res.price.round(2)
trades_res["fee"] = trades_res.fee.round(2)*-1
trades_res["amount"] = (trades_res.price*trades_res.cnt*-1 - trades_res.fee).round(2)
trades_res["usd_price"] = [get_usd(d) for d in trades_res.date]
trades_res["amount_rub"] = (trades_res.amount*trades_res.usd_price).round(2)
trades_res["cnt"] = trades_res.cnt.abs()
trades_res = trades_res.sort_values(["ticker", "type", "date"])
trades_res.loc[trades_res.duplicated(subset="ticker"), "ticker"] = ""
print("\ntrades_res:")
print(trades_res.head(2))
print("\n")


# In[ ]:


income_rub_sum = round(trades_res.amount_rub.sum(), 2)


# In[ ]:


Fname = f"Пояснительная записка {Year}.docx"
def create_doc():
    print("Формирование отчета...")
    doc = DocxTemplate("template.docx")
    context = {
        'start_date': StartDate,
        'year': Year,
        'tbl_div': div_res.to_dict(orient='records'),
        'div_sum': div_sum,
        'div_tax_paid_rub_sum': div_tax_paid_rub_sum,
        'div_tax_full_rub_sum': div_tax_full_rub_sum,
        'div_tax_rest_sum': div_tax_rest_sum,
        'tbl_div_accurals': div_accurals_res.to_dict(orient='records'),
        'div_accurals_sum': div_accurals_sum,
        'div_accurals_tax_paid_rub_sum': div_accurals_tax_paid_rub_sum,
        'div_accurals_tax_full_rub_sum': div_accurals_tax_full_rub_sum,
        'div_accurals_tax_rest_sum': div_accurals_tax_rest_sum,
        'div_final_tax_rest_sum': div_final_tax_rest_sum,
        'div_final_sum': div_final_sum,
        'div_tax_paid_final_sum': div_tax_paid_final_sum,
        'tbl_trades': trades_res.to_dict(orient='records'),
        'income_rub_sum': income_rub_sum,
        'tbl_fees': fees_res.to_dict(orient='records'),
        'fees_rub_sum': fees_rub_sum
    }
    doc.render(context)
    doc.save(Fname)
create_doc()


# In[ ]:


input("Готово.")
os.startfile(Fname, "open")


# In[ ]:




