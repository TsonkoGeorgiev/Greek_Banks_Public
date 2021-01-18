import requests, openpyxl, datetime, statistics, time, schedule
from bs4 import BeautifulSoup


def get_date(text):
    match_index = text.find('as of')
    start_index = match_index + 6
    end_index = start_index + 11
    str_date = text[start_index:end_index] #string representation of date (e.g. Sep 04 2019)
    obj_date = datetime.datetime.strptime(str_date, '%b %d %Y') #object representation of date in custom format
    str_date = obj_date.strftime('%d/%m/%Y') #convert object representation into custom format
    return str_date


def get_price(text):
    match_index = text.find('Price (EUR)')
    start_index = match_index + 56
    end_index = text.find('<', start_index)
    price = float(text[start_index:end_index])
    return price


def update_excel_sheet(date, price, stock_code, measure_dict, sheet, max_row):

    #Date
    sheet.cell(row=max_row + 1, column=1).value = date

    #PX
    key = 'price'
    sheet.cell(row=max_row + 1, column=measure_dict.get(key)).value = price

    #PX Delta(%)
    previous_price = sheet.cell(row=max_row, column=measure_dict.get(key)).value
    px_delta = round((price/previous_price - 1)*100, 2)
    key = 'px_delta'
    sheet.cell(row=max_row + 1, column=measure_dict.get(key)).value = px_delta

    #1 Year Return
    key = 'price'
    price_250_days_back = sheet.cell(row=max_row - 250, column=measure_dict.get(key)).value
    return_1_year = round((price/price_250_days_back - 1)*100, 2)
    key = 'return_1_year'
    sheet.cell(row=max_row + 1, column=measure_dict.get(key)).value = return_1_year

    if stock_code == 'TPEIR':
        #SMA50
        key = 'price'
        price_list = []
        for i in range(1, 51):
            price_list.append(sheet.cell(row=max_row - 49 + i, column=measure_dict.get(key)).value)
        sma50 = statistics.mean(price_list)
        sheet.cell(row=max_row + 1, column=4).value = sma50

        #SMA50/PX
        sheet.cell(row=max_row + 1, column=5).value = sma50 / price

        #SMA50 Delta
        previous_sma_50 = sheet.cell(row=max_row, column=4).value
        sma_delta = round((sma50/previous_sma_50-1)*100, 2)
        sheet.cell(row=max_row + 1, column=7).value = sma_delta


def job():
    if run_today():
        banks_list = ['TPEIR', 'ETE', 'EUROB', 'ALPHA']
        file = 'GR Banks.xlsx'
        wb = openpyxl.load_workbook(file, read_only=False)
        sheet = wb['Main']
        max_row = sheet.max_row

        for code in banks_list:
            res = requests.get('https://markets.ft.com/data/equities/tearsheet/historical?s=' + code + ':ATH')
            if code == 'TPEIR':
                metric_dict = {     #use a dictionary with metric as key and Excel sheet column as value
                    'price': 2,
                    'px_delta': 3,
                    'return_1_year': 6,
                }
            elif code == 'ETE':
                metric_dict = {
                    'price': 8,
                    'px_delta': 9,
                    'return_1_year': 10,
                }
            elif code == 'EUROB':
                metric_dict = {
                    'price': 11,
                    'px_delta': 12,
                    'return_1_year': 13,
                }
            else:
                metric_dict = {
                    'price': 14,
                    'px_delta': 15,
                    'return_1_year': 16,
                }
            line = res.text
            as_of_date = get_date(line)
            price_today = get_price(line)
            update_excel_sheet(as_of_date, price_today, code, metric_dict, sheet, max_row)

        wb.save(file)
        wb.close()


def get_stock_exchange_holidays_calendar():
    url = 'https://athexgroup.gr/web/guest/market-alternative-holidays'
    page = requests.get(url, verify=False)
    content_parsed = BeautifulSoup(page.content, 'html.parser')
    holidays_html = content_parsed.find(id="holidays")
    holidays = holidays_html.text.strip().split('   ')
    holidays_no_header = []

    for holiday in holidays:
        if 'Market Holidays' in holiday:
            holidays_no_header.append(holiday[17:])
        else:
            holidays_no_header.append(holiday)

    holidays_dates_only = []

    for holiday in holidays_no_header:
        for i in range(len(holiday) - 1, -1, -1):
            if holiday[i].isnumeric():
                index = i
                holidays_dates_only.append(holiday[:index + 1])
                break

    holidays_dates_only = [datetime.datetime.strptime(holiday, '%b %d, %Y') for holiday in holidays_dates_only]
    holidays_dates_only = [holiday.strftime('%d/%m/%Y') for holiday in holidays_dates_only]

    return holidays_dates_only


def run_today():
    today = datetime.datetime.today().strftime('%d/%m/%Y')
    holidays_calendar = get_stock_exchange_holidays_calendar()
    return datetime.datetime.today().isoweekday() <= 5 and today not in holidays_calendar


schedule.every().day.at("18:00").do(job)

while True:
    schedule.run_pending()
    time.sleep(1)


# schedule.every(1).minutes.do(job)
# schedule.every(20).seconds.do(job)
# schedule.every().hour.do(job)
# schedule.every(5).to(10).minutes.do(job)
# schedule.every().monday.do(job)
# schedule.every().wednesday.at("13:15").do(job)
# schedule.every().minute.at(":17").do(job)
