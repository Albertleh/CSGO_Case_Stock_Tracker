from selenium import webdriver
import openpyxl as op
import time

# Formatiert den HTMl Text zu einem Brauchbaren Format


def get_stock(price):
    stock = ""
    for i in range(len(price)):
        if price[i-1] == "$":
            while price[i] != " ":
                stock += price[i]
                i += 1
    return stock


def get_quantity(price):
    quantity = ""
    for i in price:
        if i == " ":
            break
        else:
            quantity += i
    return quantity

# Ruft die Preise und die Stückzahl ab


def ScrapeStock(url, name, cur_field, column):
    driver = webdriver.Chrome(
        "C:\\Users\\Albert\\PycharmProjects\\CSGO_Cases\\chromedriver.exe")
    driver.get(url)
    time.sleep(1)

    price = driver.find_element_by_xpath(
        '//div[@class= "market_commodity_order_summary"]')
    price = str(price.text)
    print(name, get_stock(price), get_quantity(price))

    driver.close()
    # Daten aufs Excel Sheet schreiben
    sheet[excel_felder[cur_field]] = get_stock(price)
    # Zeilenumbruch in Excel
    for el in excel_felder:
        el = el[:1] + '5'
    sheet[excel_felder[cur_field]] = get_quantity(price)
    for el in excel_felder:
        el = el[:1] + str(column)
    sheet[excel_felder[cur_field]] = get_stock(price)
    for el in excel_felder:
        el = el[:1] + str((column+1))
    sheet[excel_felder[cur_field]] = get_quantity(price)


# Main
URLS = [
    "https://steamcommunity.com/market/listings/730/Danger%20Zone%20Case",
    "https://steamcommunity.com/market/listings/730/Prisma%20Case",
    "https://steamcommunity.com/market/listings/730/Spectrum%20Case",
    "https://steamcommunity.com/market/listings/730/Operation%20Wildfire%20Case",
    "https://steamcommunity.com/market/listings/730/Horizon%20Case"
]

Names = [
    "DangerZoneCase",
    "PrismaCase",
    "SpektrumCase",
    "WildfireCase",
    "HorizonCase"
]

# Excel Zeug
workbook = op.load_workbook(filename="Stock_Sheet.xlsx")
sheet = workbook.get_sheet_by_name('Tabelle1')

column = sheet['G1']
excel_felder = ["B4", "C4", "D4", "E4", "F4"]
current_field = 0

# Loop für jedes Einzelne Case
for case in range(len(URLS)):
    ScrapeStock(URLS[case], Names[case], current_field, column)
    current_field += 1

sheet['G1'] = column+2
workbook.save(filename="Stock_Sheet.xlsx")
