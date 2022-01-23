import pandas as pd
from selenium import webdriver
from time import sleep
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException
import os
from parsel import Selector
import sys
from time import strftime

options = webdriver.ChromeOptions()
options.add_argument('--ignore-certificate-errors')
options.add_argument('--ignore-ssl-errors')
options.add_argument('--ignore-certificate-errors-spki-list')
options.add_experimental_option('excludeSwitches', ['enable-logging'])
driver = webdriver.Chrome('chromedriver.exe', chrome_options=options)
driver.maximize_window()
driver.implicitly_wait(5)

sys.setrecursionlimit(1500)


def scroll(driver, timeout):
    scroll_pause_time = timeout

    # Get scroll height
    last_height = driver.execute_script("return document.body.scrollHeight")

    while True:
        # Scroll down to bottom
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

        # Wait to load page
        sleep(scroll_pause_time)

        # Calculate new scroll height and compare with last scroll height
        new_height = driver.execute_script("return document.body.scrollHeight")
        if new_height == last_height:
            # If heights are the same it will exit the function
            break
        last_height = new_height

lego_numbers = pd.read_csv("lego_numbers.csv")
lego_numbers["Lego-Number"] = lego_numbers["Lego-Number"].astype(str)
try:
    for i, row in lego_numbers.iterrows():
        try:
            driver.get("https://www.bricklink.com/catalogPG.asp")
            try:
                driver.find_element_by_xpath('//*[@class="far fa-times text medium"]').click()
            except:
                pass
            #sending item no. to the search box
            driver.find_elements_by_xpath("//*[contains(@name, 'itemNo')]")[1].send_keys(row["Lego-Number"])
            check_boxes = driver.find_elements_by_xpath("//*[@type='CHECKBOX']")[2:]
            for y in check_boxes:
                y.click()
            sub_button = driver.find_element_by_xpath("//*[contains(@value, 'Get Value')]").click()

            try:
                driver.find_element_by_xpath('//*[@class="far fa-times text medium"]').click()
            except:
                pass

            try:
                POV_last_6_months = driver.find_elements_by_xpath("//*[@width='50%']")[0].text
                POV_last_6_months = POV_last_6_months.split("\n")[2].replace("US ", "").replace("$","")
                POV_last_6_months = float(POV_last_6_months)
                lego_numbers.loc[i, "Part out Value in $ (POV) last 6 months sales (USD)"] = "%.2f" % POV_last_6_months
                POV_last_6_months_EUR = POV_last_6_months * float(row["currency-factor"])
                lego_numbers.loc[i, "Part out Value (POV) last 6 months sales (EUR)"] = "%.2f" % POV_last_6_months_EUR
            except:
                pass

            try:
                POV_current_for_sale = driver.find_elements_by_xpath("//*[@width='50%']")[1].text
                POV_current_for_sale = POV_current_for_sale.split("\n")[2].replace("US ", "").replace("$","")
                POV_current_for_sale = float(POV_current_for_sale)
                lego_numbers.loc[i, "Part out Value in $ (POV) current for sale (USD)"] = "%.2f" % POV_current_for_sale
                POV_current_for_sale_EUR = POV_current_for_sale * float(row["currency-factor"])
                lego_numbers.loc[i, "Part out Value (POV) current for sale (EUR)"] = "%.2f" % POV_current_for_sale_EUR
            except:
                pass
            sleep(2)
        except:
            continue
except:
    pass

date_string = strftime("%d-%m-%Y")
file_name = date_string + "_Lego"

lego_numbers.to_csv(file_name + ".csv", index=False)

writer = pd.ExcelWriter(file_name + '.xlsx',
                        engine='xlsxwriter',
                        options={'strings_to_urls': False})  # pylint: disable=abstract-class-instantiated
lego_numbers.to_excel(writer, index=False, sheet_name='Lego', startrow=1, header=False)
workbook = writer.book
worksheet = writer.sheets['Lego']

# Add a header format.
header_format = workbook.add_format({
    'bold': True,
    'text_wrap': True,
    'valign': 'top',
    'fg_color': '#03fcf0',
    'border': 1
})

writer.save()
driver.quit()
