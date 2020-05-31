#! python3
# Scrapes websites to keep track of your property price.

import time
from datetime import date
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Color, Fill, Font
from openpyxl.cell import Cell
from openpyxl import Workbook
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.firefox.options import Options


def zillow_estimate(address):
    print("Getting Zillow Property Estimate...")
    url = 'https://www.zillow.com/homes/'
    driver.get(url)
    driver.find_element_by_class_name('react-autosuggest__input').clear()
    time.sleep(1)
    driver.find_element_by_class_name('react-autosuggest__input').send_keys(address)
    time.sleep(2)
    driver.find_element_by_class_name('searchBtn').click()
    time.sleep(1)
    print("Zillow Estimate -", driver.find_element_by_class_name('zestimate-value').text)
    return driver.find_element_by_class_name('zestimate-value').text


def redfin_estimate(address):
    print("Getting Redfin Property Estimate...")
    url = 'https://www.redfin.com/why-buy-with-redfin'
    driver.get(url)
    time.sleep(2)
    driver.find_element_by_xpath('/html[1]/body[1]/div[1]/div[2]/div[1]/div[1]/header[2]/div[1]/div[2]/div[1]/form[1]/div[1]/div[1]/input[1]').send_keys(address)
    time.sleep(2)
    driver.find_element_by_xpath('/html[1]/body[1]/div[1]/div[2]/div[1]/div[1]/header[2]/div[1]/div[2]/div[1]/form[1]/div[1]/div[1]/input[1]').send_keys(Keys.ENTER)
    time.sleep(1)
    print("Redfin Estimate -", driver.find_element_by_xpath('/html[1]/body[1]/div[1]/div[12]/div[1]/div[1]/section[1]/div[1]/div[1]/div[1]/div[3]/div[1]').text)
    return driver.find_element_by_xpath('/html[1]/body[1]/div[1]/div[12]/div[1]/div[1]/section[1]/div[1]/div[1]/div[1]/div[3]/div[1]').text


def trulia_estimate(address):
    print("Getting Trulia Property Estimate...")
    url = 'https://www.trulia.com/'
    driver.get(url)
    driver.find_element_by_xpath('//input[@id="banner-search"]').send_keys(address)
    time.sleep(2)
    driver.find_element_by_xpath('/html[1]/body[1]/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/div[3]/div[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[2]').click()
    time.sleep(1)
    print("Trulia Estimate -", driver.find_element_by_xpath('//div[@class="Box-sc-8ox7qa-0 iITnWe"]//h3').text)
    return driver.find_element_by_xpath('//div[@class="Box-sc-8ox7qa-0 iITnWe"]//h3').text


def homes_estimate(address):
    print("Getting Homes Property Estimate...")
    url = 'https://www.homes.com/'
    driver.get(url)
    time.sleep(3)
    driver.find_element_by_xpath('//div[@class="bg-white flex space-between items-center mt-1"]//input[@placeholder="Search by Location"]').send_keys(address)
    time.sleep(3)
    driver.find_element_by_xpath('//button[contains(@class,"search-submit-btn bg-orange p-1")]').click()
    time.sleep(3)
    print("Homes Estimate -", driver.find_element_by_xpath('//div[@class="summary-price-display mr-1 font-family-title font-size-xl whitespace-no-wrap"]').text)
    return driver.find_element_by_xpath('//div[@class="summary-price-display mr-1 font-family-title font-size-xl whitespace-no-wrap"]').text


if __name__ == '__main__':
    date_today = date.today().strftime('%m/%d/%Y')
    property_street = input('Enter property street address: ')  #.replace(' ', '-').lower()
    property_city = input('Enter property city: ')  #.replace(' ', '-').lower()
    property_state = input('Enter property state: ')
    property_zip = input('Enter property zip code: ')
    address_format = property_street + ', ' + property_city + ', ' + property_state + ', ' + property_zip
    property_fname = property_street.split(" ", 1)[1] + '.xlsx'

    print('Launching Firefox...')
    options = Options()
    options.add_argument("--headless")
    driver = webdriver.Firefox(options=options)
    driver.implicitly_wait(10)
    driver.maximize_window()

    # Clean Estimates
    zillow_value = 0
    redfin_value = 0
    trulia_value = 0
    homes_value = 0

    zillow_value = zillow_estimate(address_format)
    redfin_value = redfin_estimate(address_format)
    trulia_value = trulia_estimate(address_format)
    homes_value = homes_estimate(address_format)
    '''Work on Cookies or Realtor Captcha
    #realtor_value = realtor_estimate(address_format)'''

    driver.close()
    driver.quit()

    try:
        #wb = load_workbook(filename='HomeEstimates.xlsx')
        wb = load_workbook(filename=property_fname)
        ws = wb.active
    except FileNotFoundError:
        print("File not found, creating new file...")
        wb = Workbook()
        ws = wb.active
        ws.title = 'Home Estimate'
        ws.column_dimensions['A'].width = 10.50
        ws.cell(row=ws.max_row+1, column=1).value = date_today

        # Set Zillow values
        ws['B1'] = 'Zillow'
        zillow_cell = ws.cell(row=1, column=2)
        zillow_cell.font = Font(bold=True)
        zillow_cell.alignment = Alignment(horizontal='center')

        # Set Redfin values
        ws['C1'] = 'Redfin'
        redfin_cell = ws.cell(row=1, column=3)
        redfin_cell.font = Font(bold=True)
        redfin_cell.alignment = Alignment(horizontal='center')

        # Set Trulia values
        ws['D1'] = 'Trulia'
        trulia_cell = ws.cell(row=1, column=4)
        trulia_cell.font = Font(bold=True)
        trulia_cell.alignment = Alignment(horizontal='center')

        # Set Homes values
        ws['E1'] = 'Homes'
        homes_cell = ws.cell(row=1, column=4)
        homes_cell.font = Font(bold=True)
        homes_cell.alignment = Alignment(horizontal='center')

        print("Adding property values for", date_today)
        new_row = ws.max_row + 1
        ws.cell(row=new_row, column=1).value = date_today
        ws.cell(row=new_row, column=2).value = zillow_value
        ws.cell(row=new_row, column=3).value = redfin_value
        ws.cell(row=new_row, column=4).value = trulia_value
        ws.cell(row=new_row, column=5).value = homes_value
        wb.save(property_fname)
        print('Changes have been saved to', property_fname)

    date_cell = ws.cell(row=ws.max_row, column=1)

    if date_cell.value != date_today:
        print("Adding property values for", date_today)
        new_row = ws.max_row + 1
        ws.cell(row=new_row, column=1).value = date_today
        ws.cell(row=new_row, column=2).value = zillow_value
        ws.cell(row=new_row, column=3).value = redfin_value
        ws.cell(row=new_row, column=4).value = trulia_value
        ws.cell(row=new_row, column=5).value = homes_value
        wb.save(property_fname)
        print('Changes have been saved to', property_fname)
    elif ws:
        print("The script has already run today")
