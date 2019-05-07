# coding=utf-8
from selenium import webdriver
import logging
import os
from xlsxwriter.workbook import Workbook
import datetime
logging.basicConfig(level = logging.INFO)


def setup_selenium_driver():
    logging.info('starting init selenium')
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument('--no-sandbox')
    chrome_options.add_argument('--headless')
    selenium_driver = webdriver.Chrome(chrome_options=chrome_options)
    logging.info('finished init selenium')
    return selenium_driver


def parse(selenium_driver):
    logging.info('staring parsing')
    ## Implemented CODE by Ivan Miletic
    # define the driver
    driver = setup_selenium_driver()
    # base url
    url = 'https://www.tv-ratingen.de/de/ueber-uns/sportsuche/?Special_Club_SearchList843_page='
    result = []
    # for pages from 1 to 7 i used range()
    for x in range(1,8):
        driver.get(f"{url}{x}")
        # finding elements using xpath
        t1 = driver.find_elements_by_xpath('/html/body/div[2]/div/div[2]/div/div[2]/div[2]/div/div/div[2]/table/tbody/tr[*]/td[1]')
        t2 = driver.find_elements_by_xpath('/html/body/div[2]/div/div[2]/div/div[2]/div[2]/div/div/div[2]/table/tbody/tr[*]/td[2]')
        t3 = driver.find_elements_by_xpath('/html/body/div[2]/div/div[2]/div/div[2]/div[2]/div/div/div[2]/table/tbody/tr[*]/td[3]')
        t4 = driver.find_elements_by_xpath('/html/body/div[2]/div/div[2]/div/div[2]/div[2]/div/div/div[2]/table/tbody/tr[*]/td[4]')
        
        # create lists and append to result
        for num in range(len(t1)):
            result.append([t1[num].text, t2[num].text.split('-')[0].strip(), t2[num].text.split('-')[1].strip(), t3[num].text, t4[num].text])

    #print(lista)
    # close driver
    driver.close()
    ## END of implemeted code!
    logging.info('finished parsing results')
    return result


def save_results(data):
    logging.info('starting saving results')
    header = [['course_title', 'category', 'day', 'start_time', 'end_time',
              'location', 'description', 'level', 'trainer', 'other']]
    data = header + data
    filename = os.path.basename(__file__).split('.')[0] + '_' + str(datetime.datetime.now()) + '.xlsx'
    workbook = Workbook(filename)
    worksheet = workbook.add_worksheet()
    for row,line in enumerate(data):
        for col,entry in enumerate(line):
            worksheet.write(row, col, entry)
    workbook.close()
    logging.info('finished saving results to file {}'.format(filename))



driver = setup_selenium_driver()
scrape_results = parse(driver)
save_results(scrape_results)
