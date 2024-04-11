import time
import pandas as pd
from selenium import webdriver
from selenium.common import NoSuchElementException
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

chrome_options = webdriver.ChromeOptions()
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
# driver = webdriver.Chrome()
url = "https://www.tis.bizfile.gov.sg/ngbtisinternet/faces/oracle/webcenter/portalapp/pages/TransactionMain.jspx?selectedETransId=dirSearch"
driver.get(url)
input_field = driver.find_element(By.XPATH, '//input[@id="pt1:r1:0:sv1:it1::content"]')
UEN_input = '2003010'
input_field.send_keys(UEN_input)
submit_button = driver.find_element(By.XPATH, '//button[@id="pt1:r1:0:sv1:cb1"]')
submit_button.click()
time.sleep(50)

name_list = []
uen_list = []
status_list = []
address_list = []
industry_list = []
annual_return_list = []
annual_general_meeting_list = []
financial_statement_list = []

while True:
    name_elements = driver.find_elements(By.XPATH,
                                         '//div[@class="nameMRserchright af_panelGroupLayout"]//span[@class="bizpara1"]')
    uen_elements = driver.find_elements(By.XPATH,
                                        '//div[@class="nameMRserchright af_panelGroupLayout"]//span[@class="bizpara1 text_uppercase nameMRserchRegNo"]')
    status_elements = driver.find_elements(By.XPATH,
                                           '//div[@class="nameMRserchright af_panelGroupLayout"]//span[@class="bizpara1 text_uppercase"]')
    address_elements = driver.find_elements(By.XPATH,
                                            '//div[@class="nameMRserchright af_panelGroupLayout"]//a[@class="orgpara_nrml orgparaMRserch420 af_commandLink"]')
    industry_elements = driver.find_elements(By.XPATH,
                                             '//div[@class="nameMRserchrightdes af_panelGroupLayout"]//span[@class="bizpara1 text_uppercase"]')
    annual_return_elements = driver.find_elements(By.XPATH, '//div[@class="address_search22 af_panelGroupLayout"]')
    for parent_div in annual_return_elements:
        try:
            annual_return_element = parent_div.find_element(By.XPATH,
                                                            './/div[contains(@id, "pgl80")]//span[@class="bizpara1"]')
            annual_return_list.append(annual_return_element.text.strip())
        except NoSuchElementException:
            annual_return_list.append(None)

    annual_general_meeting_elements = driver.find_elements(By.XPATH,
                                                           '//div[@class="address_search22 af_panelGroupLayout"]')
    for parent_div in annual_general_meeting_elements:
        try:
            annual_general_meeting_element = parent_div.find_element(By.XPATH,
                                                                     './/div[contains(@id, "pgl81")]//span[@class="bizpara1"]')
            annual_general_meeting_list.append(annual_general_meeting_element.text.strip())
        except NoSuchElementException:
            annual_general_meeting_list.append(None)
    financial_statement_elements = driver.find_elements(By.XPATH,
                                                        '//div[@class="address_search22 af_panelGroupLayout"]')

    for element in financial_statement_elements:
        try:
            financial_statement_element = element.find_element(By.XPATH, './/div[contains(@id, "pgl83")]')
            if financial_statement_element.find_elements(By.XPATH, './/span[@class="bizpara1"]'):
                value = financial_statement_element.find_element(By.XPATH, './/span[contains(@class, "bizpara1")]').text
                financial_statement_list.append(value)
            else:
                years_elements = financial_statement_element.find_elements(By.XPATH,
                                                                           './/span[contains(@class, "product_year")]')
                years = ', '.join(year_element.text for year_element in years_elements)
                financial_statement_list.append(years)
        except NoSuchElementException:
            financial_statement_list.append(None)

    current_name = ''
    for element in name_elements:
        text = element.text.strip()
        if text:
            current_name += text + ' '
        else:
            name_list.append(current_name.strip())
            current_name = ''

    if current_name:
        name_list.append(current_name.strip())
    name_list = [name for name in name_list if name]
    for uen in uen_elements:
        uen_list.append(uen.text.strip())
    for status in status_elements:
        status_list.append(status.text.strip())
    for address in address_elements:
        address_list.append(address.text.strip())
    for industry in industry_elements:
        industry_list.append(industry.text.strip())

    next_button = driver.find_element(By.XPATH,
                                      '//div[@class="arrow_directSearch af_panelGroupLayout"]//a[@id="pt1:r1:0:sv1:commandLink3"]')
    if next_button.get_attribute("aria-disabled"):
        break
    else:
        next_button.click()
        time.sleep(2)

df = pd.DataFrame({
    'Name': name_list,
    'UEN': uen_list,
    'Status': status_list,
    'Address': address_list,
    'Industry': industry_list,
    'Annual Return': annual_return_list,
    'Annual General Meeting': annual_general_meeting_list,
    'Financial Statements Filed': financial_statement_list
})
excel_file_path = 'scraped_data.xlsx'
df.to_excel(excel_file_path, index=False)

driver.close()
