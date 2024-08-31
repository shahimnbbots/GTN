import math
from tkinter import filedialog
import openpyxl
import tkinter as tk
import warnings
import pandas as pd
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from selenium.webdriver import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
import os
import time

warnings.filterwarnings("ignore", message="wmf image format is not supported", category=UserWarning)


def invoice_creation(booking_number, country_value):
    service = Service()
    options = Options()
    options.add_experimental_option('detach', True)
    driver = webdriver.Chrome(service=service, options=options)
    wait = WebDriverWait(driver, 30)
    driver.get("https://network.infornexus.com/login")

    username = driver.find_element(By.CSS_SELECTOR, 'input[name="userid"]')
    username.send_keys("balaji.pandu@shahi")
    password = driver.find_element(By.CSS_SELECTOR, 'input[name="uPassword"]')
    password.send_keys("SEPL@1234")
    login_button = driver.find_element(By.CSS_SELECTOR, 'input#loginButton')
    login_button.click()
    continue_button = wait.until(ec.presence_of_element_located((By.ID, 'submitbutton')))
    continue_button.click()
    time.sleep(8)

    applications_link = wait.until(ec.element_to_be_clickable((By.ID, 'navmenu__applications')))
    applications_link.click()
    invoice_optn = driver.find_element(By.XPATH, '//*[@id="navmenu__inprogressinvoices"]')
    invoice_optn.click()
    invoice_input_label = wait.until(
        ec.presence_of_element_located((By.XPATH, '//span[contains(text(), "Invoice Number")]/ancestor::label')))
    input_field = invoice_input_label.find_element(By.XPATH,
                                                   '//*[@id="fixed-outer-filter-content"]/div/div/div[2]/div/div[2]/span/span[2]/span/input')
    booking_number_int = int(booking_number)
    booking_number_str = str(booking_number_int)
    if len(booking_number_str) == 11:
        booking_number_str += '0'
    input_field.send_keys(booking_number_str)
    print(booking_number_str)
    apply = driver.find_element(By.XPATH, '//*[@id="fixed-outer-filter-content"]/div/div/div[2]/span/button[2]')
    apply.click()
    time.sleep(3)
    # Locate the table
    table = driver.find_element(By.CSS_SELECTOR,
                                '#active > div > span > span > div > div > div.flexresults.hoverFocuses.singleFocus > table')
    tbody = table.find_element(By.TAG_NAME, 'tbody')
    rows = tbody.find_elements(By.TAG_NAME, 'tr')

    for row in rows:
        cols = row.find_elements(By.TAG_NAME, 'td')
        if len(cols) >= 4:
            try:
                qty = cols[3].find_element(By.TAG_NAME, 'span').text
                print(qty)
                # use total_quantity here
            except Exception as e:
                print(f"Could not retrieve text from the fourth column: {e}")
    try:
        inv_number = wait.until(
            ec.element_to_be_clickable((By.XPATH, f'//a[contains(text(), "{int(booking_number)}")]')))
        action_chains = ActionChains(driver)
        action_chains.double_click(inv_number).perform()
    except NoSuchElementException:
        print("Element not found")
    time.sleep(3)
    # for _ in range(4):
    #     driver.execute_script("javascript: submitUserAction('NEXT_STEP');")
    #     time.sleep(2)  # Adjust sleep time if necessary to wait for any page loads or actions to complete

    # Click on "Additional Terms"
    try:
        parent_element = wait.until(
            ec.presence_of_element_located((By.XPATH, '//td/a[contains(@href, "jumpToStep(\'AdditionalTerms\')")]')))
        additional_terms_link = parent_element.find_element(By.XPATH, './span[contains(text(), "Additional Terms")]')
        additional_terms_link.click()
        print("Clicked on 'Additional Terms'.")
    except NoSuchElementException:
        print("Could not find 'Additional Terms' link.")
    time.sleep(3)
    # Click the first and third checkboxes in the specified table
    try:
        # Locate and click the first checkbox
        first_checkbox = driver.find_element(By.NAME, 'CommercialInvoice_phrase__1_ackOfTerm_checkbox')
        if not first_checkbox.is_selected():
            first_checkbox.click()
            print("Clicked the first checkbox.")

        # Locate and click the third checkbox
        third_checkbox = driver.find_element(By.NAME, 'CommercialInvoice_phrase__3_ackOfTerm_checkbox')
        if not third_checkbox.is_selected():
            third_checkbox.click()
            print("Clicked the third checkbox.")
    except NoSuchElementException:
        print("Could not find the specified checkboxes.")
    time.sleep(3)
    # Locate the table with the checkboxes
    checkbox_table = wait.until(ec.presence_of_element_located((By.CLASS_NAME, 'pagecolor')))
    rows = checkbox_table.find_elements(By.TAG_NAME, 'tr')

    for row in rows:
        try:
            cell_text = row.find_element(By.TAG_NAME, 'font').text
            if country_value == "USA":
                if any(text in cell_text for text in [
                    'COMMERCIAL INVOICE COPY, STATING SHIPMENT AUTHORIZATION NUMBER, MANUFACTURER\'S NAME, MANUFACTURER\'S ADDRESS OR MID #, COUNTRY OF ORIGIN.',
                    'PACKING LIST COPY.',
                    'PACKING SUMMARY COPY.']):
                    checkbox = row.find_element(By.XPATH, './/input[@type="checkbox"]')
                    if not checkbox.is_selected():
                        checkbox.click()
                        print(f"Clicked checkbox for: {cell_text}")
            elif country_value == "CAN":
                if any(text in cell_text for text in [
                    'COMMERCIAL INVOICE COPY, STATING SHIPMENT AUTHORIZATION NUMBER, MANUFACTURER\'S NAME, MANUFACTURER\'S ADDRESS OR MID #, COUNTRY OF ORIGIN.',
                    'CANADA CUSTOMS INVOICE (IF DESTINATION COUNTRY IS CANADA), STATING SHIPMENT AUTHORIZATION NUMBER, MANUFACTURER\'S NAME, MANUFACTURER\'S ADDRESS OR MID #, COUNTRY OF ORIGIN.',
                    'PACKING LIST COPY.',
                    'PACKING SUMMARY COPY.']):
                    checkbox = row.find_element(By.XPATH, './/input[@type="checkbox"]')
                    if not checkbox.is_selected():
                        checkbox.click()
                        print(f"Clicked checkbox for: {cell_text}")
                        time.sleep(1)
        except NoSuchElementException:
            continue
    try:
        parent_element = wait.until(
            ec.presence_of_element_located((By.XPATH, '//td/a[contains(@href, "jumpToStep(\'Review\')")]')))
        additional_terms_link = parent_element.find_element(By.XPATH, './span[contains(text(), "Preview")]')
        additional_terms_link.click()
        print("Clicked on 'Additional Terms'.")
    except NoSuchElementException:
        print("Could not find 'Additional Terms' link.")


# Choose the file using file dialog
filetypes = (('excel files', '*.xls'), ('excel files', '*.xlsx'), ('excel files', '*.ods'))
filename = filedialog.askopenfilename(
    title='Open excel',
    initialdir="/",
    filetypes=filetypes
)

# Read the data from the chosen file
df = pd.read_excel(filename)
ci_number_a = df['CI number'].tolist()
format_value_a = df['Category'].tolist()
booking_number_a = df['Booking Number'].tolist()
country_value_a = df['Country'].tolist()

opdict = {}
current_booking_number = None
current_format = None
current_country = None

for i in range(len(ci_number_a)):
    ci_number = ci_number_a[i]
    booking_number = booking_number_a[i]
    format_value = format_value_a[i]
    country_value = country_value_a[i]

    if pd.notna(ci_number):
        if math.isnan(booking_number):  # Check for nan values
            booking_number = current_booking_number
        else:
            current_booking_number = booking_number
            current_format = format_value
            current_country = country_value
        if booking_number not in opdict:
            opdict[booking_number] = {'CI_numbers': [ci_number], 'Category': current_format, 'Country': current_country}
        else:
            opdict[booking_number]['CI_numbers'].append(ci_number)

print(opdict)

for booking_number, data in opdict.items():
    ci_numbers = data['CI_numbers']
    format_value = data['Category']
    country_value = data['Country']
    gtn(booking_number, country_value)
    break
