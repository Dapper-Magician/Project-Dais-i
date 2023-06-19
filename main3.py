import pandas as pd
from openpyxl import Workbook
from selenium import webdriver
from bs4 import BeautifulSoup
import requests
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from webdriver_manager.chrome import ChromeDriverManager

# Setup chrome options for Selenium
chrome_options = Options()
chrome_options.add_argument("--headless")  # Ensure GUI is off
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")

# Set path to chromedriver as per your configuration
webdriver_service = Service(ChromeDriverManager().install())

class DaisiBot:
    def __init__(self, voucher_path):
        self.voucher_path = voucher_path
        self.data = pd.DataFrame()
        self.driver = webdriver.Chrome()  # assuming Chrome; adjust as necessary

    def read_voucher(self):
        self.data = pd.read_excel(self.voucher_path)
        self.validate_data()  # to be implemented
        print("Voucher data read successfully.")

    def validate_data(self):
        # Just an example validation, adjust according to your needs
        if self.data.isnull().values.any():
            print("Data contains null values. Please check the voucher.")
        else:
            print("Data validated successfully.")

    def create_template(self):
        # Implement Excel template creation here using openpyxl
        pass

    def setup_webdriver(self):
        self.driver = webdriver.Firefox(executable_path='path/to/geckodriver')  # replace path with your actual path
        print("Webdriver setup successfully.")

    def login_to_daisi(username, password):
        driver = webdriver.Chrome(service=webdriver_service, options=chrome_options)
        driver.get("https://www.iccbdaisi.org/public/login/loginPage.jsp?")  # Confirm is website

        # Fill in username
        username_field = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'userName')))
        username_field.send_keys(username)

        # Fill in password
        password_field = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'password')))
        password_field.send_keys(password)

        # Click login button
        login_button = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME, 'login')))
        login_button.click()

        # Click on Daisi-Access-V2
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//a[contains(text(), "V2 DAISI")]'))).click()

        return driver

    # Execute login
    driver = login_to_daisi('LLloyd', 'CMAA3043')

    def navigate_to_page(self, page_name):
        # Replace 'page_name' with the actual XPath or other locator strategy of the page link/Ask Proteus if shortcut
        page_link = WebDriverWait(self.driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, f"//a[text()='{page_name}']")))
        page_link.click()
        print(f"Navigated to {page_name} page.")

    def select_year(driver, year):
        # Click dropdown
        dropdown = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CLASS_NAME, 'ui-dropdown-trigger-icon')))
        dropdown.click()

        # Select year from dropdown
        year_option = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, f'//li[@aria-label="{year}"]')))
        year_option.click()

    def navigate_to_students(driver):
        # Click on the "Students" header dropdown
        students_header = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//span[text()="Students"]')))
        students_header.click()

    def add_new_student(driver):
        # Click on the "Add New" option under "Students"
        add_new_option = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//span[text()="Add New"]')))
        add_new_option.click()

################# Add Student input form
    def enter_student_search(driver, student_name):
        # Enter student name in search
        search_field = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//input[@placeholder="Enter Class/Section No to Search"]')))
        search_field.send_keys(student_name)


    def navigate_to_classes(driver):
        # Click on the "Classes" header dropdown
        classes_header = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//span[text()="Classes"]')))
        classes_header.click()

    def list_search_classes(driver):
        # Click on the "List/Search" option under "Classes"
        list_search_option = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//span[text()="List/Search"]')))
        list_search_option.click()

    def list_current_fy(driver):
        # Click on the "List Current FY" option under "Classes"
        list_current_fy_option = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//span[text()="List Current FY"]')))
        list_current_fy_option.click()







    # Execute functions
    select_year(driver, 2023)
    enter_student_search(driver, 'student_name')  # replace with actual student name
    navigate_to_students(driver)
    navigate_to_classes(driver)
    add_new_student(driver)
    list_search_classes(driver)
    list_current_fy(driver)


    def enter_data(self, student_data):
        for data_field, value in student_data.items():
            # Replace 'data_field' with the actual name or id or XPath of the HTML input fields
            input_field = WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.NAME, data_field)))
            input_field.clear()
            input_field.send_keys(value)
            print(f"Entered {value} into {data_field} field.")
        submit_button = WebDriverWait(self.driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//button[@type='submit']")))
        submit_button.click()
        print("Data submitted successfully.")



    # Remaining methods...

    def run(self, username, password, year):
        self.read_voucher()
        self.validate_data()
        self.setup_webdriver()
        self.login(username, password)
        self.select_year(year)
        self.navigate_to_page("Class Page")
        for student_data in self.voucher_data:
            self.navigate_to_page("Attendance Page")
            self.enter_data(student_data)
            self.navigate_to_page("Roster Page")
            self.enter_data(student_data)


bot = DaisiBot("path/to/your/voucher.xlsx")
bot.run("username", "password", "2023")
