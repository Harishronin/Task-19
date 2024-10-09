"""
 Name : Harish kumar
 Date : 09-Oct-2024
 Program 1 : Using data driven testing framework,page object model,explicit wait,expected conditions,pytest
1.Create a excel file which comprises test id,username,password,time,time of test,tester name,test results for login in to portal
2.https://opensource-demo.orangehrmlive.com/web/index.php/auth/login
3.Login into portal using username and password provided in the excel file.try use of 5 username and password
4.If the login is successful your python code will write in the excel file whether your test passed or test failed.
5.Do not use sleep().
 """

#page object model(POM)
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

class LoginPage:
    def __init__(self, driver):
        self.driver = driver
        self.username_field = (By.NAME, "username")
        self.password_field = (By.NAME, "password")
        self.login_button = (By.XPATH, "//button[@type='submit']")

    def enter_username(self, username):
        WebDriverWait(self.driver, 10).until(
            EC.visibility_of_element_located(self.username_field)).clear()
        self.driver.find_element(*self.username_field).send_keys(username)

    def enter_password(self, password):
        WebDriverWait(self.driver, 10).until(
            EC.visibility_of_element_located(self.password_field)).clear()
        self.driver.find_element(*self.password_field).send_keys(password)

    def click_login(self):
        WebDriverWait(self.driver, 10).until(
            EC.element_to_be_clickable(self.login_button)).click()

    def is_login_successful(self):
        try:
            WebDriverWait(self.driver, 10).until(
                EC.visibility_of_element_located((By.XPATH, "//span[text()='Dashboard']")))
            return True
        except:
            return False

#Pytest
import pytest
import openpyxl
from selenium import webdriver
from page_objects.login_page import LoginPage
from datetime import datetime

# data from Excel
def load_test_data(file_name):
    workbook = openpyxl.load_workbook(file_name)
    sheet = workbook.active
    test_data = []
    for row in range(2, sheet.max_row + 1):
        test_data.append({
            'test_id': sheet.cell(row, 1).value,
            'username': sheet.cell(row, 2).value,
            'password': sheet.cell(row, 3).value,
            'time': sheet.cell(row, 4).value,
            'tester_name': sheet.cell(row, 5).value
        })
    return test_data, workbook, sheet

# test result in Excel
def write_test_result(sheet, row, result, workbook):
    sheet.cell(row, 6).value = result
    workbook.save('test_data.xlsx')

# pytest fixture for setup
@pytest.fixture
def setup():
    driver = webdriver.Chrome()  # Make sure to download ChromeDriver
    driver.get('https://opensource-demo.orangehrmlive.com/web/index.php/auth/login')
    driver.maximize_window()
    yield driver
    driver.quit()

# Test case for login
@pytest.mark.parametrize("test_case", load_test_data('test_data.xlsx')[0])
def test_login(test_case, setup):
    driver = setup
    login_page = LoginPage(driver)

    # username and password from test case
    username = test_case['username']
    password = test_case['password']

    # Perform login actions
    login_page.enter_username(username)
    login_page.enter_password(password)
    login_page.click_login()

    # Check if login was successful
    if login_page.is_login_successful():
        result = 'PASS'
    else:
        result = 'FAIL'

    # Write test result back to Excel
    data, workbook, sheet = load_test_data('test_data.xlsx')
    write_test_result(sheet, test_case['test_id'] + 1, result, workbook)

