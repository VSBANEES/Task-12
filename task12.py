from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from openpyxl import Workbook
from datetime import datetime


def login(username, password):
    driver = webdriver.Chrome()
    driver.get("https://opensource-demo.orangehrmlive.com/")
    driver.find_element(By.ID, 'username').send_keys(username)
    driver.find_element(By.ID, 'password').send_keys(password)
    driver.find_element(By.ID, 'submit').click()

    # Check if login successful
    try:
        WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.ID, 'conf_message')))
        test_result = "Pass"
    except:
        test_result = "Fail"

    driver.quit()
    return test_result


def main():
    # Create Excel workbook and sheet
    wb = Workbook()
    ws = wb.active
    ws.append(["Test ID", "Username", "Password", "Date", "Time of Test", "Name of Tester", "Test Result"])

    # Define usernames and passwords
    credentials = [
        {"username": "Admin", "password": "admin123"},
        {"username": "Admin", "password": "admin@123"},
        {"username": "admin", "password": "admin123"},
        {"username": "admin", "password": "admin@123"},
        {"username": "Admin", "password": "admin123"}
    ]

    # Perform login with each set of credentials and record results in Excel
    for i, cred in enumerate(credentials, start=1):
        username = cred["username"]
        password = cred["password"]
        test_result = login(username, password)
        ws.append([i, username, password, datetime.now().strftime("%Y-%m-%d"),
                   datetime.now().strftime("%H:%M:%S"), "Tester Name", test_result])

    # Save Excel file
    wb.save("login_test_results.xlsx")


if __name__ == "__main__":
    main()
