# selenium 4
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import TimeoutException

from openpyxl import load_workbook
import time

wb = load_workbook(filename = 'D:\Python_project\automation_data_input1\data.xlsx')

sheetRange = wb['Sheet1']

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))

driver.get("https://demoqa.com/webtables")
driver.maximize_window()
driver.implicitly_wait(10)

# looping input 
i = 2
while i <= len(sheetRange['A']):
    Firstname = sheetRange['A'+str(i)].value
    Lastname = sheetRange['B'+str(i)].value
    email = sheetRange['C'+str(i)].value
    age = sheetRange['D'+str(i)].value
    salary = sheetRange['E'+str(i)].value
    department = sheetRange['F'+str(i)].value
    
    driver.find_element(By.ID, "addNewRecordButton").click()

    try:
        WebDriverWait(driver,10).until(EC.visibility_of_element_located((By.XPATH, "/html/body/div[4]/div/div")))

        driver.find_element(By.ID, "firstName").send_keys(Firstname)
        driver.find_element(By.ID, "lastName").send_keys(Lastname)
        driver.find_element(By.ID, "userEmail").send_keys(email)
        driver.find_element(By.ID, "age").send_keys(age)
        driver.find_element(By.ID, "salary").send_keys(salary)
        driver.find_element(By.ID, "department").send_keys(department)
        driver.find_element(By.ID, "submit").click()

    except TimeoutException:
        print("form gak muncul kang")
        pass

    time.sleep(1)
    i = i + 1

print("beres kang")