# selenium 4
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

from openpyxl import load_workbook
import time

wb = load_workbook(filename="D:\Python project\automation-data-input1\data.xlsx")

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
    
