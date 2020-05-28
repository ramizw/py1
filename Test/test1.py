from selenium import webdriver
from selenium.webdriver.common.by import By
import time
import os
import xlrd
import xlwt

driver = webdriver.Chrome("..\\drivers\\chromedriver.exe")
driver.set_page_load_timeout("10")
baseURL = "https://vast-dawn-73245.herokuapp.com/"
driver.maximize_window()
driver.implicitly_wait(3)
driver.get(baseURL)

path = "..\\data\\TestData.xlsx"
inputXl = xlrd.open_workbook(path)
inputSheet = inputXl.sheet_by_index(0)

wb = xlwt.Workbook()
ws = wb.add_sheet("Result")

for i in range (1,inputSheet.nrows):
    for j in range(1,inputSheet.ncols):
        inputDate = []
        inputDate.append(inputSheet.cell_value(i,j))
        print(inputDate)
        driver.find_element_by_xpath("//input[@class='form-control']").send_keys(inputDate)
        driver.find_element_by_xpath("//input[@class='btn btn-default']").click()
        result = driver.find_element_by_xpath("/html/body/div[2]/div/div[2]/div").text
        print(result)


time.sleep(5)
driver.quit()