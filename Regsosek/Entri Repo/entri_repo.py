from openpyxl import load_workbook
from selenium import webdriver
from selenium.webdriver.support import  expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import pyautogui
import time

driver = webdriver.Chrome()
driver.get("https://sso.bps.go.id/")

WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="password"]')))
driver.maximize_window()

driver.find_element(By.XPATH, '//*[@id="username"]').send_keys('mangihut')
driver.find_element(By.XPATH, '//*[@id="password"]').send_keys('tonang051')
driver.find_element(By.XPATH, '//*[@id="kc-login"]').click()
WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="totp"]')))


driver.get("https://regsosek-repo.cloud.bps.go.id/entry")
WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="app"]/div/main/div/div/div[2]/div/form/button[2]')))
driver.find_element(By.XPATH, '//*[@id="app"]/div/main/div/div/div[2]/div/form/button[2]').click()

WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="app"]/div[1]/div/div[2]/div[2]/div[1]/input')))
driver.get("https://regsosek-repo.cloud.bps.go.id/entry")
WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="app"]/div[1]/div/div[2]/div[2]/div[2]/input')))
driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div/div[2]/div[2]/div[1]/input').send_keys('balige')
pyautogui.hotkey('enter')

driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div/div[2]/div[2]/div[2]/input').send_keys('30')
pyautogui.hotkey('enter')
# driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div/div[2]/div[2]/div[2]/input').click()
# driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div/div[2]/div[2]/div[1]/input').send_keys('balige')
# driver.find_element(By.XPATH, '//*[@id="multiselect-option-1206030"]').click()
# driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div/div[2]/div[2]/div[2]/input').send_keys('')
# driver.find_element(By.XPATH, '//*[@id="multiselect-option-1206030018"]').click()