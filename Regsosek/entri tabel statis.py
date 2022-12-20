from openpyxl import load_workbook
from selenium import webdriver
from selenium.webdriver.support import  expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import pyautogui
import time

driver = webdriver.Chrome()
driver.get("https://tobasamosirkab.bps.go.id/backend")

WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="dialog"]')))
driver.maximize_window()

driver.find_element(By.CLASS_NAME, 'ui-dialog-titlebar-close').click()
driver.find_element(By.ID, 'LoginForm_username').send_keys('dapas')
driver.find_element(By.ID, 'LoginForm_password').send_keys('nokia8250')
driver.find_element(By.NAME, 'yt0').click()
WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="mainmenu"]')))
driver.find_element(By.XPATH, '//*[@id="yw0"]/li[4]').click()
WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.ID, 'dropdownlistWrappersubjectHtml')))
WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.ID, 'dropdownlistContentsubjectCSAHtml')))
WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.ID, 'dropdownlistContentflagKategoriHtml')))
    

wb = load_workbook(filename="E:\BPS TOBA\Selenium Automatitation\data.xlsx")
sheetRange = wb['Sheet1']
i = 15
current = 0
while i <= len(sheetRange['A']):
    subject = sheetRange['A'+str(i)].value
    csa = sheetRange['B'+str(i)].value
    kategori = sheetRange['C'+str(i)].value
    jdlIndo = sheetRange['D'+str(i)].value
    jdlInggris = sheetRange['E'+str(i)].value
    pathXls = sheetRange['F'+str(i)].value
    pathHtml = sheetRange['G'+str(i)].value
    
    driver.find_element(By.ID, 'fileExcelIndo').send_keys(pathXls)
    driver.find_element(By.ID, 'fileExcelInggris').send_keys(pathXls)

    driver.find_element(By.ID, 'dropdownlistContentsubjectHtml').click()
    pyautogui.typewrite(subject)
    driver.find_element(By.ID, 'listitem6innerListBoxsubjectHtml').click()

    driver.find_element(By.ID, 'dropdownlistContentsubjectCSAHtml').click()
    pyautogui.typewrite(csa)
    driver.find_element(By.ID, 'listitem3innerListBoxsubjectCSAHtml').click()

    driver.find_element(By.ID, 'dropdownlistContentflagKategoriHtml').click()
    pyautogui.typewrite(kategori)
    driver.find_element(By.ID, 'listitem1innerListBoxflagKategoriHtml').click()

    driver.find_element(By.ID, 'jdlIndo').send_keys(jdlIndo)
    driver.find_element(By.ID, 'jdlInggris').send_keys(jdlInggris)

    driver.execute_script("window.open('')")
    driver.switch_to.window(driver.window_handles[current + 1])
    driver.get("file:///" + pathHtml)

    time.sleep(2)
    pyautogui.hotkey('ctrl','u')
    time.sleep(2)
    pyautogui.hotkey('ctrl','a')
    time.sleep(1)
    pyautogui.hotkey('ctrl','c')

    driver.switch_to.window(driver.window_handles[current])
    driver.find_element(By.ID, 'kodeHtmlIndo').clear()
    driver.find_element(By.ID, 'kodeHtmlInggris').clear()
    time.sleep(2)
    driver.find_element(By.ID, 'kodeHtmlIndo').click()
    pyautogui.hotkey('ctrl','v')
    driver.find_element(By.ID, 'kodeHtmlInggris').click()
    pyautogui.hotkey('ctrl','v')

    driver.find_element(By.ID, 'simpanLinkHtml').click()
    time.sleep(3)

    driver.switch_to.window(driver.window_handles[current + 2])
    driver.execute_script("window.open('')")
    driver.switch_to.window(driver.window_handles[current + 3])
    driver.get("https://tobasamosirkab.bps.go.id/backend/index.php/dataDynamic/menu/id/4")
    WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.ID, 'dropdownlistWrappersubjectHtml')))
    WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.ID, 'dropdownlistContentsubjectCSAHtml')))
    WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.ID, 'dropdownlistContentflagKategoriHtml')))
    
    i += 1
    current += 3
