#https://zwbetz.com/download-chromedriver-binary-and-add-to-your-path-for-automated-functional-testing/
from selenium import webdriver
import time
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import WebDriverException


driver = webdriver.Chrome()
file_name = r'C:\Users\docha\OneDrive\Leka\PPA LTR\E-taotluskeskkond01.html'
file_name = 'file://' + file_name
driver.get(file_name)
#driver.get("https://www.python.org")
#search_bar = driver.find_element(By.NAME, 'q')
#search_bar.clear()
#search_bar.send_keys("getting started with python")
#search_bar.send_keys(Keys.RETURN)
#print(driver.current_url)
#driver.maximize_window()
#time.sleep(5)
print(driver.current_url)
inputElement = driver.find_element(By.NAME, 'phoneNumber')
inputElement.send_keys('555443211')
#inputElement.submit()

print('Подтвердить и перейти на следующую страницу? y/n')
move_to_2 = input()
if move_to_2 in ('y', 'Y', 'z', 'Z', 'н', 'Н'):
    continueButton = driver.find_element(By.XPATH, '//*[@id="etk_main_view"]/div/form/div/div[2]/div[2]/div/button')
    continueButton.submit()
time.sleep(2)
driver.close()
