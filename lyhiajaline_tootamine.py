from selenium import webdriver
import time
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import lyhiajaline_config as lc


url = r'https://etaotlus.politsei.ee/ltr/#!/login'

driver = webdriver.Chrome()
driver.get(url)
#time.sleep(10)

#работает, но не кликабельно
#smart_id_button = driver.find_element(By.XPATH, "//*[name()='use' and @*='#icon-smart-id']")
#<svg class="icon icon-smart-id"><use xlink:href="#icon-smart-id"></use></svg>


#<a class="c-tab-login__nav-link is-active" href="#" lang="en" data-tab="smart-id">
#    <span class="c-tab-login__nav-label">
#        <svg class="icon icon-smart-id"><use xlink:href="#icon-smart-id"></use></svg>
#        Smart-ID
#    </span>
#</a>
time.sleep(3)
sisene_button = driver.find_element(By.XPATH, '//*[@id="etk_main_view"]/section/div[1]/div/div/form/div/div/div/button')
sisene_button.click()

time.sleep(3)
choice_buttons = driver.find_elements(By.CLASS_NAME, "c-tab-login__nav-link")
#WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH,
#        "//*[name()='svg:image' and starts-with(@class, 'holder') and contains(@xlink:href, 'some')]"))).click()

#smart_id_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH,
#        '//div[@class="icon icon-smart-id"]/*[name()="svg"]'))).click()


#print(choice_buttons)
#print('Перейти на smart-ID? Да - 1 ')
#go_to_smart_id = input('')
#if go_to_smart_id == '1':
#    choice_buttons[2].click()
choice_buttons[2].click()

minu_isikukood = driver.find_element(By.ID, 'sid-personal-code')
minu_isikukood.send_keys(lc.minu_isikukood)
minu_isikukood.send_keys(Keys.ENTER)
time.sleep(3)


#uus_taotlus
WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH,
                                                            '//*[@id="main-top-container"]/div/button'))).click()
#uus_taotlus_button = driver.find_element(By.CLASS_NAME, 'ng-binding ng-scope')
#uus_taotlus_button.click()

time.sleep(2)


tooandja_registrikood = driver.find_element(By.NAME, 'identifier')
tooandja_registrikood.send_keys(lc.tooandja_registrikood)

tooandja_registrikood.send_keys(Keys.ENTER)

tegevusala = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.TAG_NAME, 'select')))
#tegevusala = driver.find_element(By.TAG_NAME, 'select')
tegevusala.click()
Select(tegevusala).select_by_visible_text(lc.tegevusala)

firma_telefon = driver.find_element(By.NAME, 'phoneNumber')
firma_telefon.send_keys(lc.firma_telefon)

firma_email = driver.find_element(By.NAME, 'email')
firma_email.send_keys(lc.firma_email)

represent = driver.find_element(By.XPATH, '//*[@id="representativePerson0"]')
represent.click()

minu_telefon = driver.find_element(By.NAME, 'authorizedPersonPhoneNumber')
minu_telefon.send_keys(lc.minu_telefon)

minu_email = driver.find_element(By.NAME, 'authorizedPersonEmail')
minu_email.send_keys(lc.minu_email)

print('Загрузи доверенность и нажми тут 1. Кнопку "продолжать" на сайте не нажимать: ')
jatka_button_choise = input()
if jatka_button_choise == '1':
    jatka_button = driver.find_element(By.XPATH, '//*[@id="etk_main_view"]/div/form/div/div[2]/div[2]/div/button')
    jatka_button.click()


