from selenium import webdriver
import time
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import lyhiajaline_config as lc


def get_element(path_type, element_path, custom_time_out=None,):
    """
    :param path_type: тип локатора
    :param element_path: путь элемента
    :return:
    """
    __type = {
        'xpath': By.XPATH,
        'css': By.CSS_SELECTOR,
        'id': By.ID,
        'class': By.CLASS_NAME
    }
    if custom_time_out is not None:
        time_out = custom_time_out
    else:
        time_out = 20

    return WebDriverWait(driver, time_out).until(EC.presence_of_element_located((__type.get(path_type), element_path)))


url = r'https://etaotlus.politsei.ee/ltr/#!/login'

driver = webdriver.Chrome()
driver.get(url)
driver.implicitly_wait(300)
sec_wait = 20
wait = WebDriverWait(driver, sec_wait)

#работает, но не кликабельно
#smart_id_button = driver.find_element(By.XPATH, "//*[name()='use' and @*='#icon-smart-id']")
#<svg class="icon icon-smart-id"><use xlink:href="#icon-smart-id"></use></svg>


#<a class="c-tab-login__nav-link is-active" href="#" lang="en" data-tab="smart-id">
#    <span class="c-tab-login__nav-label">
#        <svg class="icon icon-smart-id"><use xlink:href="#icon-smart-id"></use></svg>
#        Smart-ID
#    </span>
#</a>

# кнопка sisene на первой страницу
sisene_button = driver.find_element(By.XPATH, '//*[@id="etk_main_view"]/section/div[1]/div/div/form/div/div/div/button')
sisene_button.click()

#переход на логин со smart id
choice_buttons = driver.find_elements(By.CLASS_NAME, "c-tab-login__nav-link")
choice_buttons[2].click()


minu_isikukood = driver.find_element(By.ID, 'sid-personal-code')
minu_isikukood.send_keys(lc.minu_isikukood)
minu_isikukood.send_keys(Keys.ENTER)
time.sleep(3)


#uus_taotlus
#uus_taotlus_button = wait.until(EC.element_to_be_clickable((By.XPATH,
#                                                            '//*[@id="main-top-container"]/div/button'))).click()
uus_taotlus_button = get_element(By.XPATH, '//*[@id="main-top-container"]/div/button',30).click()

tooandja_registrikood = driver.find_element(By.NAME, 'identifier')
tooandja_registrikood.send_keys(lc.tooandja_registrikood)

tooandja_registrikood.send_keys(Keys.ENTER)

tegevusala = wait.until(EC.element_to_be_clickable((By.TAG_NAME, 'select')))
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

#TODO проверить загрузку доверенности. выглядит странно.
#грузим доверенность, время на появление кнопки Kustuta до 600 секунд

get_element(By.CLASS_NAME, 'btn btn-danger btn-xs pull-right', 600)
#wait.until(EC.element_to_be_clickable((By.CLASS_NAME, 'btn btn-danger btn-xs pull-right')))


#print('Загрузи доверенность и нажми тут 1. Кнопку "продолжать" на сайте не нажимать, переход будет автоматом: ')
#jatka_button_choise = input()
#if jatka_button_choise == '1':
#    jatka_button = driver.find_element(By.XPATH, '//*[@id="etk_main_view"]/div/form/div/div[2]/div[2]/div/button')
#    jatka_button.click()


#переходим к загрузке данных человека
#TODO сделать базу в экселе и конвертировать ее в словарь.
database_from_excel = {'firstName': 'YURII',
                       'lastName': 'BENDER',
                       'personCode': '37308220092',
                       'birthCountry': 'Ukraina',
                       'citizenship': 'Ukraina',
                       'otherCitizenship': '',
                       'email': 'tanja.svcentre@gmail.com',
                       'phoneNumber': '+37253985689',
                       'basisOfStay': 'viisa alusel',
                       'addressType1': 'Eesti'}

#TODO загрузить фотографию

get_element(By.CLASS_NAME, 'btn btn-default btn-outline ng-scope', 600)
#wait.until(EC.element_to_be_clickable((By.CLASS_NAME, 'btn btn-default btn-outline ng-scope')))


tootaja_first_name = driver.find_elemet(By.NAME, 'firstName')
tootaja_first_name.send_keys(database_from_excel['firstName'])

tootaja_last_name = driver.find_element(By.NAME, 'lastName')
tootaja_last_name.send_keys(database_from_excel['lastName'])

tootaja_isikukood = driver.find_element(By.NAME, 'personCode')
tootaja_isikukood.send_keys(database_from_excel['personCode'])

#TODO дата рождения и пол

tootaja_countryBirth = wait.until(EC.element_to_be_clickable((By.NAME, 'birthCountry')))
tootaja_countryBirth.click()
Select(tootaja_countryBirth).select_by_visible_text(database_from_excel['birthCountry'])

tootaja_citizenship = wait.until(EC.element_to_be_clickable((By.NAME, 'citizenship')))
tootaja_citizenship.click()
Select(tootaja_citizenship).select_by_visible_text(database_from_excel['citizenship'])

tootaja_other_citizenship = wait.until(EC.element_to_be_clickable((By.NAME, 'otherCitizenship')))
tootaja_other_citizenship.click()
Select(tootaja_other_citizenship).select_by_visible_text(database_from_excel['otherCitizenship'])

tootaja_email = driver.find_element(By.NAME, 'email')
tootaja_email.send_keys(database_from_excel['email'])

tootaja_phone = driver.find_element(By.NAME, 'phoneNumber')
tootaja_phone.send_keys(database_from_excel['phoneNumber'])

tootaja_basisOfStay = wait.until(EC.element_to_be_clickable((By.NAME, 'basisOfStay')))
tootaja_basisOfStay.click()
Select(tootaja_basisOfStay).select_by_visible_text(database_from_excel['basisOfStay'])

eesti_adress_country = driver.find_element(By.ID, 'addressType1')
eesti_adress_country.click()

#TODO заполнить нормально
aadress_string = driver.find_element(By.XPATH, '//*[@id="geoservice"]/input')

tootaja_postal = driver.find_element(By.NAME, 'localPostalCode')

tootaja_doc = driver.find_element(By.NAME, 'TravelDoc')

tootaja_doc_country = wait.until(EC.element_to_be_clickable((By.NAME, 'travelDocumentIssueCountry')))
tootaja_doc_country.click()

tootaja_doc_date = driver.find_element(By.NAME, 'travelDocumentIssueDate')

tootaja_doc_until = driver.find_element(By.NAME, 'travelDocumentValidUntil')


#TODO загрузить паспорт

get_element(By.CLASS_NAME, 'btn btn-danger ng-scope', 600)
#wait.until(EC.element_to_be_clickable((By.CLASS_NAME, 'btn btn-danger ng-scope')))


tootaja_too_address = driver.find_element(By.XPATH, '//*[@id="geoservice"]/input')

tootaja_too_indeks = driver.find_element(By.NAME, 'localPostalCode')

tootaja_too_type = wait.until(EC.element_to_be_clickable((By.NAME, 'tyoe')))
tootaja_too_type.click()

tootaja_palk = driver.find_element(By.NAME, 'grossSalary')

tootaja_kuus = wait.until(EC.element_to_be_clickable((By.NAME, 'grossSalaryUnit')))
tootaja_kuus.click()

tootaja_occupation = wait.until(EC.element_to_be_clickable((By.NAME, 'occupation')))
tootaja_occupation.click()

#TODO загрузить договор
tootaja_contract = wait.until(EC.presence_of_element_located((By.NAME, 'contractDocument')))
tootaja_contract.send_keys(r'C:\Users\docha\OneDrive\Leka\PPA LTR\taotlus.pdf')

tootaja_start = driver.find_element(By.NAME, 'workStartDate')
