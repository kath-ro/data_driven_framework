import time
from selenium import webdriver
from Util.ObjectMap import find_element,find_elements
from Util.ParseConfig import *
from Conf.ProjVar import PageElementLocator_file_path
from PageObject import  HomePage,LoginPage
class addressPage():

    def __init__(self,driver):
        self.driver = driver

    def get_create_button(self):
        button = read_ini_file_option(
        PageElementLocator_file_path,"126mail_addContactsPage","addContactsPage.createContactsBtn")
        element = find_element(self.driver,button.split(">")[0],button.split(">")[1])
        return element

    def click_create_button(self):
        button = self.get_create_button()
        button.click()
        time.sleep(1)

    def get_name(self):
        input_box = read_ini_file_option(
        PageElementLocator_file_path, "126mail_addContactsPage", "addContactsPage.contactPersonName")
        element = find_element(self.driver, input_box.split(">")[0], input_box.split(">")[1])
        return element

    def input_name(self,name):
        if name is None: return
        input_box = self.get_name()
        input_box.send_keys(name)

    def get_email(self):
        input_box = read_ini_file_option(
            PageElementLocator_file_path, "126mail_addContactsPage", "addContactsPage.contactPersonEmail")
        element = find_element(self.driver, input_box.split(">")[0], input_box.split(">")[1])
        return element

    def input_email(self, email):
        if email is None: return
        input_box = self.get_email()
        input_box.send_keys(email)

    def get_mobile(self):
        input_box = read_ini_file_option(
            PageElementLocator_file_path, "126mail_addContactsPage", "addContactsPage.contactPersonMobile")
        element = find_element(self.driver, input_box.split(">")[0], input_box.split(">")[1])
        return element

    def input_mobile(self, mobile):
        if mobile is None: return
        input_box = self.get_mobile()
        input_box.send_keys(mobile)

    def get_comment(self):
        input_box = read_ini_file_option(
            PageElementLocator_file_path, "126mail_addContactsPage", "addContactsPage.contactPersonComment")
        element = find_element(self.driver, input_box.split(">")[0], input_box.split(">")[1])
        return element

    def input_comment(self,comment):
        if comment is None: return
        input_box = self.get_comment()
        input_box.send_keys(comment)

    def get_save_button(self):
        button = read_ini_file_option(
            PageElementLocator_file_path, "126mail_addContactsPage", "addContactsPage.savecontacePerson")
        element = find_element(self.driver, button.split(">")[0], button.split(">")[1])
        return element

    def click_save_button(self):
        button = self.get_save_button()
        button.click()
        time.sleep(5)
        #self.driver.quit()

    def get_star_flag(self):
        check_box = read_ini_file_option(
            PageElementLocator_file_path, "126mail_addContactsPage", "addContactsPage.starContacts")
        element = find_element(self.driver, check_box.split(">")[0], check_box.split(">")[1])
        return element

    def click_star_flag(self,flag):
        if flag is None: return
        if ("y" in flag) or ("yes" in flag) or ("是" in flag):
            check_box = self.get_star_flag()
            check_box.click()

if __name__ == "__main__":
    driver = webdriver.Chrome(executable_path="d:\\chromedriver")
    driver.get("https://www.126.com")
    login_page = LoginPage.loginPage(driver)
    login_page.click_login_link()
    login_page.switch_to_iframe()
    login_page.input_username("testgloryroad2020")
    login_page.input_password("123456789!!")
    login_page.click_login_button()

    home_page = HomePage.homePage(driver)
    home_page.click_address_link()

    address_page = addressPage(driver)
    address_page.click_create_button()
    address_page.input_name("吴老师")
    address_page.input_email("2055739@qq.com")
    address_page.input_comment("随便写写")
    address_page.click_star_flag("是")
    address_page.input_mobile("123444444444")
    address_page.click_save_button()
