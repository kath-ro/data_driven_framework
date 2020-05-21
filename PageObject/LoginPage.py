import time
from selenium import webdriver
from Util.ObjectMap import find_element,find_elements
from Util.ParseConfig import *
from Conf.ProjVar import PageElementLocator_file_path

class loginPage():

    def __init__(self,driver):
        self.driver = driver

    '''[126mail_login]
       loginPage.loginlink=xpath>//a[contains(text(),'密码登录')]
    '''
    def get_login_link(self):#获取登录页面
        link = read_ini_file_option(
            PageElementLocator_file_path,"126mail_login","loginPage.loginlink") #定位表达式密码登录loginPage.loginlink
        element = find_element(self.driver,link.split(">")[0],link.split(">")[1])
        return element

    def click_login_link(self):#点击登录按钮
        self.get_login_link().click()
        time.sleep(5)

    """loginPage.frame=xpath>//iframe[contains(@id,'x-URS-iframe')]"""
    def get_iframe(self):#获取iframe框
        link = read_ini_file_option(
            PageElementLocator_file_path,"126mail_login","loginPage.frame")
        element = find_element(self.driver,link.split(">")[0],link.split(">")[1])
        return element

    def switch_to_iframe(self):#切换iframe框
        iframe = self.get_iframe()
        self.driver.switch_to.frame(iframe)
        time.sleep(1)

    """loginPage.username=xpath>//input[@name='email']"""
    def get_username(self):#获取输入用户名框
        link = read_ini_file_option(
            PageElementLocator_file_path,"126mail_login","loginPage.username")
        element = find_element(self.driver,link.split(">")[0],link.split(">")[1])
        return element

    def input_username(self,username):#输入用户名框
        self.get_username().send_keys(username)

    def get_password(self):
        link = read_ini_file_option(
            PageElementLocator_file_path,"126mail_login","loginPage.password")
        element = find_element(self.driver,link.split(">")[0],link.split(">")[1])
        return element

    def input_password(self,password):
        self.get_password().send_keys(password)

    def get_login_button(self):
        link = read_ini_file_option(
            PageElementLocator_file_path, "126mail_login", "loginPage.loginbutton")
        element = find_element(self.driver, link.split(">")[0], link.split(">")[1])
        return element

    def click_login_button(self):
        button = self.get_login_button()
        button.click()
        time.sleep(5)
        self.driver.switch_to.default_content()



if __name__ == "__main__":
    driver = webdriver.Chrome(executable_path="d:\\chromedriver")
    driver.get("https://www.126.com")
    login_page = loginPage(driver)
    #login_page.click_login_link()
    login_page.switch_to_iframe()
    login_page.input_username("testgloryroad2020")
    login_page.input_password("123456789!!")
    login_page.click_login_button()
