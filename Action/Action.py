from selenium import webdriver
from PageObject import LoginPage,HomePage,AddressPage

def login(useranme,password):
    '''封装的登录的函数'''
    driver = webdriver.Chrome(executable_path="d:\\chromedriver")
    driver.get("https://www.126.com")
    login_page = LoginPage.loginPage(driver)
    #login_page.click_login_link()
    login_page.switch_to_iframe()
    login_page.input_username(useranme)
    login_page.input_password(password)
    login_page.click_login_button()
    return driver

def add_person_info(driver,name,email,star_flag,mobile,comment):
    '''封装的添加一个联系人的函数'''

    home_page = HomePage.homePage(driver)
    home_page.click_address_link()
    address_page =AddressPage.addressPage(driver)
    address_page.click_create_button()
    address_page.input_name(name)
    address_page.input_email(email)
    address_page.input_comment(comment)
    address_page.click_star_flag(star_flag)
    address_page.input_mobile(mobile)
    address_page.click_save_button()

if __name__ =="__main__":
    driver = login("testgloryroad2020","123456789!!") #登录
    add_person_info(driver,"testgloryroad2020","123456789!!","吴老师","423333@qq.com","y","13343039393","随便写写")#新建联系人
