import time
from selenium import webdriver
from Util.ObjectMap import find_element,find_elements
from Util.ParseConfig import *
from Conf.ProjVar import PageElementLocator_file_path

class homePage():

    def __init__(self,driver):
        self.driver = driver

    def get_address_link(self):
        link = read_ini_file_option(
            PageElementLocator_file_path,"126mail_homePage","homePage.addressbook")
        element = find_element(self.driver,link.split(">")[0],link.split(">")[1])
        return element

    def click_address_link(self):
        link  = self.get_address_link()
        link.click()
        time.sleep(5)


