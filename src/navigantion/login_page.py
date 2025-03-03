# navigation/login_page.py
import sys
import os
current_dir = os.path.dirname(os.path.abspath(__file__))
src_dir = os.path.join(current_dir, '..') 
sys.path.append(src_dir)

from selenium.webdriver.common.by import By
from navigantion.base_page import BasePage

class LoginPage(BasePage):
    def __init__(self, driver):
        super().__init__(driver)
        self.username_input = (By.ID, "user")
        self.password_input = (By.ID, "password")
        self.login_button = (By.ID, "loginButton")

    def login(self, username, password):
        """Realiza login no sistema."""
        self.enter_text(*self.username_input, username)
        self.enter_text(*self.password_input, password)
        self.click_element(*self.login_button)
