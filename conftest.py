import pytest
import openpyxl
import json
import time
from selenium import webdriver

@pytest.fixture
def setup_browser():
    driver = webdriver.Chrome('D:\Selenium Practices\chromedriver-win64\chromedriver.exe')
    driver.maximize_window()
    yield driver
    time.sleep(5)
    driver.quit()