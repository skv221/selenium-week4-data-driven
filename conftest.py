import pytest
from selenium import webdriver

@pytest.fixture
def setup_browser():
    driver = webdriver.Chrome('D:\Selenium Practices\chromedriver-win64\chromedriver.exe')
    driver.maximize_window()
    yield driver
    driver.quit()