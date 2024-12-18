import pytest
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

@pytest.mark.parametrize("name, email, message", [
    ("Aditi Sharma","aditi.sharma@gmail.com",True),
    ("Rajesh Gupta","rajesh.gupta@",False),
    ("Vikram Malhotra","vikram.malhotra@gmail.com",True),
    ("","priya.verma@gmail.com",False)
])
def test_submit(setup_browser, name, email, message):
    driver = setup_browser
    driver.get("https://qavalidation.com/demo-form/")
    
    driver.find_element(By.NAME, "g4072-fullname").send_keys(name)
    driver.find_element(By.NAME, "g4072-email").send_keys(email)
    driver.find_element(By.XPATH, "//button[@type='submit']").click()
    
    try:
        successMessage = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH,"//h4[@id='contact-form-success-header']"))
        )
    except:
        successMessage = False
    finally:
        assert bool(successMessage) == message
    