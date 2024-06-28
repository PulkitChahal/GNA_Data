import os
import time
from datetime import datetime, timedelta
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
import pytesseract as tess
tess.pytesseract.tesseract_cmd = r'C:\Users\pulki\AppData\Local\Programs\Tesseract-OCR\tesseract.exe'
from PIL import Image
import cv2

class wbes_gna_report:
    def __init__(self):
        self.main_directory = r'C:\GNA\Data'
        self.final_directory = r'C:\GNA\Data Upload'
        self.file_directory = os.path.join(f'{self.main_directory}\GNA_Report')
        self.clear_or_create_directory(self.file_directory)
        pass

    def create_directory(self, directory):
        if not os.path.exists(directory):
            os.makedirs(directory)

    def clear_or_create_directory(self, directory):
        if os.path.exists(directory):
            for file in os.listdir(directory):
                file_path_full = os.path.join(directory, file)
                if os.path.isfile(file_path_full):
                    os.remove(file_path_full)
        else:
            os.makedirs(directory)

    def download_gna_report_wbes(self):
        start_date = datetime.now().strftime('%d-%m-%Y')
        start_date_obj = datetime.strptime(start_date, '%d-%m-%Y')
        last_date = start_date_obj - timedelta(days=10)
        screenshot_path = os.path.join(self.file_directory, 'captcha.png')

        option = Options()
        prefs = {'download.default_directory': self.file_directory}
        option.add_experimental_option('prefs', prefs)
        chrome_driver_path = r'C:\Users\pulki\.cache\selenium\chromedriver\win64\125.0.6422.76\chromedriver.exe'
        service = Service(executable_path=chrome_driver_path)
        driver = webdriver.Chrome(service=service, options=option)

        # while start_date_obj.date() != last_date.date():
        driver.get('https://wbes.srldc.in/Report/GNA')
        wait = WebDriverWait(driver, 15)

        # Wait for the overlay to disappear
        wait.until(EC.invisibility_of_element_located((By.CLASS_NAME, 'blockOverlay')))

        # Find date_field part on page
        date_field_text = wait.until(EC.presence_of_element_located((By.ID, 'datepicker')))
        date_field_text.click()
        date_field_text.clear()
        date_field_text.send_keys(start_date_obj.strftime('%d-%m-%Y'))

        # Select region part
        region_dropdown = wait.until(EC.presence_of_element_located((By.ID, 'ddlRegion')))
        select_region = Select(region_dropdown)
        select_region.select_by_value('1')
        time.sleep(2)

        # Click on show_data
        show_data_button = wait.until(EC.element_to_be_clickable((By.ID, 'btnShow')))
        show_data_button.click()
        time.sleep(5)

        # Wait for captcha dialog to appear
        captcha_element = wait.until(EC.presence_of_element_located((By.CLASS_NAME, 'btn-toolbar')))
        driver.execute_script("arguments[0].scrollIntoView();", date_field_text)
        captcha_element.screenshot(screenshot_path)
        time.sleep(2)

        img = cv2.imread(screenshot_path)
        config = ('-l eng --oem 1 --psm 3')
        text = tess.image_to_string(img, config=config)
        print(text + 'hello')
        text = text.split('\n')[0]
        print(text + 'hello')
        time.sleep(5)

        # Enter the captcha solution
        captcha_input = wait.until(EC.presence_of_element_located((By.ID, 'txtCaptcha')))
        captcha_input.send_keys(text)

        # Click on submit button
        submit_button = wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[6]/div[3]/div/button')))
        submit_button.click()
        time.sleep(20)

        # Move to the next date (if looping through multiple dates)
        # start_date_obj -= timedelta(days=1)

    def get_data(self):
        wbes_gna_report.download_gna_report_wbes(self)


if __name__ == '__main__':
    wbes_gna = wbes_gna_report()
    wbes_gna.get_data()
    pass
