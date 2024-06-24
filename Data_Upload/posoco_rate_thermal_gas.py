import os
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import requests

class thermal_rate:

    def __init__(self):
        self.file_directory = r'C:/GNA/Data/Posoco_thermal_rate'
        if not os.path.exists(self.file_directory):
            os.makedirs(self.file_directory)

    def download_data_for_thermal_rate(self):
        options = Options()
        prefs = {'download.default_directory': self.file_directory}
        options.add_experimental_option('prefs', prefs)
        chromedriver_path = r'C:\Users\pulki\.cache\selenium\chromedriver\win64\125.0.6422.76\chromedriver.exe'
        driver = webdriver.Chrome(executable_path=chromedriver_path, options=options)
        # driver = webdriver.Chrome(options=options)

        driver.get('https://posoco.in/en/reports/ancillary-services-monthly-reports/tras-provider-details/')
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "ctl00_InnerContent_grdreport")))



if __name__ == '__main__':
    thermal_rate = thermal_rate()
    thermal_rate.download_data_for_thermal_rate()
    pass
