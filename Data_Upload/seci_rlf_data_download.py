import os
import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select


class seci_tender_results:
    def __init__(self):
        self.main_directory = r'C:/GNA/Data/seci_tender_results'
        self.clear_or_create_directory(self.main_directory)
        self.file_directory = r'C:/GNA/Data/seci_tender_results/download_files'
        self.clear_or_create_directory(self.file_directory)
        pass

    def clear_or_create_directory(self, directory):
        if os.path.exists(directory):
            for file in os.listdir(directory):
                file_path_full = os.path.join(directory, file)
                if os.path.isfile(file_path_full):
                    os.remove(file_path_full)
        else:
            os.makedirs(directory)

    def download_seci_tender_results(self):
        options = Options()
        prefs = {'download.default_directory': self.file_directory}
        options.add_experimental_option('prefs', prefs)
        chromedriver_path = r'C:\Users\pulki\.cache\selenium\chromedriver\win64\125.0.6422.76\chromedriver.exe'
        service = Service(executable_path=chromedriver_path)
        driver = webdriver.Chrome(service=service, options=options)
        url = 'https://www.seci.co.in/Bidder/view/tender/results/all-award/list/bidder'
        driver.get(url)
        wait = WebDriverWait(driver, 15)

        find_label = wait.until(EC.presence_of_element_located((By.XPATH, '/html/body/form/div['
                                                                          '3]/section/div/div/div/div/div/div[1]/div['
                                                                          '1]/div/label/select')))
        find_label.click()
        # Select value 100 from the dropdown
        select = Select(find_label)
        select.select_by_value('100')

        table = wait.until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[3]/section/div/div/div/div/div/div[2]/div')))
        links = wait.until(EC.presence_of_element_located((By.TAG_NAME,'a')))
        print(links)

        # Print the text of each link
        for link in links:
            print(link.text)


        time.sleep(20)


if __name__ == '__main__':
    seci_results = seci_tender_results()
    seci_results.download_seci_tender_results()
    pass
