import os
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import requests


class posoco_thermal_rate:

    def __init__(self):
        self.file_directory = r'C:/GNA/Data/Posoco_thermal_rate'
        if not os.path.exists(self.file_directory):
            os.makedirs(self.file_directory)

    def download_url(self, url, destination):
        # Download the file from the given URL and save it to the destination
        response = requests.get(url)
        with open(destination, 'wb') as f:
            f.write(response.content)

    def download_data_for_thermal_rate(self):
        # Set up Chrome options to specify the download directory
        options = Options()
        prefs = {'download.default_directory': self.file_directory}
        options.add_experimental_option('prefs', prefs)

        # Path to the ChromeDriver executable
        chromedriver_path = r'C:\Users\pulki\.cache\selenium\chromedriver\win64\125.0.6422.76\chromedriver.exe'

        # Initialize the Chrome WebDriver with the service and options
        service = Service(executable_path=chromedriver_path)
        driver = webdriver.Chrome(service=service, options=options)

        # Open the target URL
        driver.get('https://posoco.in/en/reports/ancillary-services-monthly-reports/tras-provider-details/')

        # Wait until the table with the specified ID is present on the page
        table = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.ID, "wpdmmydls-cb3e24c425e66d2d1f886e662b2f8bc6")))

        # Find all links within the table
        links = table.find_elements(By.TAG_NAME, 'a')

        # Iterate through each link and download the file
        for link in links:
            href = link.get_attribute('href')
            filename = link.text
            print(f"Downloading: {filename}")
            try:
                # Try downloading as a PDF
                destination = f"{self.file_directory}/{filename}.xlsx"
                self.download_url(href, destination)
            except:
                try:
                    # If PDF download fails, try downloading as an Excel file
                    destination = f"{self.file_directory}/{filename}.pdf"
                    self.download_url(href, destination)
                except:
                    # If both attempts fail, skip the file
                    print(f"Failed to download: {filename}")


class posoco_gas_rate:
    def __init__(self):
        self.file_directory = r'C:/GNA/Data/Posoco_gas_rate'
        if not os.path.exists(self.file_directory):
            os.makedirs(self.file_directory)
        pass

    def download_url(self, url, destination):
        response = requests.get(url)
        with open(destination, 'wb') as f:
            f.write(response.content)

    def download_data_for_gas_rate(self):
        options = Options()
        prefs = {'download.default_directory': self.file_directory}
        options.add_experimental_option('prefs', prefs)
        chromedriver_path = r'C:\Users\pulki\.cache\selenium\chromedriver\win64\125.0.6422.76\chromedriver.exe'
        service = Service(executable_path=chromedriver_path)
        driver = webdriver.Chrome(service=service, options=options)
        driver.get('https://posoco.in/en/reports/ancillary-services-monthly-reports/tras-provider-details-gas/')

        # Wait until the table with the specified ID is present on the page
        table = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.ID, "wpdmmydls-b5d7e336be749724eb02fa1371963c12")))

        # Find all links within the table
        links = table.find_elements(By.TAG_NAME, 'a')

        # Iterate through each link and download the file
        for link in links:
            href = link.get_attribute('href')
            filename = link.text
            print(f"Downloading: {filename}")
            try:
                # Try downloading as a PDF
                destination = f"{self.file_directory}/{filename}.xlsx"
                self.download_url(href, destination)
            except:
                try:
                    # If PDF download fails, try downloading as an Excel file
                    destination = f"{self.file_directory}/{filename}.pdf"
                    self.download_url(href, destination)
                except:
                    # If both attempts fail, skip the file
                    print(f"Failed to download: {filename}")


if __name__ == '__main__':
    thermal_rate = posoco_thermal_rate()
    gas_rate = posoco_gas_rate()
    # thermal_rate.download_data_for_thermal_rate()
    gas_rate.download_data_for_gas_rate()
    pass
