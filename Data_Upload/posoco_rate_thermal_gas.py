import os
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import requests
import pandas as pd
import tabula


class posoco_thermal_rate:

    def __init__(self):
        self.main_directory = r'C:/GNA/Data/Posoco_thermal_rate'
        self.file_directory = os.path.join(self.main_directory, '\download_files')
        self.clear_or_create_directory(self.file_directory)
        self.output_directory = os.path.join(f'{self.main_directory}\edited_xlsx_files')
        self.clear_or_create_directory(self.output_directory)
        pass

    def clear_or_create_directory(selfself, directory):
        if os.path.exists(directory):
            for file in os.listdir(directory):
                full_file_path = os.path.join(directory, file)
                if os.path.isfile(full_file_path):
                    os.remove(full_file_path)
        else:
            os.makedirs(directory)

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

    def edit_pdf_files(self):
        pdf_files = []
        for file in os.listdir(self.main_directory):
            if file.endswith('.pdf') and file.__contains__('Rate'):
                pdf_files.append(file)

        for pdf_file in pdf_files:
            input_file = os.path.join(self.main_directory, pdf_file)
            output_file = os.path.join(self.output_directory, os.path.splitext(pdf_file)[0] + '.xlsx')
            print(pdf_file)
            tables = tabula.read_pdf(input_file, pages='all')
            print(tables)
            modified_tables = []
            for i, table in enumerate(tables):
                df = pd.DataFrame(table)
                if i != -1:
                    df.loc[-1] = df.columns
                    df.index = df.index + 1
                    df = df.sort_index()
                    df.columns = range(len(df.columns))

                    modified_tables.append(df)

            merged_df = pd.concat(modified_tables,ignore_index=True)
                #     merged_df._append(merged_df, df)
            print(merged_df)

            merged_df.to_excel(output_file, index = False)
            print(f'File Saved at {output_file}')

    def edit_xlsx_files(self):
        xlsx_files = []
        for file in os.listdir(self.main_directory):
            if file.endswith('.xlsx') and file.__contains__('Rate'):
                xlsx_files.append(file)

        for xlsx_file in xlsx_files:
            input_file = os.path.join(self.main_directory, xlsx_file)
            output_file = os.path.join(self.output_directory, os.path.splitext(xlsx_file)[0] + '.xlsx')
            print(xlsx_file)

    def get_data(self):
        # posoco_thermal_rate.download_data_for_thermal_rate(self)
        posoco_thermal_rate.edit_pdf_files(self)


class posoco_gas_rate:
    def __init__(self):
        self.main_directory = r'C:/GNA/Data/Posoco_gas_rate'
        self.file_directory = os.path.join(self.main_directory, '\download_files')
        self.clear_or_create_directory(self.file_directory)
        self.output_directory = os.path.join(self.main_directory, '\edited_xlsx_files')
        self.clear_or_create_directory(self.output_directory)
        pass

    def clear_or_create_directory(selfself, directory):
        if os.path.exists(directory):
            for file in os.listdir(directory):
                full_file_path = os.path.join(directory, file)
                if os.path.isfile(full_file_path):
                    os.remove(full_file_path)
        else:
            os.makedirs(directory)

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

    def edit_xlsx_files(self):
        xlsx_files = []
        for file in os.listdir(self.main_directory):
            if file.endswith('.xlsx') and file.__contains__('rates'):
                xlsx_files.append(file)

        for xlsx_file in xlsx_files:
            input_file = os.path.join(self.main_directory, xlsx_file)
            output_file = os.path.join(self.output_directory, os.path.splitext(xlsx_file)[0] + '.xlsx')
            print(xlsx_file)
            try:
                df = pd.read_excel(input_file)
                print(df.head)
            except:
                pass

    def get_data(self):
        posoco_gas_rate.download_data_for_gas_rate(self)
        posoco_gas_rate.edit_xlsx_files(self)


if __name__ == '__main__':
    thermal_rate = posoco_thermal_rate()
    thermal_rate.get_data()

    gas_rate = posoco_gas_rate()
    # gas_rate.get_data()
    pass
