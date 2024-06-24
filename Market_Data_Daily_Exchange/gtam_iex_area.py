import os
import glob
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from urllib.parse import urlparse, parse_qs
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
# from selenium.webdriver.support.ui import Select
import shutil
import time
from datetime import datetime


def run():
    try:
        options = Options()
        # options.add_argument("--headless") # Run selenium under headless mode
        prefs = {"download.default_directory": r"C:\GNA\Market Data\test"}
        options.add_experimental_option("prefs", prefs)
        driver = webdriver.Chrome(options=options)
        driver.get("https://www.iexindia.com/marketdata/G-TAM_Details.aspx")
    except Exception as e:
        return "Webpage Not Found"

    try:
        # Click on update button
        update = driver.find_element(By.ID, "ctl00_InnerContent_btnUpdateReport")
        update.click()
    except Exception as e:
        return "Date Not Selected"
    
    try:
        try:
            time.sleep(5)
            # Wait till data is available
            driver.find_element(
                By.XPATH,
                "/html/body/form/div[3]/section[2]/div/div/div/span[2]/div/table/tbody/tr[5]/td[3]/div/div[1]/div/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr/td[1]/table",
            )
        except Exception as e:
            return "Data Not Available"

        # Wait till download excel file button is present
        WebDriverWait(driver, 15).until(
            EC.presence_of_element_located(
                (
                    By.XPATH,
                    "/html/body/form/div[3]/section[2]/div/div/div/span[2]/div/table/tbody/tr[4]/td/div/div/div[4]/table/tbody/tr/td/div[2]/div[1]/a",
                )
            )
        )

        # Click on download button
        table = driver.find_element(
            By.ID, "ctl00_InnerContent_reportViewer_ctl05_ctl04_ctl00"
        )
        table.click()

        # Download excel file
        anchor = driver.find_element(
            By.XPATH,
            "/html/body/form/div[3]/section[2]/div/div/div/span[2]/div/table/tbody/tr[4]/td/div/div/div[4]/table/tbody/tr/td/div[2]/div[1]/a",
        )
        try:
            anchor.click()
        except Exception as e:
            print("error in html", e)
    except Exception as e:
        return "Excel File Download could Not Start"
    
    try:
        # Switch between tabs
        download_tab = driver.window_handles[1]
        driver.switch_to.window(download_tab)
    except Exception as e:
        return "Could Not Switch Tab"

    try:
        download_url = driver.current_url
        parsed_url = urlparse(download_url)
        parsed_query = parse_qs(parsed_url.query)
        print(parsed_query)
    except Exception as e:
        return "URL Could Not Parse"

    try:
        path_to_add = r"C:\GNA\Market Data\test"
        if not os.path.exists(path_to_add):
            os.makedirs(path_to_add)
        time.sleep(4)
        for file in glob.glob(os.path.join(path_to_add, "*.xlsx")):
            print(file)
        else:
            pass

        try:
            # Set Name for File
            new_file = (
                "GTAM IEX Area Price_" + datetime.now().strftime("%d.%m.%y") + ".xlsx"
            )
        except Exception as e:
            return "File Name Not Changed"

        try:
            # Make Copy of File to New Folder
            local_path = r"C:\GNA\Market Data\All Data"
            if not os.path.exists(local_path):
                os.makedirs(local_path)
            shutil.copyfile(
                os.path.join(path_to_add, file), rf"{local_path}/{new_file}"
            )
            time.sleep(2)

            file_store = r"C:\GNA\Market Data\GTAM All Exchanges"
            if not os.path.exists(file_store):
                os.makedirs(file_store)
            shutil.copyfile(
                os.path.join(path_to_add, file), rf"{file_store}/{new_file}"
            )
        except Exception as e:
            return "File Not Shifted To Local Path"
        finally:
            # Remove File from test Folder
            os.remove(os.path.join(path_to_add, file))

        for file in glob.glob(os.path.join(path_to_add, "*.xlsx.crdownload")):
            print("Error file found:", file)
        else:
            pass
    except Exception as e:
        return "Error in File !Try Again!"

    return "Success"


# run()
