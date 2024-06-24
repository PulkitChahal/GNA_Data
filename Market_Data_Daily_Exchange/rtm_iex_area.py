import os
import glob
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from urllib.parse import urlparse, parse_qs
import shutil
import time
from datetime import datetime
from selenium.webdriver.support.ui import Select
from datetime import timedelta


def run():
    try:
        options = Options()
        # options.add_argument("--headless") # Run selenium under headless mode
        prefs = {"download.default_directory": r"D:\test"}
        options.add_experimental_option("prefs", prefs)
        driver = webdriver.Chrome(options=options)
        driver.get("https://www.iexindia.com/marketdata/rtm_areaprice.aspx")
    except Exception as e:
        return "Webpage Not Found"

    try:
        # Select Delivery Period
        dropdown = Select(
            driver.find_element(
                By.XPATH,
                "/html/body/form/div[3]/section[2]/div/div/div/div[1]/div[1]/label[2]/select",
            )
        )
        your_value = "-1"
        dropdown.select_by_value(your_value)
        time.sleep(1)

        # Update Report
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
                "/html/body/form/div[3]/section[2]/div/div/div/span[3]/div/table/tbody/tr[5]/td[3]/div/div[1]/div/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr[1]/td[1]/table",
            )
        except Exception as e:
            return "Data Not Available"

        # Wait till download Excel file button is present
        WebDriverWait(driver, 15).until(
            EC.presence_of_element_located(
                (
                    By.XPATH,
                    "/html/body/form/div[3]/section[2]/div/div/div/span[3]/div/table/tbody/tr[4]/td/div/div/div[4]/table/tbody/tr/td/div[2]/div[1]/a",
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
            "/html/body/form/div[3]/section[2]/div/div/div/span[3]/div/table/tbody/tr[4]/td/div/div/div[4]/table/tbody/tr/td/div[2]/div[1]/a",
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
        path_to_add = r"D:\test"
        if not os.path.exists(path_to_add):
            os.makedirs(path_to_add)
        time.sleep(4)
        for file in glob.glob(os.path.join(path_to_add, "*.xlsx")):
            print(file)
        else:
            pass

        try:
            # Set Name for File
            prev_date = datetime.today() - timedelta(days=1)
            new_date = prev_date.strftime("%d.%m.%y")
            new_file = (
                "RTM IEX Area Price_" + datetime.now().strftime(new_date) + ".xlsx"
            )
        except Exception as e:
            return "File Name Not Changed"

        try:
            # Make Copy of File to New Folder
            local_path = r'D:/Market Data/All Data'
            if not os.path.exists(local_path):
                os.makedirs(local_path)
            shutil.copyfile(
                os.path.join(path_to_add, file), rf"{local_path}/{new_file}"
            )
            time.sleep(2)

            file_store = r"D:/Market Data/RTM Area Price IEX"
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
