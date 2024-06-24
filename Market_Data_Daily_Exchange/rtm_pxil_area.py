import os
import glob
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import shutil
import time
from datetime import datetime
from datetime import timedelta


def run():
    try:
        options = Options()
        # options.add_argument("--headless") # Run selenium under headless mode
        prefs = {"download.default_directory": r"D:\test"}
        options.add_experimental_option("prefs", prefs)
        driver = webdriver.Chrome(options=options)
        driver.get("https://powerexindia.in/Pages/Market.html#RTM/")
    except Exception as e:
        return "Webpage Not Found"

    market_spot = driver.find_element(
        By.XPATH, "/html/body/main/div/section[3]/div/div/div/div/ul/li[1]/a"
    )

    try:
        # Select Real Time Spot
        real_time = driver.find_element(
            By.XPATH,
            "/html/body/main/div/section[3]/div/div/div/div/div/div[2]/div/div[2]/div[1]/nav/div/a[4]",
        )
        driver.execute_script("arguments[0].scrollIntoView();", market_spot)
        time.sleep(2)
        real_time.click()
        time.sleep(1)
    except Exception as e:
        return "Area Clearing Price Not Selected"

    try:
        # Select Toggle Filter
        toggle_filter = driver.find_element(
            By.XPATH,
            "/html/body/main/div/section[3]/div/div/div/div/div/div[2]/div/div[2]/div[2]/div/div[2]/div/div/div[1]/button",
        )
        driver.execute_script("arguments[0].scrollIntoView();", real_time)
        time.sleep(2)
        toggle_filter.click()
        time.sleep(1)
    except Exception as e:
        return "Toggle Filter Not Selected"

    try:
        # Select Delivery Date Range
        delivery_range = driver.find_element(
            By.XPATH,
            "/html/body/main/div/section[3]/div/div/div/div/div/div[2]/div/div[2]/div[2]/div/div[2]/div/div/div[2]/form/div[1]/div/input",
        )
        delivery_range.click()
        time.sleep(2)

        # Select Delivery Date
        delivery_date = WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/div[18]/div[1]/ul/li[2]')))
        delivery_date.click()
        time.sleep(2)

        # Submit
        submit = driver.find_element(
            By.XPATH,
            "/html/body/main/div/section[3]/div/div/div/div/div/div[2]/div/div[2]/div[2]/div/div[2]/div/div/div[2]/form/div[2]/div/button",
        )
        submit.click()
        # time.sleep(2)
    except Exception as e:
        return "Date Not Selected"

    try:
        time.sleep(5)
        # Wait till Excel file is not clickable
        WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable(
                (
                    By.XPATH,
                    "/html/body/main/div/section[3]/div/div/div/div/div/div[2]/div/div[2]/div[2]/div/div[2]/div/div/div[6]/div/div/div[1]/div/button[2]",
                )
            )
        )

        # Download Excel File
        table = driver.find_element(
            By.XPATH,
            "/html/body/main/div/section[3]/div/div/div/div/div/div[2]/div/div[2]/div[2]/div/div[2]/div/div/div[6]/div/div/div[1]/div/button[2]",
        )
        driver.execute_script("arguments[0].scrollIntoView();", submit)
        time.sleep(2)
        try:
            driver.execute_script("arguments[0].click();", table)
        except Exception as e:
            print(e)
            pass
    except Exception as e:
        return "Excel File Download could Not Start"

    try:
        path_to_add = r"D:/test"
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
                    "RTM PXIL Area Price_" + datetime.now().strftime(new_date) + ".xlsx"
            )
        except Exception as e:
            return "File Name Not Changed"

        try:
            # Make Copy of File to New Folder
            local_path = r"D:/Market Data/All Data"
            if not os.path.exists(local_path):
                os.makedirs(local_path)
            shutil.copyfile(
                os.path.join(path_to_add, file), rf"{local_path}/{new_file}"
            )
            time.sleep(2)

            file_store = r"D:/Market Data/RTM Area Price PXIL"
            if not os.path.exists(file_store):
                os.makedirs(file_store)
            shutil.copyfile(
                os.path.join(path_to_add, file), rf"{file_store}/{new_file}"
            )
            time.sleep(2)
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