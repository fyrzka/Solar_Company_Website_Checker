import pandas as pd
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
import time

# read the Excel file into a DataFrame
df = pd.read_excel(r'C:\Users\YiranFan\webscraping\website test.xlsx')

# create a new Chrome browser instance
driver = webdriver.Chrome()

# set the maximum time to wait for a page to load (in seconds)
timeout = 10

# iterate over the rows of the DataFrame
for index, row in df.iterrows():
    try:
        # navigate to the webpage
        driver.get(row['Company'])

        # Wait for the page to load
        WebDriverWait(driver, timeout).until(EC.presence_of_element_located((By.TAG_NAME, 'body')))

        # get the webpage's source code
        source_code = driver.page_source.lower()  # convert source code to lowercase

        # check for each keyword in the source code
        keywords = ['Hoymiles', 'Deye', 'APsystems', 'Huawei']
        for keyword in keywords:
            if keyword.lower() in source_code:  # convert keyword to lowercase
                df.loc[index, keyword] = 'Yes'
            else:
                df.loc[index, keyword] = ''
    except TimeoutException as e:
        print(f"Timeout: Page took too long to load - {e}")
        continue
    except Exception as e:
        print(f"Error accessing website: {e}")
        continue

# close the browser
driver.quit()

# write the DataFrame back to the Excel file
df.to_excel(r'C:\Users\YiranFan\webscraping\website test.xlsx', index=False)
