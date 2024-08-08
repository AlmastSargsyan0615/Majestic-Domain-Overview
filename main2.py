import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException, WebDriverException
import time
import json
import csv

# Load login details from user.json
with open('user.json') as file:
    login_details = json.load(file)
    start_number = login_details["start_number"]
    end_number = login_details["end_number"]
    url_login = login_details["url_login"]
    username = login_details["username"]
    userpassword = login_details["userpassword"]

# Initialize the Chrome driver
options = webdriver.ChromeOptions()
driver = webdriver.Chrome(options=options)

def login(driver):
    print("--------------------------------login now---------------------------------------------")
    driver.get(url_login)
    time.sleep(3)  # Adjust the wait time as necessary
    driver.find_element(By.ID, "amember-login").send_keys(username)
    driver.find_element(By.ID, "amember-pass").send_keys(userpassword)
    driver.find_element(By.XPATH, "//input[@type='submit' and @value='Login']").click()
    print("--------------------------------login end---------------------------------------------")
    time.sleep(3)  # Adjust the wait time as necessary

def open_and_extract():
    try:
        # Perform login
        login(driver)

        # Find and click the button (assuming it's the first button on the page after login)
        button = driver.find_element(By.TAG_NAME, 'button')
        ActionChains(driver).key_down(Keys.CONTROL).click(button).key_up(Keys.CONTROL).perform()
        time.sleep(2)  # Wait for the new tab to open

        # Switch to the new tab
        driver.switch_to.window(driver.window_handles[1])
    except Exception as e:
        print(f"An error occurred: {e}")

def navigate_in_new_tab(domain, writer):
    print(domain)
    try:
        # Construct the URLs based on the domain
        url1 = f'https://majestic.noxtools.com/reports/site-explorer?oq={domain}&IndexDataSource=F&q={domain}'
        url2 = f'https://majestic.noxtools.com/reports/site-explorer/link-profile?q={domain}&oq={domain}&scope=&IndexDataSource=F'
        
        # Navigate to the first URL in the new tab
        driver.get(url1)
        time.sleep(5)  # Wait for the page to load

        # Get some data from the first URL (e.g., the title of the page)
        try:
            TF = driver.find_element(By.ID, "trust_flow_chart").text
        except NoSuchElementException:
            TF = "N/A"
        print(f"Data from first URL in new tab: {TF}")

        try:
            CF = driver.find_element(By.ID, "citation_flow_chart").text
        except NoSuchElementException:
            CF = "N/A"
        print(f"Data from first URL in new tab: {CF}")
        
        # Navigate to the second URL in the new tab
        driver.get(url2)
        time.sleep(5)  # Wait for the page to load

        try:
            referring_domains = driver.find_element(By.XPATH, "/html/body/div[5]/div[2]/div/div/div[2]/div/div[2]/div[1]/span[2]").text
        except NoSuchElementException:
            referring_domains = "N/A"
        print(f"Data from second URL in new tab: {referring_domains}")

        try:
            TF_referring_domains = driver.find_element(By.XPATH, "/html/body/div[5]/div[2]/div/div/div[2]/div/div[2]/div[2]/span[2]").text
        except NoSuchElementException:
            TF_referring_domains = "N/A"
        print(f"Data from second URL in new tab: {TF_referring_domains}")

        # Write the data to the CSV file
        writer.writerow([domain, TF, CF, referring_domains, TF_referring_domains])

    except Exception as e:
        print(f"An error occurred: {e}")

# Load domains from Excel file
df = pd.read_excel('urls.xlsx')
print("DataFrame columns:", df.columns)

# Extract domains within the range of start_number and end_number
domains = df.iloc[start_number - 1:end_number].iloc[:, 0].tolist()  # Adjust for zero-based index

# Create CSV file with start_number and end_number in the filename
csv_filename = f"results_{start_number}_to_{end_number}.csv"
with open(csv_filename, mode='w', newline='') as file:
    writer = csv.writer(file)
    writer.writerow(['Domain', 'Trust Flow', 'Citation Flow', 'Referring Domains', 'TF Referring Domains'])

    # Login and open the initial tab
    open_and_extract()

    # Navigate to different URLs constructed from domains in the new tab and write data to CSV
    for domain in domains:
        navigate_in_new_tab(domain, writer)

# Close the driver
driver.quit()
