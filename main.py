from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import json
import time

def convert_excel_to_json(excel_file, json_file):

    # Read the Excel file into a DataFrame
    df = pd.read_excel(excel_file)

    # Convert DataFrame to JSON and save it to a file
    df.to_json(json_file, orient='records')


def send_whatsapp_messages_from_json(json_file, wait_time=10):
    with open(json_file, 'r') as file:
        data = json.load(file)

    # Set up Chrome WebDriver
    chrome_options = Options()
    chrome_options.add_argument("--disable-infobars")
    chrome_options.add_argument("--disable-extensions")
    #chrome_options.add_argument("--headless")  # uncomment this line if you want to not see the browser

    # Path to which your chromedriver.exe is located  
    driver_path = "C:/Users/DELL/Downloads/chromedriver-win64/chromedriver-win64/chromedriver.exe" # Change with your actual path
    driver = webdriver.Chrome(executable_path=driver_path, options=chrome_options)

    for entry in data:
        country_code = '+20' # Change according to your country :")
        whatsapp_number = f"{country_code}{entry['whatsapp_numbers_column_name']}" # Change with your column name

        # For every column in the xlsx file you want to add it to your message
        firstColumn = str(entry['firstColumn_name']) # Change with your column name
        secondColumn = str(entry['secondColumn_name']) # Change with your column name

        # Construct the message to be sent
        whatsapp_message = f"Hi, your firstColumn is {firstColumn}, and about the secondColumn is {secondColumn}"

        # Construct the URL with the phone number
        whatsapp_url = f"https://web.whatsapp.com/send/?phone=%2B{whatsapp_number}&text&type=phone_number&app_absent=0"

        # Open the chat URL
        driver.get(whatsapp_url)
        # Give time to load the chat
        time.sleep(10)

        # Locate the chat input field
        input_box = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//div[@title='Type a message']"))
        )

        # Type the message
        input_box.send_keys(whatsapp_message)

        # Press Enter to send the message
        input_box.send_keys(Keys.RETURN)

        # Wait for the specified time (e.g., 10 seconds)
        time.sleep(wait_time)

    # Close the browser after sending all messages
    driver.quit()



if __name__ == "__main__":

    # Provide the input Excel file and the desired output JSON file (empty file)
    # Must be with the same directory of the main.py file
    excel_file_path = 'test.xlsx' # change with your actual file name
    json_file_path = 'output.json' # change with your actual file name

    convert_excel_to_json(excel_file_path, json_file_path)
    send_whatsapp_messages_from_json(json_file_path)
