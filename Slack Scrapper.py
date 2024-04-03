from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import pandas as pd
import time
from datetime import datetime, timedelta
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By


# Path to your Chrome profile
chrome_profile_path = 'C:\\Users\\awaish\\AppData\\Local\\Google\\Chrome\\User Data\\Default'

# Configure Chrome options
chrome_options = Options()
chrome_options.add_argument('--user-data-dir=' + chrome_profile_path)

# Initialize Chrome driver
driver = webdriver.Chrome(options=chrome_options)

try:
    
    user_input = int(input("How many days you want to scrap? Enter in 1 to 31: "))
    if(user_input > 31 | user_input < 0): 
        exit
        
        
    # URL to scrape
    url = 'https://app.slack.com/client/T5BBJGFRT/C0319QZCNDR'

    # Load the page
    driver.get(url)

    time.sleep(5)

    # Define the date according to user
    user_days = (datetime.now() - timedelta(days = user_input)).date()
    
    data = []

    # Scroll to load more messages until the last three days' messages are reached
    while True:
        # Get page source
        page_source = driver.page_source

        # Parse the page source using BeautifulSoup
        soup = BeautifulSoup(page_source, 'html.parser')

        # Find all elements containing sender name, message text, and message time
        message_containers = soup.find_all("div", class_="c-message_kit__gutter c-message_kit__gutter--compact")

        driver.find_element(By.TAG_NAME, 'body').send_keys(Keys.PAGE_UP)
        time.sleep(1)
        loop_break = 0

        # Extract sender name, message text, message time, and date from each container
        for container in message_containers:
            date_elem = container.find("a", class_="c-link c-timestamp")
            sender_name_elem = container.find("button", class_="c-link--button c-message__sender_button")
            message_text_elem = container.find("div", class_="p-rich_text_section")
            time_elem = container.find("span", class_="c-timestamp__label")
            
            if sender_name_elem and message_text_elem and time_elem and date_elem:
                sender_name = sender_name_elem.text.strip()
                message_text = message_text_elem.text.strip()
                message_time = time_elem.text.strip()
                message_date = float(date_elem['data-ts'])  # Extracting aria-label attribute
                message_date = datetime.fromtimestamp(message_date).date()

            data.append([sender_name, message_text, message_date, message_time ])

            if message_date < user_days:
                loop_break = 1
                break

        if(loop_break == 1): 
            break
        time.sleep(1)



    # Create a DataFrame
    df = pd.DataFrame(data, columns=['Sender', 'Message Text', 'Message Date', 'Message Time' ])
    df = df.drop_duplicates()
    df = df.sort_values(by=['Message Date', 'Message Time'], ascending=False)

    # Export DataFrame to Excel
    df.to_excel(r'C:\Users\awaish\OneDrive\Desktop\slack_messages.xlsx', index=False)

    print("Export successful.")
    
except Exception as e:
    print("An error occurred:", str(e))
finally:
    # Close the browser
    driver.quit()