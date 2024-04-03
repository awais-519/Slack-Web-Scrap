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
    if not 0 < user_input <= 31: 
        exit()

    # URL to scrape
    url = 'https://app.slack.com/client/T5BBJGFRT/C0319QZCNDR'

    # Load the page
    driver.get(url)

    time.sleep(5)

    # Define the date according to user
    user_days = (datetime.now() - timedelta(days=user_input)).date()
    
    data = {}
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
        

        loop_break = False

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
                
                
                if sender_name not in data:
                    data[sender_name] = {}
                if message_date not in data[sender_name]:
                    data[sender_name][message_date] = set()  # Use a set to avoid duplicates
                
                data[sender_name][message_date].add((message_time, message_text))
                
                print(message_date, user_days)
                if message_date < user_days:
                    loop_break = True
                    break
        if loop_break:
            break
        time.sleep(1)

    # Create a DataFrame
    final_data = []
    for sender, dates in data.items():
        for date, messages in dates.items():
                sorted_messages = sorted(messages, key=lambda x: x[0])  # Sort messages based on time
                formatted_messages = [f"{text} @ {time}" for time, text in sorted_messages]
                final_data.append([sender, date, ' '.join(formatted_messages)])

    df = pd.DataFrame(final_data, columns=['Sender', 'Message Date', 'Data'])
    df = df.sort_values(by=['Message Date'], ascending=False)

    # Export DataFrame to Excel
    df.to_excel(r'C:\Users\awaish\OneDrive\Desktop\slack_messages_per_person.xlsx', index=False)

    print("Export successful.")
    
except Exception as e:
    print("An error occurred:", str(e))
finally:
    # Close the browser
    driver.quit()
