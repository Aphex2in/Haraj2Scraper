from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from alive_progress import alive_bar
import time
import keyboard
from bs4 import BeautifulSoup
import time
import openpyxl
import threading

url = 'https://haraj.com.sa/search/%D8%BA%D8%B1%D9%81%D8%A9%20%D9%84%D9%84%D8%A7%D9%8A%D8%AC%D8%A7%D8%B1'  # Replace this with the URL you have when you search for something.
# Set up the Chrome WebDriver
driver = webdriver.Chrome()

# Variables for keyboard input
stop_event = threading.Event()

# Function to check for Esc key press
def check_esc_key():
    global esc_pressed
    while not stop_event.is_set():
        if keyboard.is_pressed('esc'):
            stop_event.set()
            break

# Start the keyboard listener in a separate thread
keyboard_thread = threading.Thread(target=check_esc_key)
keyboard_thread.start()

try:
    print('NOTE: DO NOT CLOSE THE WINDOW MANUALLY. PRESS ESC HERE IF YOU HAVE FINISHED.')
    loop = 0
    Refresh = 50 # Change this to your desirable amount. for example: 50 x Scrolls will go for 15,000 pages.
    Scrolls = 300 # DO NOT CHANGE THIS. HARAJ LIMITS THE AMOUNT OF SCROLLS YOU DO AFTER THAT NO MORE DEALS WILL POP UP.
    page_source = ''
    with alive_bar(Refresh*Scrolls) as bar:
        while loop in range(Refresh) and not stop_event.is_set():
            driver.get(url)
            button = driver.find_element(By.CSS_SELECTOR, '[class*="my-[5px] flex cursor-pointer items-center justify-center whitespace-nowrap border disabled:opacity-60 bg-transparent text-text-primary border-text-primary hover:bg-background hover:shadow-1 rounded-[10px] h-[40px] py-[10px] px-[15px]  !h-auto !border-gray-300 !py-3 !px-6 !text-elementary dark:!border-secondary-input-gray dark:!text-secondary md:!py-4 md:!px-8"]')
            button.click()
            # Scroll down to trigger loading of additional data
            for i in range(Scrolls):  # Adjust the number of scrolls as needed
                #print(str(i)+'/'+str(50), end='\r')
                driver.find_element(By.TAG_NAME, 'body').send_keys(Keys.END)
                time.sleep(1)  # Wait for the content to load, adjust as needed
                bar()
                if stop_event.is_set() is True:
                    break
            loop = loop+1
            page_source += driver.page_source

    # Extract the page source after scrolling
    soup = BeautifulSoup(page_source, 'html.parser')

    items = soup.find_all('div')

    # Extract text from each span element inside the parent divs and store it in a list
    prices = []
    titles = []
    cities = []
    for item in items:
        title = item.find('div', class_='box-border items-center overflow-hidden text-ellipsis')
        try:
            price = title.find_next_sibling('div', class_='flex w-[100%] content-center self-end mb-[5px]')
            if price:
                h2_elements = title.find('h2', class_='overflow-hidden text-ellipsis')
                if h2_elements not in titles:
                    titles.append(h2_elements)
                    price_value = price.find('span')
                    prices.append(price_value)
                    city = item.find('a', href=lambda href: href and href.startswith('/city/'))
                    cities.append(city)
        except:
            continue

    # Save the data to an Excel file
    excel_filename = 'scraped_data.xlsx'
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.append(['Title', 'Price', 'City'])  # Write the header..


    # Write data to the Excel file
    for title, price, city in zip(titles, prices, cities):
        sheet.append([title.text.strip(), price.text.strip(), city.text.strip()])

    workbook.save(excel_filename)

    print(f'Data has been saved to {excel_filename}')
    
finally:
    # Close the WebDriver
    stop_event.set()
    keyboard_thread.join()
    driver.quit()
