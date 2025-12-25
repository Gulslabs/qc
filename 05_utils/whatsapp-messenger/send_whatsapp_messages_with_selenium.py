import time
import csv
import random
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager


def setup_driver():
    service = Service(ChromeDriverManager().install())
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    driver = webdriver.Chrome(service=service, options=options)
    driver.get('https://web.whatsapp.com')
    input("Scan QR Code and press ENTER to continue...")
    return driver


def send_messages(driver, contacts, message_template):
    sent_count = 0

    for contact in contacts:
        print(f"Processing contact: {contact}")
        try:
            number = contact.get('number')
            if not number:
                print("Skipping contact with missing number.")
                continue

            message = message_template.format(**contact)
            print(f"Sending about ({contact.get('Student_Name', 'Unknown')}) to ({number})...")

            driver.get(f"https://web.whatsapp.com/send?phone={number}")
            time.sleep(10)  # Wait for chat to load

            message_box = driver.find_element(By.XPATH, '//div[@contenteditable="true"][@data-tab="10"]')
            for line in message.split('\n'):
                message_box.send_keys(line)
                message_box.send_keys(Keys.SHIFT, Keys.ENTER)  # Line break
                
            message_box.send_keys(Keys.ENTER)

            print(f"Message sent to {contact.get('Student_Name', 'Unknown')} ({number})")
            sent_count += 1

            delay = random.randint(15, 20)
            print(f" Waiting {delay} seconds...")
            time.sleep(delay)

            if sent_count % 20 == 0:
                print("Sent 20 messages. Waiting 15 minutes to avoid blocking...")
                time.sleep(900)

        except Exception as e:
            print(f" Failed to send to {contact.get('Student_Name', 'Unknown')} ({number}): {e}")
            time.sleep(5)


def load_template(template_file):
    with open(template_file, 'r', encoding='utf-8') as file:
        return file.read()


def load_contacts(csv_file):
    contacts = []
    with open(csv_file, mode='r', newline='', encoding='utf-8') as file:
        reader = csv.DictReader(file)
        for row in reader:
            contacts.append(row)  # Use full row dictionary
    return contacts


def main():
    contacts = load_contacts("contacts.csv")
    message_template = load_template("message_template.txt")
    driver = setup_driver()
    send_messages(driver, contacts, message_template)
    print("All messages sent successfully!")
    driver.quit()


if __name__ == "__main__":
    main()
