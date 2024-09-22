import pandas as pd
import importlib
from email_sender import send_email
from datetime import datetime
import time
import random
import db
from tkinter import messagebox

def get_scraper(url):
    if 'site_url.com' in url:
        scraper_module = importlib.import_module('scrapers.site')
        return scraper_module.RobolinkMarketScraper(url)
    else:
        from site_scraper import SiteScraper
        return SiteScraper(url)

def update_excel(read_ex, categories):
    excel_data = pd.read_excel(read_ex, sheet_name=None)

    total_links = sum([len(excel_data[category]['Link']) for category in categories if category in excel_data])

    if total_links > 50:
        messagebox.showwarning("Warning", f"Too many links! There are {total_links} links. Please select another file with a maximum of 50 links.")
        return False

    avg_sleep_time = (120 + 250) / 2
    total_estimated_time = (total_links * avg_sleep_time) / 60
    print(f"Estimated total time: {total_estimated_time:.2f} minutes")

    for category in categories:
        if category in excel_data:
            df = excel_data[category]
            linkler = df['Link']
            ids = df['ID']
            i = 0

            for link in linkler:
                current_date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                productID = ids[i]
                i += 1
                scraper = get_scraper(link)
                scraper.fetch()
                name = scraper.get_data()
                db.update_product(name, productID, current_date)
                print(f"{link} completed successfully.")

                a = random.uniform(120, 250)
                print(f"Sleeping for {a:.2f} seconds.")
                time.sleep(a)

                remaining_links = total_links - i
                remaining_time = (remaining_links * avg_sleep_time) / 60
                print(f"Estimated remaining time: {remaining_time:.2f} minutes")

            db.excel_table(category)

    print("Product information updated and saved successfully.")
    return True

def email_send(tomail):
    send_email(
        subject="Updated Product Information",
        body="Please find attached the updated product information.",
        attachment_path=f".\Links_output.xlsx",
        tomail=tomail
    )

try:
    def job(file, categories):
        update_excel(file, categories)
except Exception as e:
    print(f"An error occurred: {e}")
    input("Press Enter to continue:")