import os
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time
import random
import requests


def initialize_driver():
    """Initialize Selenium WebDriver."""
    options = webdriver.ChromeOptions()
    driver = webdriver.Chrome(options=options)  # Ensure ChromeDriver is in PATH
    return driver


def search_pdf(driver, query):
    """Search Google for a PDF URL based on the query."""
    try:
        driver.get("https://www.google.com")
        search_box = driver.find_element(By.NAME, "q")
        search_box.send_keys(query)
        search_box.send_keys(Keys.RETURN)
        time.sleep(3)  # Allow search results to load

        # Find all search result links
        results = driver.find_elements(By.CSS_SELECTOR, "a")
        for result in results:
            href = result.get_attribute("href")
            if href and href.endswith(".pdf"):  # Look for PDF links
                return href
    except Exception as e:
        print(f"Error during search: {e}")
    return None


def download_pdf(url, save_path):
    """Download a PDF from a URL and save it to the specified path."""
    try:
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
        }
        response = requests.get(url, headers=headers, stream=True)
        response.raise_for_status()
        os.makedirs(os.path.dirname(save_path), exist_ok=True)
        with open(save_path, "wb") as pdf_file:
            for chunk in response.iter_content(chunk_size=8192):
                pdf_file.write(chunk)
        print(f"Downloaded: {save_path}")
    except Exception as e:
        print(f"Failed to download {url}: {e}")
        return str(e)
    return None


def search_and_download_pdfs(file_path, output_folder, failed_log_file, results_log_file):
    """Search and download PDFs using application numbers from an Excel file."""
    try:
        os.makedirs(output_folder, exist_ok=True)
        failed_entries = []
        success_entries = []
        driver = initialize_driver()

        # Read Excel file
        data = pd.read_excel(file_path, usecols=["APPLICATION_NUMBER"])

        for _, row in data.iterrows():
            patent_no = str(row["APPLICATION_NUMBER"]).strip()

            # Check if APPLICATION_NUMBER is NaN
            if pd.isna(row["APPLICATION_NUMBER"]):
                print(f"Stopping the process due to NaN APPLICATION_NUMBER: {row}")
                break  # Stop the loop entirely

            # Search by APPLICATION_NUMBER
            query = f'"{patent_no}" filetype:pdf'
            print(f"Searching for: {query}")
            pdf_url = search_pdf(driver, query)
            if pdf_url:
                pdf_name = f"{patent_no}.pdf"
                save_path = os.path.join(output_folder, pdf_name)
                error = download_pdf(pdf_url, save_path)
                if not error:
                    success_entries.append({"APPLICATION_NUMBER": patent_no, "PDF_URL": pdf_url})
                else:
                    failed_entries.append({"APPLICATION_NUMBER": patent_no, "Error": error})
            else:
                print(f"No PDF found for: {patent_no}")
                failed_entries.append({"APPLICATION_NUMBER": patent_no, "Error": "PDF not found"})

            # Random delay to avoid being flagged
            delay = random.randint(2, 6)
            time.sleep(delay)

        driver.quit()  # Close the browser

        # Save failed entries to Excel
        if failed_entries:
            failed_df = pd.DataFrame(failed_entries)
            failed_df.to_excel(failed_log_file, index=False, sheet_name="Failed Entries")
            print(f"Failed entries saved to {failed_log_file}")
        else:
            print("No failed entries.")

        # Save successful results to Excel
        if success_entries:
            success_df = pd.DataFrame(success_entries)
            success_df.to_excel(results_log_file, index=False, sheet_name="Successful Downloads")
            print(f"Successful entries saved to {results_log_file}")
        else:
            print("No successful downloads.")

        print("Process completed.")
    except Exception as e:
        print(f"An error occurred: {e}")


# Usage
excel_file_path = "./Ayush Applications Filed and Granted 10092024.xlsx"  # Replace with your Excel file path
output_folder_path = "./downloaded_pdfs"  # Replace with your desired output folder
failed_log_path = "./Failed_Patent_Numbers.xlsx"  # File to store failed entries in Excel
results_log_path = "./Successful_Patent_Downloads.xlsx"  # File to store successful entries in Excel

search_and_download_pdfs(excel_file_path, output_folder_path, failed_log_path, results_log_path)
