from google.colab import files
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time, os, shutil
from pathlib import Path

base_path = "/content"

# clean existing content
for item in os.listdir(base_path):
    item_path = os.path.join(base_path, item)
    if os.path.isdir(item_path):
        shutil.rmtree(item_path)  # delete folder and its contents
    else:
        os.remove(item_path)       # delete individual file

print("Press Choose File and Upload multiple PDF files")
# colab's upload file selection menu
uploaded = files.upload()

# create folder named pdfs in the colab directory
folder_name = "pdfs"
os.makedirs(folder_name, exist_ok=True)

# move the uploaded file to the pdfs folder
for filename in uploaded.keys():
    os.rename(filename, os.path.join(folder_name, filename))

print(f"Files are now in folder: {folder_name}")
print(os.listdir(folder_name))

# === Folders ===
input_folder = "/content/pdfs"   # folder containing PDFs
download_dir = "/content/downloads"    # where converted files will be saved
Path(download_dir).mkdir(parents=True, exist_ok=True)

# Chrome options (headless) -> for debugging and not running on colab, remove this
chrome_options = Options()
chrome_options.add_argument("--headless=new")     # headless mode
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.add_argument("--window-size=1920,1080")

# tell Chrome where to save files without asking
prefs = {
    "download.default_directory": download_dir,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True,
}
chrome_options.add_experimental_option("prefs", prefs)

driver = webdriver.Chrome(options=chrome_options)

# Allow downloads in headless
try:
    driver.execute_cdp_cmd(
        "Page.setDownloadBehavior",
        {"behavior": "allow", "downloadPath": download_dir},
    )
except Exception:
    # Fallback for newer Chrome/CDP versions
    driver.execute_cdp_cmd(
        "Browser.setDownloadBehavior",
        {"behavior": "allow", "downloadPath": download_dir}
    )

# === Wait for the download to finish by detecting file is in directory===
def wait_for_downloads(path, timeout=180):
    end = time.time() + timeout
    # Take a snapshot of existing files before starting
    existing_files = set(os.listdir(path))

    while time.time() < end:
        files = set(os.listdir(path))
        new_files = files - existing_files  # detect newly added files
        # ignore directories
        new_files = {f for f in new_files if os.path.isfile(os.path.join(path, f))}
        # check for completed download
        for f in new_files:
            if not f.endswith(".crdownload") and f.lower().endswith((".xlsx", ".xls", ".csv")):
                return os.path.join(path, f)
        time.sleep(1)
    return None

# === Process all PDFs in the input folder ===
pdf_files = [f for f in os.listdir(input_folder) if f.lower().endswith(".pdf")]
print(f"Found {len(pdf_files)} PDF(s) to convert.")

for pdf in pdf_files:
    pdf_path = os.path.join(input_folder, pdf)
    driver.get("https://www.ilovepdf.com/pdf_to_excel")
    time.sleep(2)  # wait for page to load

    # Wait for the hidden input[type=file]
    file_input = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='file']"))
    )

    # Upload file
    file_input.send_keys(pdf_path)
    print(f"Uploaded: {pdf}")

    time.sleep(2)  # wait for upload

     # Wait for the "One sheet" button and click it
    one_sheet_btn = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[1]/div[2]/div[2]/div[5]/ul/li[1]/div"))
    )
    one_sheet_btn.click()


    # Click the convert button after upload
    convert_btn = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.ID, "processTask"))
    )
    convert_btn.click()


    # Wait for download button
    download_btn = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.ID, "pickfiles"))
    )
    download_btn.click()

    # Wait for download to finish
    downloaded = wait_for_downloads(download_dir, timeout=300)
    if downloaded:
        print(f"Downloaded file: {downloaded}")
    else:
        print(f"Download timed out for: {pdf}")

driver.quit()
print("All PDFs processed.")

# Zip the folder
shutil.make_archive("output", 'zip', "downloads")

# Download the zip file
files.download("output.zip")