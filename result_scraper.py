import requests
from bs4 import BeautifulSoup
from PIL import Image
import pytesseract
import io
import re
import os
import time
import csv
import tempfile
import random
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from concurrent.futures import ThreadPoolExecutor, as_completed

# Explicitly set the path to tesseract.exe
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

# --- Configuration ---
BASE_URL = "https://jcboseustymca.co.in/Forms/Student/ResultStudents.aspx"
RESULT_URL = "https://jcboseustymca.co.in/Forms/Student/PrintReportCardNew.aspx"
CAPTCHA_URL = "https://jcboseustymca.co.in/Handler/GenerateCaptchaImage.ashx"
OUTPUT_DIR = "results"
MAX_WORKERS = 4 # A safer number of concurrent threads to avoid server issues

# --- Captcha functions (unchanged) ---
def clean_captcha(text):
    text = text.strip().upper()
    text = re.sub(r'[^A-Z0-9]', '', text)
    text = text.replace("0", "O").replace("1", "I").replace("5", "S").replace("8", "B")
    return text

def solve_captcha(img_bytes):
    try:
        img = Image.open(io.BytesIO(img_bytes)).convert("L")
        img = img.point(lambda x: 0 if x < 140 else 255)
        text = pytesseract.image_to_string(
            img, config="--psm 7 -c tessedit_char_whitelist=ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
        )
        return clean_captcha(text)
    except Exception as e:
        print(f"  [!] Error during OCR: {e}")
        return ""

# --- MODIFIED: Screenshot function with ZOOM capability ---
def save_html_as_image(html_content, output_path):
    """Renders HTML, zooms out, and saves a full-height screenshot."""
    print(f"  📸 Capturing screenshot to {output_path}...")
    try:
        options = webdriver.ChromeOptions()
        options.add_argument("--headless")
        options.add_argument("--window-size=1200,800")
        options.add_argument("--hide-scrollbars")
        options.add_argument("--log-level=3")
        options.add_experimental_option('excludeSwitches', ['enable-logging'])

        with tempfile.NamedTemporaryFile(delete=False, suffix=".html", mode='w', encoding='utf-8') as fp:
            fp.write(html_content)
            temp_path = "file://" + os.path.abspath(fp.name).replace('\\', '/')

        driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=options)
        driver.get(temp_path)
        time.sleep(2)

        # --- NEW: Replicate the "Zoom Out" to 70% ---
        driver.execute_script("document.body.style.zoom='70%'")
        print("    Zoomed out to 70% for full view.")
        time.sleep(2) # Allow zoom to apply before measuring height

        js_get_height = "return Math.max( document.body.scrollHeight, document.body.offsetHeight, document.documentElement.clientHeight, document.documentElement.scrollHeight, document.documentElement.offsetHeight );"
        total_height = driver.execute_script(js_get_height)
        
        driver.set_window_size(1200, total_height + 50)
        time.sleep(2)

        driver.save_screenshot(output_path)
        driver.quit()
        os.unlink(fp.name)
        print("  ✅ Screenshot saved successfully.")
        return True
    except Exception as e:
        print(f"  [!] Failed to save screenshot: {e}")
        return False

# --- Stricter parsing function (unchanged) ---
def parse_result_details(soup):
    try:
        name_span = soup.find('span', {'id': 'lblname'})
        if not name_span or not name_span.text.strip():
            return None

        name = name_span.text.strip()
        sgpa_span = soup.find('span', {'id': 'lblResult'})
        sgpa = sgpa_span.text.strip() if sgpa_span else "N/A"
        cgpa_span = soup.find('span', {'id': 'lblCgpaResult'})
        cgpa = cgpa_span.text.strip() if cgpa_span and cgpa_span.text.strip() else "N/A"
        
        return {"name": name, "sgpa": sgpa, "cgpa": cgpa}

    except AttributeError:
        return None

# --- Main result fetching logic (unchanged) ---
def fetch_result(session, roll_number, semester):
    max_attempts = 15
    for attempt in range(max_attempts):
        print(f"  Attempt {attempt + 1}/{max_attempts} for {roll_number}...")
        try:
            response = session.get(BASE_URL)
            soup = BeautifulSoup(response.text, "html.parser")

            viewstate = soup.find("input", {"id": "__VIEWSTATE"})["value"]
            viewstategen = soup.find("input", {"id": "__VIEWSTATEGENERATOR"})["value"]
            eventvalidation = soup.find("input", {"id": "__EVENTVALIDATION"})["value"]
            captcha_img_bytes = session.get(CAPTCHA_URL).content
            captcha_text = solve_captcha(captcha_img_bytes)

            if len(captcha_text) < 5: continue

            payload = {
                "__VIEWSTATE": viewstate, "__VIEWSTATEGENERATOR": viewstategen,
                "__EVENTVALIDATION": eventvalidation, "txtRollNo": roll_number,
                "ddlSem": semester.zfill(2), "txtCaptcha": captcha_text,
                "btnResult": "View Result",
            }

            session.post(BASE_URL, data=payload, timeout=20)
            report_page = session.get(RESULT_URL, timeout=20)
            soup_result = BeautifulSoup(report_page.text, "html.parser")
            
            parsed_data = parse_result_details(soup_result)

            if parsed_data:
                print(f"  🎉 Success! Found Result for {roll_number}")
                return {"status": "success", "html": report_page.text, "details": parsed_data}
            else:
                print(f"    Login failed for {roll_number}. Retrying...")
                time.sleep(2)
        except Exception as e:
            print(f"  [!] Error for {roll_number}: {e}. Retrying...")
            time.sleep(5)
            
    return {"status": "failed", "html": "", "details": {"name": "N/A", "sgpa": "N/A", "cgpa": "N/A"}}

# --- MODIFIED: Worker function with a stagger delay ---
def process_roll_number(roll_number, semester):
    """Complete, isolated process for one roll number with a delay."""
    # --- NEW: Stagger requests to avoid server race conditions ---
    time.sleep(random.uniform(0.5, 2.0))
    
    print(f"[*] Starting job for Roll Number: {roll_number}")
    with requests.Session() as session:
        result = fetch_result(session, roll_number, semester)
    
    result['roll_number'] = roll_number

    if result['status'] == 'success':
        roll_dir = os.path.join(OUTPUT_DIR, roll_number)
        os.makedirs(roll_dir, exist_ok=True)
        image_path = os.path.join(roll_dir, f"result_sem_{semester}.png")
        save_html_as_image(result["html"], image_path)

    return result

# --- Main execution block (MULTITHREADED) ---
if __name__ == "__main__":
    roll_no_file = input("Enter the path to the .txt file containing roll numbers: ")
    semester = input("Enter the semester for all roll numbers (e.g., 1, 2, 03...): ")

    if not os.path.exists(roll_no_file):
        print(f"Error: The file '{roll_no_file}' was not found.")
        exit()

    with open(roll_no_file, 'r') as f:
        roll_numbers = [line.strip() for line in f if line.strip()]

    print(f"\nFound {len(roll_numbers)} roll numbers. Starting processing with {MAX_WORKERS} workers.")
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    all_results = []

    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        future_to_roll = {executor.submit(process_roll_number, roll, semester): roll for roll in roll_numbers}
        for future in as_completed(future_to_roll):
            try:
                result_data = future.result()
                all_results.append(result_data)
            except Exception as exc:
                roll = future_to_roll[future]
                print(f"[!] {roll} generated an exception: {exc}")

    summary_csv_path = os.path.join(OUTPUT_DIR, 'batch_summary.csv')
    with open(summary_csv_path, 'w', newline='', encoding='utf-8') as csvfile:
        csv_writer = csv.writer(csvfile)
        csv_writer.writerow(["RollNumber", "Name", "SGPA", "CGPA", "Status"])
        all_results.sort(key=lambda x: x['roll_number'])
        for result in all_results:
            details = result['details']
            csv_writer.writerow([
                result['roll_number'], details['name'], details['sgpa'],
                details['cgpa'], result['status'].title()
            ])

    print(f"\n--- Batch processing complete! ---")
    print(f"Check the '{OUTPUT_DIR}' folder for results.")
    print(f"A summary has been saved to '{summary_csv_path}'.")

