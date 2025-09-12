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
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
import concurrent.futures # The library for multithreading

# --- Configuration (Unchanged) ---
BASE_URL = "https://jcboseustymca.co.in/Forms/Student/ResultStudents.aspx"
RESULT_URL = "https://jcboseustymca.co.in/Forms/Student/PrintReportCardNew.aspx"
CAPTCHA_URL = "https://jcboseustymca.co.in/Handler/GenerateCaptchaImage.ashx"
OUTPUT_DIR = "results"

# --- All helper functions (solve_captcha, save_html_as_image, etc.) remain exactly the same ---
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
        print(f"  [!] OCR Error for thread {os.getpid()}: {e}")
        return ""

def save_html_as_image(html_content, output_path):
    print(f"  📸 Capturing screenshot to {output_path}...")
    try:
        options = webdriver.ChromeOptions()
        options.add_argument("--headless")
        options.add_argument("--window-size=1200,1000")
        options.add_argument("--hide-scrollbars")
        # Suppress console logging from Chrome
        options.add_experimental_option('excludeSwitches', ['enable-logging'])
        
        with tempfile.NamedTemporaryFile(delete=False, suffix=".html", mode='w', encoding='utf-8') as fp:
            fp.write(html_content)
            temp_path = "file://" + fp.name
        driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=options)
        driver.get(temp_path)
        time.sleep(2)
        total_height = driver.execute_script("return document.body.parentNode.scrollHeight")
        driver.set_window_size(1200, total_height)
        time.sleep(1)
        driver.save_screenshot(output_path)
        driver.quit()
        os.unlink(fp.name)
        print(f"  ✅ Screenshot saved for {os.path.basename(os.path.dirname(output_path))}.")
        return True
    except Exception as e:
        print(f"  [!] Screenshot Error for {os.path.basename(os.path.dirname(output_path))}: {e}")
        return False

def parse_result_details(soup):
    details = {}
    def get_text_by_id(element_id):
        try:
            return soup.find('span', {'id': element_id}).text.strip()
        except AttributeError: return "N/A"
    details['student_info'] = {'name': get_text_by_id('lblname')}
    details['result_summary'] = {'sgpa': get_text_by_id('lblResult'), 'cgpa': get_text_by_id('lblCgpaResult')}
    return details

def fetch_result(session, roll_number, semester):
    max_attempts = 15
    for attempt in range(max_attempts):
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

            if soup_result.find("span", {"id": "lblResult"}) and soup_result.find("span", {"id": "lblResult"}).text.strip():
                print(f"  🎉 Success for {roll_number} on attempt {attempt + 1}")
                parsed_data = parse_result_details(soup_result)
                return {"status": "success", "html": report_page.text, "details": parsed_data}
            
            # If login fails, a small pause before retrying
            time.sleep(1)

        except Exception:
             # Broad exception to catch any request/parsing errors and just retry
            continue
            
    print(f"  ❌ Failed to fetch result for {roll_number} after {max_attempts} attempts.")
    return {"status": "failed"}

# --- NEW: Worker Function for Multithreading ---
def process_roll_number(roll_number, semester):
    """
    This function handles the entire process for a single roll number.
    It's what each thread will run.
    """
    print(f"  Starting thread for {roll_number}...")
    with requests.Session() as session:
        result_data = fetch_result(session, roll_number, semester)
    
    # Prepare the data to be returned for the CSV file
    if result_data and result_data["status"] == "success":
        details = result_data['details']
        roll_dir = os.path.join(OUTPUT_DIR, roll_number)
        os.makedirs(roll_dir, exist_ok=True)
        
        image_path_abs = os.path.join(roll_dir, f"result_sem_{semester}.png")
        save_html_as_image(result_data["html"], image_path_abs)
        
        return {
            "RollNumber": roll_number,
            "Name": details['student_info']['name'],
            "SGPA": details['result_summary']['sgpa'],
            "CGPA": details['result_summary']['cgpa'],
            "Status": "Success"
        }
    else:
        return {
            "RollNumber": roll_number, "Name": "N/A", "SGPA": "N/A", "CGPA": "N/A", "Status": "Failed"
        }

# --- Main Execution Block (MODIFIED for Multithreading) ---
if __name__ == "__main__":
    roll_no_file = input("Enter the path to the .txt file containing roll numbers: ")
    semester = input("Enter the semester for all roll numbers (e.g., 1, 2, 03...): ")

    if not os.path.exists(roll_no_file):
        print(f"Error: The file '{roll_no_file}' was not found.")
        exit()

    with open(roll_no_file, 'r') as f:
        roll_numbers = [line.strip() for line in f if line.strip()]

    print(f"\nFound {len(roll_numbers)} roll numbers to process for semester {semester}.")
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    
    # Set the maximum number of concurrent threads.
    # This prevents sending too many requests to the server at once.
    MAX_WORKERS = min(10, len(roll_numbers))
    
    all_results = []
    
    # Using ThreadPoolExecutor to process roll numbers in parallel
    with concurrent.futures.ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        # Create a "future" for each roll number. This submits the jobs to the thread pool.
        future_to_roll = {executor.submit(process_roll_number, roll, semester): roll for roll in roll_numbers}
        
        for future in concurrent.futures.as_completed(future_to_roll):
            try:
                # Get the result from the completed thread
                result = future.result()
                all_results.append(result)
            except Exception as exc:
                roll = future_to_roll[future]
                print(f"  [!] An error occurred while processing {roll}: {exc}")
                all_results.append({"RollNumber": roll, "Name": "N/A", "SGPA": "N/A", "CGPA": "N/A", "Status": "Error"})

    # --- Write results to CSV after all threads are finished ---
    summary_csv_path = os.path.join(OUTPUT_DIR, 'batch_summary.csv')
    # Sort results by Roll Number to maintain order
    all_results.sort(key=lambda x: x['RollNumber'])
    
    with open(summary_csv_path, 'w', newline='', encoding='utf-8') as csvfile:
        fieldnames = ["RollNumber", "Name", "SGPA", "CGPA", "Status"]
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(all_results)

    print(f"\n\n--- Batch processing complete! ---\nCheck the '{OUTPUT_DIR}' folder for results.")
    print(f"A summary has been saved to '{summary_csv_path}'.")
