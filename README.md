# ReSCo: Result Scraper and Compiler

**ReSCo** is a powerful Python automation script designed to batch-download student results from the J.C. Bose University of Science and Technology, YMCA, Faridabad portal. It handles CAPTCHA solving, processes multiple roll numbers from a text file, and saves the results in an organised manner.

---

## Features

- **Batch Processing:** Fetches results for multiple roll numbers listed in a `.txt` file.
- **Automated CAPTCHA Solving:** Uses Pytesseract OCR to read and solve the image CAPTCHA automatically.
- **Multithreaded for Speed:** Processes up to 10 roll numbers concurrently for a significant speed boost over sequential scraping.
- **Organised Output:**
  - Creates a master results folder.
  - Generates a separate folder for each roll number.
  - Saves a full-page screenshot (`.png`) of the result card in the respective folder.
  - Compiles all data into a master `batch_summary.csv` for easy viewing.
- **Robust and Resilient:** Includes retry logic for failed CAPTCHA attempts and handles common web scraping errors gracefully.

---

## Tech Stack

- **Python:** Core programming language.
- **Requests:** For handling HTTP requests and sessions.
- **BeautifulSoup4:** For parsing HTML and extracting data.
- **Selenium:** To render the final result page and capture a full-page screenshot.
- **Pillow (PIL):** For image processing before OCR.
- **Pytesseract:** An OCR tool to read the CAPTCHA images.
- **webdriver-manager:** To automatically manage the browser drivers for Selenium.

---

## Setup and Usage

### Prerequisites

- Python 3.x  
- Google Tesseract OCR installed and accessible in your system's PATH.
- https://github.com/Harshitmishra09/resco/blob/main/tesseract-ocr-w64-setup-5.5.0.20241111.exe
- https://github.com/UB-Mannheim/tesseract/wiki

### Installation

Clone the repository and set up environment:

```bash
git clone https://github.com/Harshitmishra09/resco.git
cd resco

python -m venv venv
# On Windows
venv\Scripts\activate
# On macOS/Linux
source venv/bin/activate
