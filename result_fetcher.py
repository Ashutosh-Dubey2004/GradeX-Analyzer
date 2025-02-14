import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import NoAlertPresentException
import requests 

import cv2
import numpy as np
from PIL import Image
from io import BytesIO
import pytesseract
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"


def process_captcha(image_element):
    image = Image.open(BytesIO(image_element.screenshot_as_png))
    gray = cv2.cvtColor(np.array(image), cv2.COLOR_RGB2GRAY)
    processed_img = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 31, 2)
    extracted_text = pytesseract.image_to_string(processed_img, config="--psm 6").strip().replace(" ", "")

    return extracted_text


def fetch_result_data(URL, roll_no, semester):
    options = Options()
    options.add_experimental_option("excludeSwitches",["enable-automation"])
    driver = webdriver.Chrome(options=options)
    driver.get(URL)

    rollNo_input = driver.find_element(By.XPATH, "//input[@id='ctl00_ContentPlaceHolder1_txtrollno']")
    rollNo_input.send_keys(roll_no)

    sem = Select(driver.find_element(By.XPATH,"//select[@id='ctl00_ContentPlaceHolder1_drpSemester']"))
    sem.select_by_value(semester)

    image_element = driver.find_element(By.XPATH, "//img[@alt='Captcha']")
    extracted_text = process_captcha(image_element)

    capchaInput = driver.find_element(By.XPATH, "(//input[@id='ctl00_ContentPlaceHolder1_TextBox1'])[1]")
    capchaInput.send_keys(extracted_text)

    viewResult = driver.find_element(By.XPATH, "(//input[@id='ctl00_ContentPlaceHolder1_btnviewresult'])[1]")

    time.sleep(5)    
    viewResult.click()

    time.sleep(1)
    driver.quit()



def fetch_results(course, semester,prefixRollNO, roll_start, roll_end):
    results = []
    if course == 'DDMCA':
        URL = "https://result.rgpv.ac.in/Result/McaDDrslt.aspx"
    elif course == 'MCA':
        URL = "https://result.rgpv.ac.in/Result/MCArslt.aspx"
    else:
        print("Error")
    
    for roll_no in range(roll_start, roll_end + 1):
        full_roll_no = f"{prefixRollNO}{str(roll_no).zfill(2)}"
        print(f"Fetching result for: {full_roll_no}")
        
        studentData = fetch_result_data(URL, full_roll_no, semester)
        results.append(studentData)

    return results

if __name__ == '__main__':
    fetch_results('DDMCA','7','0827CA21DD',5,15)