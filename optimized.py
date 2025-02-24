import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, WebDriverException

from processCaptcha import process_captcha

abort = False

def extractSubjects(driver):
    try:
        print("Extracting subjects...")
        base_xpath = "/html/body/form/div[3]/div/div[2]/table/tbody/tr[7]/td[1]/div/table/tbody/tr[3]/td[1]/table[{}]/tbody/tr[1]/td[1]"

        subjects = []
        for n in range(2, 9):
            try:
                # xpath = base_xpath.format(n)
                # element = driver.find_element(By.XPATH, xpath)
                # subjects.append(element.text)
                subjects.append(driver.find_element(By.XPATH, base_xpath.format(n)).text)
            except NoSuchElementException:
                break

        return subjects
    except Exception as e:
        print(f"Error extracting subjects: {e}")
        return []

def extractResult(driver):
    try:
        print("Extracting Result...")
        # For Grades
        base_xpath = "/html/body/form/div[3]/div/div[2]/table/tbody/tr[7]/td[1]/div/table/tbody/tr[3]/td[1]/table[{}]/tbody/tr[1]/td[4]"

        grades = []
        for n in range(2, 9):   
            try:
                # xpath = base_xpath.format(n)
                # element = driver.find_element(By.XPATH, xpath)
                # grades.append(element.text)
                grades.append(driver.find_element(By.XPATH, base_xpath.format(n)).text)
            except NoSuchElementException:
                break
        # For Result, SGPA, CGPA
        try : 
            result = driver.find_element(By.ID,"ctl00_ContentPlaceHolder1_lblResultNewGrading").text
            sgpa = driver.find_element(By.ID,"ctl00_ContentPlaceHolder1_lblSGPA").text
            cgpa = driver.find_element(By.ID,"ctl00_ContentPlaceHolder1_lblcgpa").text
        except NoSuchElementException:
            print("Exception in Result, SGPA, CGPA")
        
        return grades + [sgpa, cgpa, result]
    except Exception as e:
        print(f"Error extracting grades: {e}")
        return []

def extractStudentInfo(driver):
    try:
        print("Extracting student information...")
        xpaths = {
            "name": "/html/body/form/div[3]/div/div[2]/table/tbody/tr[7]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/span",
            "rollNo": "/html/body/form/div[3]/div/div[2]/table/tbody/tr[7]/td/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td[4]/span",
            "course": "/html/body/form/div[3]/div/div[2]/table/tbody/tr[7]/td/div/table/tbody/tr[2]/td/table/tbody/tr[3]/td[2]/span",
            "branch": "/html/body/form/div[3]/div/div[2]/table/tbody/tr[7]/td/div/table/tbody/tr[2]/td/table/tbody/tr[3]/td[4]",
            "semester": "/html/body/form/div[3]/div/div[2]/table/tbody/tr[7]/td/div/table/tbody/tr[2]/td/table/tbody/tr[4]/td[2]",
            "status": "/html/body/form/div[3]/div/div[2]/table/tbody/tr[7]/td/div/table/tbody/tr[2]/td/table/tbody/tr[4]/td[4]/span"
        }

        studentInfo = []
        for key, xpath in xpaths.items():
            try:
                # element = driver.find_element(By.XPATH, xpath)
                # studentInfo.append(element.text.strip())
                studentInfo.append(driver.find_element(By.XPATH, xpath).text.strip())
            except NoSuchElementException:
                studentInfo.append("N/A")

        return studentInfo
    except Exception as e:
        print(f"Error extracting student info: {e}")
        return []

def extractCompleteStudentInfo(driver):
    try:
        # studentInfo = extractStudentInfo(driver)
        # grades = extractResult(driver)
        # return studentInfo + grades
        return extractStudentInfo(driver) + extractResult(driver)
    except Exception as e:
        print(f"Error extracting complete student info: {e}")
        return []

def collectResultData(driver, isFirstStudent):
    try:
        print(f"Collecting result data... {isFirstStudent}")
        if isFirstStudent:
            headerRow = ['Name', 'Roll No.', 'Course', 'Branch', 'Semester', 'Status'] + extractSubjects(driver) + ['SGPA', 'CGPA', 'Result']
            data = [headerRow]
        else:
            data = []

        # temp = extractCompleteStudentInfo(driver)
        # data.append(temp)
        data.append(extractCompleteStudentInfo(driver))
        return data
    except Exception as e:
        print(f"Error collecting result data: {e}")
        return []

def checkInvalidRollNumber(driver):
    try:
        WebDriverWait(driver, 5).until(EC.alert_is_present())
        alert = driver.switch_to.alert
        print(f"Alert Detected: {alert.text}")
        alert.accept()
        time.sleep(1)
        return True
    except TimeoutException:
        return False
    except Exception as e:
        print(f"Error checking invalid roll number: {e}")
        return False

def retrieveStudentResult(driver, URL, rollNo, semester, isFirstSuccess):
    try:
        driver.get(URL)

        # rollNoInput = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_txtrollno")
        # rollNoInput.send_keys(rollNo)
        driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_txtrollno").send_keys(rollNo)
        
        # sem = Select(driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_drpSemester"))
        # sem.select_by_value(semester)

        Select(driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_drpSemester")).select_by_value(semester)

        attempt = 1
        while attempt <= 2:
            print(f"Attempt {attempt} for {rollNo}")
            time.sleep(2)

            imageElement = driver.find_element(By.XPATH, "//img[@alt='Captcha']")
            extractedText = process_captcha(imageElement)
            capchaInput = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_TextBox1")
            capchaInput.clear()
            capchaInput.send_keys(extractedText)
            
            viewResult = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_btnviewresult")

            time.sleep(3)
            viewResult.click()

            if checkInvalidRollNumber(driver):
                print(f"Alert detected on attempt {attempt} for {rollNo}. Retrying...")
                attempt += 1
                if attempt > 2:
                    print(f"Skipping {rollNo} after 2 failed attempts.")
                    return None, isFirstSuccess
                continue

            time.sleep(2)
            resultData = collectResultData(driver, isFirstSuccess)
            print(f"First successful result fetched for: {rollNo}" if isFirstSuccess else f"Result Fetched for {rollNo}")
            return resultData, False  # isFirstSuccess False after first success

    except Exception as e:
        print(f"Error occurred for {rollNo}: {e}")
    return None, isFirstSuccess

def retrieveMultipleResults(course, semester, prefixRollNo, rollStart, rollEnd):
    print("Retrieving multiple results...")

    isFirstSuccess = True
    data = []

    # if course == 'DDMCA':
    #     URL = "https://result.rgpv.ac.in/Result/McaDDrslt.aspx"
    # elif course == 'MCA':
    #     URL = "https://result.rgpv.ac.in/Result/MCArslt.aspx"
    # else:
    #     print("Error: Invalid Course")
    #     return []
    URL = "https://result.rgpv.ac.in/Result/McaDDrslt.aspx" if course == 'DDMCA' else "https://result.rgpv.ac.in/Result/MCArslt.aspx"
    
    options = Options()
    # options.add_argument("--headless=new")  #Chrome Window Not Visible
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument("--log-level=3")  # Suppresses warnings
    driver = webdriver.Chrome(options=options)

    try:
        for rollNo in range(rollStart, rollEnd + 1):
            if abort:
                print("Abort")
                return 
            
            fullRollNo = f"{prefixRollNo}{str(rollNo).zfill(2)}"
            print(f"Fetching result for: {fullRollNo}")

            studentData, isFirstSuccess = retrieveStudentResult(driver, URL, fullRollNo, semester, isFirstSuccess)

            if studentData is None:
                print(f"Skipped Roll No: {fullRollNo}")
                continue

            data.extend(studentData)

    except WebDriverException as e:
        print(f"WebDriver Error: {e}")
    finally:
        driver.quit() 
    return data

if __name__ == '__main__':
    data = retrieveMultipleResults('DDMCA', '7', '0827CA21DD', 5, 6)
    # data = retrieveMultipleResults('MCA', '3', '0827CA2310', 5, 6)
    from pprint import pprint
    pprint(data)
