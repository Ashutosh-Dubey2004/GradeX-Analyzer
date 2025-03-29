import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, WebDriverException, StaleElementReferenceException

from processCaptcha import process_captcha,load_easyocr

def init_easyocr():
    load_easyocr()

abort = False

def reset(driver):
    driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_btnReset").click()

def checkInvalidRollNumber(driver):
    try:
        WebDriverWait(driver, 5).until(EC.alert_is_present())
        alert = driver.switch_to.alert
        alert_text = alert.text.lower()  # Convert text to lowercase for case-insensitive matching
        print(f"Alert Detected: {alert.text}")

        if "you have entered a wrong text" in alert_text:
            alert.accept()  # Accept the alert
            time.sleep(1)
            return "retry"

        elif "result for this enrollment no. not found" in alert_text:
            alert.accept()
            return "invalid_roll"

        alert.accept()
        return None  # Some other unknown alert

    except TimeoutException:
        return None  
    except Exception as e:
        print(f"Error checking invalid roll number: {e}")
        return None

def extractSubjects(driver,course,semester):
    try:
        print("Extracting subjects...")
        if course == "DDMCA":
            base_xpath = "/html/body/form/div[3]/div/div[2]/table/tbody/tr[7]/td[1]/div/table/tbody/tr[3]/td[1]/table[{}]/tbody/tr[1]/td[1]"
        else:
            base_xpath = "/html/body/form/div[3]/div/div[2]/table/tbody/tr[9]/td[1]/div/table/tbody/tr[3]/td[1]/table[{}]/tbody/tr[1]/td[1]"

        subjects = []
        if course == 'MCA' and semester == 4:
            totalSubject = range(2, 7)
        else:
            totalSubject = range(2, 9)

        for n in totalSubject:
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

def extractResult(driver,course, semester):
    try:
        print("Extracting Result...")
        # For Grades 7 -DDMCA , 9-MCA
        if course == "DDMCA":
            base_xpath = "/html/body/form/div[3]/div/div[2]/table/tbody/tr[7]/td[1]/div/table/tbody/tr[3]/td[1]/table[{}]/tbody/tr[1]/td[4]"
        else:
            base_xpath = "/html/body/form/div[3]/div/div[2]/table/tbody/tr[9]/td[1]/div/table/tbody/tr[3]/td[1]/table[{}]/tbody/tr[1]/td[4]"

        grades = []
        if course == 'MCA' and semester == 4:
            totalSubject = range(2, 7)
        else:
            totalSubject = range(2, 9)

        for n in totalSubject:   
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
        Ids = {
            "name": "ctl00_ContentPlaceHolder1_lblNameGrading",
            "rollNo": "ctl00_ContentPlaceHolder1_lblRollNoGrading",
            "course": "ctl00_ContentPlaceHolder1_lblProgramGrading",
            "branch": "ctl00_ContentPlaceHolder1_lblBranchGrading",
            "semester": "ctl00_ContentPlaceHolder1_lblSemesterGrading",
            "status": "ctl00_ContentPlaceHolder1_lblStatusGrading"
        }

        studentInfo = []
        for key, id in Ids.items():
            try:
                # element = driver.find_element(By.ID, id)
                # studentInfo.append(element.text.strip())
                studentInfo.append(driver.find_element(By.ID, id).text.strip())
            except NoSuchElementException:
                studentInfo.append("N/A")

        return studentInfo
    except Exception as e:
        print(f"Error extracting student info: {e}")
        return []

def extractCompleteStudentInfo(driver,course,semester):
    try:
        # studentInfo = extractStudentInfo(driver)
        # grades = extractResult(driver)
        # return studentInfo + grades
        return extractStudentInfo(driver) + extractResult(driver,course,semester)
    except Exception as e:
        print(f"Error extracting complete student info: {e}")
        return []

def collectResultData(driver, course, semester, isFirstStudent):
    try:
        print(f"Collecting result data... {isFirstStudent}")
        data = []
        if isFirstStudent:
            headerRow = ['Name', 'Roll No.', 'Course', 'Branch', 'Semester', 'Status'] + extractSubjects(driver,course,semester) + ['SGPA', 'CGPA', 'Result']
            data.append(headerRow)

        # temp = extractCompleteStudentInfo(driver)
        # data.append(temp)
        # data.append(extractCompleteStudentInfo(driver,course))

        studentData = extractCompleteStudentInfo(driver, course, semester)
        if studentData:
            data.append(studentData)
        else:
            print("Warning: No student data extracted.")

        reset(driver)
        return data
    
    except Exception as e:
        print(f"Error collecting result data: {e}")
        return []

def retrieveStudentResult(driver, rollNo, course, semester, isFirstSuccess):
    currentURL = driver.current_url
    try:
        # rollNoInput = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_txtrollno")
        # rollNoInput.send_keys(rollNo)
        driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_txtrollno").send_keys(rollNo)

        # sem = Select(driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_drpSemester"))
        # sem.select_by_value(semester)
        Select(driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_drpSemester")).select_by_value(str(semester))

        # Retry Logic for Captcha
        attempt = 1  # Initialize attempt counter
        while True:  # Infinite loop until correct captcha
            print(f"Attempt {attempt} for {rollNo}")

            time.sleep(2)  # Allow time for captcha to load
            try:
                imageElement = driver.find_element(By.XPATH, "//img[@alt='Captcha']")
                extractedText = process_captcha(imageElement)
                captchaInput = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_TextBox1")
                captchaInput.clear()
                captchaInput.send_keys(extractedText)

                viewResult = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_btnviewresult")

                time.sleep(2)
                viewResult.click()

                # ALERT CHECK
                alert_status = checkInvalidRollNumber(driver)

                if alert_status == "retry":
                    print("Retrying due to incorrect captcha...")
                    attempt += 1
                    continue  # Retry with new captcha

                elif alert_status == "invalid_roll":
                    print(f"Skipping Roll Number {rollNo} (Invalid).")
                    driver.get(currentURL)
                    time.sleep(2)
                    return None, False  # Move to next roll number

                # time.sleep(2)
                # Fetch and return result data
                resultData = collectResultData(driver, course, semester, isFirstSuccess)

                print(f"First successful result fetched for {rollNo}" if isFirstSuccess else f"Result fetched for {rollNo}")

                return resultData, False  # Set isFirstSuccess to False after first success

            except StaleElementReferenceException:
                print("Element became stale, retrying...")
                time.sleep(2)  
                continue  
            
            except NoSuchElementException as e:
                print(f"Error during attempt {attempt} for {rollNo}: {e}")

            attempt+=1

    except (NoSuchElementException, TimeoutException) as e:
        print(f"Error occurred for {rollNo}: {e}")
    except Exception as e:
        print(f"Unexpected error for {rollNo}: {e}")

    driver.get(currentURL)  # Reset the page
    time.sleep(2)
    return None, isFirstSuccess  

def retrieveMultipleResults(course, semester, prefixRollNo, rollStart, rollEnd):
    print("Retrieving multiple results...")

    if course not in ["DDMCA", "MCA"]:
        print("Error: Invalid Course")
        return []

    URL = "https://result.rgpv.ac.in/Result/"
    # if course == "DDMCA":
    #     course_id = "radlstProgram_12"
    # elif course == "MCA":
    #     course_id = "radlstProgram_17"
    # else:
    #     print("Invalid Course")
    course_id = {"DDMCA": "radlstProgram_12", "MCA": "radlstProgram_17"}[course]

    options = Options()
    # options.add_argument("--headless=new")  #Chrome Window Not Visible
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument("--log-level=3")  # Suppress warnings

    try:
        driver = webdriver.Chrome(options=options)
        driver.get(URL)

        # Select course
        try:
            driver.find_element(By.ID, course_id).click()
        except NoSuchElementException:
            print("Error: Course selection failed.")
            driver.quit()
            return []

        data = []
        isFirstSuccess = True

        for rollNo in range(rollStart, rollEnd + 1):
            if abort:
                print("Aborted")
                break  

            fullRollNo = f"{prefixRollNo}{str(rollNo).zfill(2)}"
            print(f"Fetching result for: {fullRollNo}")

            try:
                studentData, isFirstSuccess = retrieveStudentResult(driver, fullRollNo, course, semester, isFirstSuccess)
                if studentData:
                    data.extend(studentData)
                else:
                    print(f"Skipped Roll No: {fullRollNo}")

            except (TimeoutException, WebDriverException) as e:
                print(f"Error fetching {fullRollNo}: {e}")
                driver.quit()
                return []

        return data

    except WebDriverException as e:
        print(f"WebDriver Error: {e}")
        return []

    finally:
        driver.quit()  

if __name__ == '__main__':
    # data = retrieveMultipleResults('DDMCA', '7', '0827CA21DD', 5, 8)
    data = retrieveMultipleResults('MCA', '3', '0827CA2310', 5, 8)
    from pprint import pprint
    pprint(data)
