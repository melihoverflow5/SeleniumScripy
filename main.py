from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import pandas as pd

# Replace 'your_excel_file.xlsx' with the path to your actual Excel file
file_path = 'data.xlsx'

# Read the Excel file
data_frame = pd.read_excel(file_path)

data = data_frame.to_numpy().tolist()


row = 0

# Setup WebDriver (make sure to specify the correct path if not on PATH)
driver = webdriver.Chrome()

# URL of the Google Form
form_url = 'https://docs.google.com/forms/d/e/1FAIpQLSd1zZYIKP1UxASmXodmz4Gd5SAvKMLtgn3Ifm6fgjs4ok6LIg/viewform'
driver.get(form_url)

try:
    # Wait for the form to load
    wait = WebDriverWait(driver, 10)
    
    next_button = driver.find_element(By.XPATH, "//span[text()='Sonraki']")
    if next_button:
        next_button.click()
        
    time.sleep(1)
    birth_year_input = wait.until(EC.element_to_be_clickable((By.CLASS_NAME, "whsOnd")))
    birth_year_input.clear()  # Clear any pre-filled value
    birth_year_input.send_keys(data[row][1])
    
    if data[row][2] == "Kadın":
        female_option = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[@aria-label='Kadın']")))
        female_option.click()
    else:
        male_option = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[@aria-label='Erkek']")))
        male_option.click()
    
    if data[row][3] == "Evli":
        evli_option = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[@data-value='Evli']")))
        evli_option.click()
    else:
        bekar_option = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[@data-value='Evli değil']")))
        bekar_option.click()
    
    if data[row][4] == "Çekirdek ailemle":
        cekirdek_option = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[@data-value='Kendi çekirdek ailemle']")))
        cekirdek_option.click()
    elif data[row][4] == "Kendi başıma":
        kendi_basima_option = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[@data-value='Tek başıma']")))
        kendi_basima_option.click()
    elif data[row][4] == "Aile büyükleriyle":
        aile_buyukleriyle_option = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[@data-value='Aile büyükleri ile']")))
        aile_buyukleriyle_option.click()
    else:
        diger_option = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[@data-value='Diğer']")))
        diger_option.click()
    
    cocuk_option = wait.until(EC.element_to_be_clickable((By.XPATH, f"//div[@data-value='{data[row][5]}']")))
    cocuk_option.click()
    
    
    all_input_elements = driver.find_elements(By.CLASS_NAME, "whsOnd")
    kilo_input = all_input_elements[1]
    kilo_input.clear()  # Clear any pre-filled value
    kilo_input.send_keys(data[row][6])
    boy_input = all_input_elements[2]
    boy_input.clear()  # Clear any pre-filled value
    boy_input.send_keys(data[row][7])
    
    questions = driver.find_elements(By.CLASS_NAME, "Qr7Oae")
    aktivite_option = questions[8].find_element(By.XPATH, f".//div[@data-value='{data[row][8]}']")
    aktivite_option.click()
    
    ogrenci_option = questions[9].find_element(By.XPATH, f".//div[@data-value='{data[row][9]}']")
    ogrenci_option.click()
    
    
    next_button = driver.find_element(By.XPATH, "//span[text()='Sonraki']")
    
    next_button.click()
    time.sleep(2)


    if data[row][9] == 'Hayır':
        time.sleep(60)
    
    questions = driver.find_elements(By.CLASS_NAME, "lLfZXe")
    for i in range(52):
        if data[row][22+i] == "Düzenli Olarak":
            q = questions[i].find_element(By.XPATH, f".//div[@data-value='Düzenli Olarak']")
        else:
            q = questions[i].find_element(By.XPATH, f".//div[@data-value='{data[row][22+i].capitalize()}']")
        q.click()

    next_button = driver.find_element(By.XPATH, "//span[text()='Sonraki']")
    
    next_button.click()
    time.sleep(2)
    
    questions = driver.find_elements(By.CLASS_NAME, "Qr7Oae")
    if data[row][74] == "3+ saat":
        questions[1].find_element(By.XPATH, f".//div[@data-value='3 saatten fazla']").click()
    elif data[row][74] == "2-3 saat":
        questions[1].find_element(By.XPATH, f".//div[@data-value='2-3 saat arası']").click()
    elif data[row][74] == "1-2 saat":
        questions[1].find_element(By.XPATH, f".//div[@data-value='1-2 saat arası']").click()
    else:
        questions[1].find_element(By.XPATH, f".//div[@data-value='1 saatten az']").click()

    if data[row][75] == "3+ saat":
        questions[2].find_element(By.XPATH, f".//div[@data-value='3 saatten fazla']").click()
    elif data[row][75] == "2-3 saat":
        questions[2].find_element(By.XPATH, f".//div[@data-value='2-3 saat arası']").click()
    elif data[row][75] == "1-2 saat":
        questions[2].find_element(By.XPATH, f".//div[@data-value='1-2 saat arası']").click()
    else:
        questions[2].find_element(By.XPATH, f".//div[@data-value='1 saatten az']").click()
    
    questions[3].find_element(By.XPATH, f".//div[@data-value='{data[row][76]}']").click()
    
    
    next_button = driver.find_element(By.XPATH, "//span[text()='Sonraki']")
    
    next_button.click()
    time.sleep(2)
    
    questions = driver.find_elements(By.CLASS_NAME, "Qr7Oae")
    for i in range(9):
        questions[1+i].find_element(By.XPATH, f".//div[@data-value='{data[row][77+i]}']").click()
    
    time.sleep(20)
    
except Exception as e:
    print(f"An error occurred: {e}")

finally:
    # Optionally close the driver or do other cleanup here
    driver.quit()