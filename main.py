import pandas as pd
from selenium import webdriver
from selenium.common import TimeoutException
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

google_chorme_address = r'C:\Program Files\Google\Chrome\ChromeDriver\chromedriver.exe'
feishu_url = 'https://nankai.feishu.cn/next/messenger/'
original_file_path = r"./old.xlsx"
new_file_path = r"./new.xlsx"

# 设置Chrome驱动
service = Service(google_chorme_address)
driver = webdriver.Chrome(service=service)


def visit_feishu(url):
    driver.get(url)
    print("Page opened", url)


def scan_to_login():
    print("Please scan with your phone...")
    try:
        wait = WebDriverWait(driver, 60)
        wait.until(EC.presence_of_element_located((By.ID, 'a11y-open-btn')))
        print("Logged in successfully!")
    except TimeoutException:
        print("Login timeout or QR code scan failed.")


def get_search_box():
    wait = WebDriverWait(driver, 10)

    print("wait to located...")
    body = wait.until(EC.presence_of_element_located((By.XPATH, '/html/body')))
    print("body window located!")
    body.send_keys(Keys.CONTROL + 'k')

    print("pressed ctrl+k")

    print("wait to located...")
    search_box = wait.until(EC.presence_of_element_located(
        (By.XPATH, '//*[contains(@class, "zone-container editor-kit-container notranslate chrome window chrome88")]')))
    print("Search box located!")
    return search_box


def search_info(box, name_list, record_list, delay_sec=3):
    for student in name_list:
        if pd.isna(student):
            record_list.append(('', '', ''))
            continue

        box.click()
        box.send_keys(Keys.CONTROL + "a")
        box.send_keys(Keys.BACKSPACE)
        box.send_keys(student)
        print(f"Entered: {student}")

        time.sleep(delay_sec)
        try:
            stu_id = driver.find_element(By.XPATH,
                                         '//*[contains(@class, "common-items-wrapper smart-search-tab")]/div[1]/div[2]/div[2]/div/div[1]/span')
            stu_college = driver.find_element(By.XPATH,
                                              '//*[contains(@class, "common-items-wrapper smart-search-tab")]/div[1]/div[2]/div[2]/div/div[2]/div/span')
            record_tuple = (student, stu_id.text, stu_college.text)
            record_list.append(record_tuple)
            print(record_tuple, "done")
        except Exception as e:
            print(f"Error occurred for student {student}: {e}")
            record_list.append((student, '', ''))


def auto_complete_stu_info():
    data_raw = pd.read_excel(original_file_path)
    name_list = data_raw['姓名'].tolist()
    columns = ['姓名', '学号', '学院']
    result_list = []

    visit_feishu(feishu_url)
    scan_to_login()
    search_box = get_search_box()
    search_info(search_box, name_list, result_list)

    result = pd.DataFrame(result_list, columns=columns)
    result.to_excel(new_file_path, index=False)
    print("Total done!!!")

    driver.quit()


try:
    auto_complete_stu_info()
except Exception as e:
    print(f"Error occurred: {e}")
finally:
    driver.quit()
