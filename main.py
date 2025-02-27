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
directory_path=r"./"
# directory_path=r"D:/Desktop/"
original_file_path = directory_path + r"old.xlsx"
new_file_path = directory_path + r"new.xlsx"

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
    li_contactor = wait.until(EC.presence_of_element_located((By.XPATH, '//div[text()="联系人"]')))
    print("li_contactor located!")
    li_contactor.click()

    print("wait to located...")
    search_box = wait.until(EC.presence_of_element_located(
        (By.XPATH, '//*[contains(@class, "zone-container editor-kit-container notranslate chrome window chrome88")]')))
    print("Search box located!")
    time.sleep(1)
    return search_box


def search_info(box, name_list, record_list, get_first, basic_filter, delay_sec=1):
    # get_first: TRUE-只记录首个结果 / FALSE-记录查找到的所有结果
    # basic_filter: (get_first=FALSE 时有效) TRUE-保证名字完全匹配，且排除“未激活”或“暂停使用”的用户 / FALSE-不筛选
    # delay_sec: 将搜索框中的内容删除后，查找列表消失所用的最大时间（只需保证时间段结束后不会把上一个查找结果录入，无需保证在这个时间段内就完成下一个查找）

    wait = WebDriverWait(driver, 10)

    for student in name_list:
        if pd.isna(student):
            record_list.append(('', '', ''))
            continue

        box.send_keys(student)
        print(f"Entered: {student}")

        time.sleep(delay_sec)

        if get_first:
            index = 1
            info_wrapper = f'//*[contains(@class, "common-items-wrapper other-search-tab")]/div[{index}]/div[2]/div[2]/div'
            try:
                stu_id = wait.until(EC.presence_of_element_located((By.XPATH,
                                             info_wrapper + '/div[1]/span'))).text
                # 综合查询：id //*[contains(@class, "common-items-wrapper smart-search-tab")]/div[1]/div[2]/div[2]/div/div[1]/span
                stu_college = wait.until(EC.presence_of_element_located((By.XPATH,
                                                  info_wrapper + '/div[2]/div/span'))).text
                # 综合查询：college //*[contains(@class, "common-items-wrapper smart-search-tab")]/div[1]/div[2]/div[2]/div/div[2]/div/span
                record_tuple = (student, stu_id, stu_college)
                record_list.append(record_tuple)
                print(record_tuple, "done")
            except Exception as e:
                print(f"Error occurred for student {student}: {e}")
                record_list.append((student, 'ERROR', 'ERROR'))

        else:
            div_wrapper = wait.until(EC.presence_of_element_located(
                (By.XPATH, '//*[contains(@class, "common-items-wrapper other-search-tab")]')))
            div_count = len(div_wrapper.find_elements(By.XPATH, './div'))
            print("div_count =",div_count)

            for index in range(1, div_count + 1):
                info_wrapper = f'//*[contains(@class, "common-items-wrapper other-search-tab")]/div[{index}]/div[2]'
                try:

                    # [class]user-chatter-info-name
                    stu_name = wait.until(EC.presence_of_element_located((By.XPATH,
                                                     info_wrapper + '/div[1]/div[1]/div[1]'))).text.strip()
                    stu_note = wait.until(EC.presence_of_element_located((By.XPATH,
                                                     info_wrapper + '/div[1]/div[1]/div[2]'))).text
                    if basic_filter:
                        if index > 1 and not stu_name == student or stu_note == "未激活" or stu_note == "暂停使用":
                            continue

                    try:
                        stu_id = wait.until(EC.presence_of_element_located((By.XPATH,
                                                     info_wrapper + '/div[2]/div/div[1]/span'))).text
                        stu_college = wait.until(EC.presence_of_element_located((By.XPATH,
                                                          info_wrapper + '/div[2]/div/div[2]/div/span'))).text
                        if index==1:
                            record_tuple = (student, stu_name, stu_id, stu_college, stu_note)
                        else:
                            record_tuple = ('', stu_name, stu_id, stu_college, stu_note)
                        record_list.append(record_tuple)
                        print(record_tuple, "done")
                    except Exception as e:
                        print(f"Error occurred for student {student}: {e}")
                        record_list.append((student, 'ERROR', 'ERROR', 'ERROR', 'ERROR'))
                    # time.sleep(0.5)

                except Exception:
                    # 如果没有找到相关元素，就跳过当前用户卡片
                    print("Skipping user card.")
                    continue

        box.send_keys(Keys.CONTROL + "a")
        box.send_keys(Keys.BACKSPACE)


def auto_complete_stu_info(get_first=False, basic_filter=True):
    data_raw = pd.read_excel(original_file_path)
    name_list = data_raw['姓名'].tolist()
    if not get_first:
        columns = ['查找', '姓名', '学号', '学院', '备注']
    else:
        columns = ['姓名', '学号', '学院']
    result_list = []

    visit_feishu(feishu_url)
    scan_to_login()
    search_box = get_search_box()
    search_info(search_box, name_list, result_list, get_first,basic_filter)

    result = pd.DataFrame(result_list, columns=columns)
    result.to_excel(new_file_path, index=False)


try:
    auto_complete_stu_info(get_first=False, basic_filter=False)
    print("Data entry has been completed!")
except Exception as e:
    print(f"Error occurred: {e}")
finally:
    driver.quit()
