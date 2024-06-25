from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import openpyxl
import time
service = Service(executable_path="chromedriver.exe")
driver = webdriver.Chrome(service=service)
driver.get("https://meroshare.cdsc.com.np/#/login")

branchcode = ''
userid = ''
password = ''

wait = WebDriverWait(driver, 20)
wait.until(EC.presence_of_element_located((By.TAG_NAME, "app-login")))

user_id  = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="username"]')))
pw_id = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="password"]')))
branch = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="selectBranch"]')))

branch.click()
branch_id = driver.find_element(By.CLASS_NAME, 'select2-search__field')
branch_id.click()
branch_id.send_keys(branchcode)
branch_id.send_keys(Keys.ENTER)
user_id.send_keys(userid)
pw_id.send_keys(password)
time.sleep(2)

login_btn = wait.until(EC.presence_of_element_located((By.CLASS_NAME, 'sign-in')))
login_btn.click()
time.sleep(1)

btn_port = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="sideBar"]/nav/ul/li[5]/a')))
btn_port.click()
time.sleep(5)

table = driver.find_element(By.XPATH, '//*[@id="main"]/div/app-my-portfolio/div/div[2]/div/div/table')
rows = table.find_elements(By.TAG_NAME, "tr")

all_data = []
for row in rows[1:-1]:
    cells = row.find_elements(By.TAG_NAME, "td")
    data = {'Name':cells[1].text, 'Price':cells[5].text}
    print(cells[1].text, cells[5].text)
    all_data.append(data)
print(all_data)

df = pd.DataFrame(all_data)
df.to_excel("portfolio.xlsx")

driver.quit()