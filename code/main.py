import os
from selenium import webdriver
from utils import download

# save directory
data_dir = os.path.abspath(os.path.join(os.getcwd(), '../dataset'))

# driver setting
options = webdriver.ChromeOptions()
options.add_argument('--headless') # headless를 하는 경우 chrome 브라우저를 표시하지 않기 때문에 보다 빠른 크롤링이 가능합니다.
driver = webdriver.Chrome(options=options)

# crawling
# driver = download.kbo_crowd(driver, data_dir)
driver = download.statiz_record(driver, data_dir)
driver.close()


