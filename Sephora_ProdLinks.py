import pandas

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
import time

from selenium.webdriver.common.proxy import Proxy, ProxyType

from webdrivermanager import ChromeDriverManager


from bs4 import BeautifulSoup
from openpyxl import load_workbook

driver = webdriver.Chrome()
driver.get("https://www.sephora.com/shop/facial-treatments")
driver.refresh()

number_of_products = driver.find_element(By.CSS_SELECTOR, "p[data-at=number_of_products]")
number_of_products = number_of_products.get_attribute("innerHTML")
number_of_products = str(number_of_products).split(" ")[0]
number_of_products = int(number_of_products)

print(number_of_products)
    
product_elements = driver.find_elements(By.CSS_SELECTOR,"a[class=css-klx76]")
scroll_height = 400
#while len(product_elements) < number_of_products-30:
while len(product_elements) < 570:
    time.sleep(1)
   
    if driver.find_elements(By.XPATH, "//button[@class='css-1p9axos eanm77i0']"):
        driver.find_element(By.XPATH, "//button[@class='css-1p9axos eanm77i0']").click()
        
    driver.execute_script("window.scrollTo(0, "+str(scroll_height)+")")
    append_list = driver.find_elements(By.CSS_SELECTOR,"a[class=css-klx76]")
    for i in append_list:
        if i not in product_elements:
            product_elements.append(i)
    scroll_height += 400
    
    if driver.find_elements(By.XPATH, "//button[@class='css-1kna575']"):
        driver.find_element(By.XPATH, "//button[@class='css-1kna575']").click()
        
    
    print("Number of Links: " + str(len(product_elements)))
    
print("loop complete")
print("Number of Links: " + str(len(product_elements)))

product_links = []
for product in product_elements:
    product_links.append(product.get_attribute("href"))
    
print(len(product_links))

pandas.DataFrame(product_links).to_excel('sephoralinks.xlsx', header = True, index = False)

print("complete")
