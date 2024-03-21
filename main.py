from selenium import webdriver
from selenium.common import NoSuchElementException
from selenium.common import TimeoutException
from selenium.common.exceptions import ElementNotInteractableException
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import time
import pandas as pd
import os

class Scraper:
    def __init__(self):
        self.headless = False
        self.driver = None
        self.list_of_scrapped_data = []

    def config_driver(self):
        options = webdriver.ChromeOptions()
        if self.headless:
            options.add_argument("--headless")
        s = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=s, options=options)
        self.driver = driver
        self.driver.maximize_window()

    def scrap_data(self,given_name, investor_name, address):
        wait = WebDriverWait(self.driver, 20)
        self.list_of_scrapped_data = []


        sections_xpath = [
            "//span[contains(., 'Amount Pending With Company')]/following-sibling::table[1]",
            "//span[contains(., 'Amount Transfered to IEPF')]/following-sibling::table[1]",
            "//font[contains(., 'Amount Refunded from IEPF')]/following-sibling::table[1]",
            "//font[contains(., 'Shares Transferred to IEPF account')]/following-sibling::table[1]",
            "//font[contains(., 'Shares Refunded from IEPF account')]/following-sibling::table[1]"
        ]

        #1
        try:
            xpath = sections_xpath[0]
            try:
                table = WebDriverWait(self.driver, 2).until(EC.presence_of_element_located((By.XPATH, xpath)))
            except: 
                xpath = "//font[contains(., 'Amount Pending With Company')]/following-sibling::table[1]"
                table = WebDriverWait(self.driver, 0.1).until(EC.presence_of_element_located((By.XPATH, xpath)))

            rows = table.find_elements(By.XPATH, ".//tbody/tr[@class='RowDataAlt' or @class='RowData']")
            for row in rows:
                row_data = [td.text for td in row.find_elements(By.XPATH, ".//td")]
                item = {
                    "Given Name": given_name,
                    "Inversor Name" : investor_name,
                    "Address" : address,
                    "Type" : "Amount Pending With Company",
                    "Company Name": row_data[0],
                    "Investment Type": row_data[1],
                    "Amount Due": row_data[2],
                    "Proposed Date Of Transfer To IEPF" : row_data[3],
                    "Actual Date Of Transfer To IEPF" : "",
                    "Number of Shares": "",
                    "Nominal Value":""             
                }   
                print(item)
                self.list_of_scrapped_data.append(item)
        except Exception as e:
            pass

        #2
        try:
            try:
                xpath = sections_xpath[1]
                table = WebDriverWait(self.driver, 2).until(EC.presence_of_element_located((By.XPATH, xpath)))
            except:
                xpath = "//font[contains(., 'Amount Transfered to IEPF')]/following-sibling::table[1]"
                table = WebDriverWait(self.driver, 0.1).until(EC.presence_of_element_located((By.XPATH, xpath)))

            rows = table.find_elements(By.XPATH, ".//tbody/tr[@class='RowDataAlt' or @class='RowData']")
            for row in rows:
                row_data = [td.text for td in row.find_elements(By.XPATH, ".//td")]
                item = {
                    "Given Name": given_name,
                    "Inversor Name" : investor_name,
                    "Address" : address,
                    "Type" : "Amount Transfered to IEPF",
                    "Company Name": row_data[0],
                    "Investment Type": row_data[1],
                    "Amount Due": row_data[2],
                    "Proposed Date Of Transfer To IEPF" : "",
                    "Actual Date Of Transfer To IEPF" : row_data[3],
                    "Number of Shares": "",
                    "Nominal Value":""             
                }   
                print(item)
                self.list_of_scrapped_data.append(item)
        except Exception as e:
            pass

        #3
        try:
            xpath = sections_xpath[2]
            table = wait.until(EC.presence_of_element_located((By.XPATH, xpath)))
            rows = table.find_elements(By.XPATH, ".//tbody/tr[@class='RowDataAlt' or @class='RowData']")
            for row in rows:
                row_data = [td.text for td in row.find_elements(By.XPATH, ".//td")]
                item = {
                    "Given Name": given_name,
                    "Inversor Name" : investor_name,
                    "Address" : address,
                    "Type" : "Amount Refunded from IEPF",
                    "Company Name": row_data[0],
                    "Investment Type": row_data[1],
                    "Amount Due": row_data[2],
                    "Proposed Date Of Transfer To IEPF" : row_data[3],
                    "Actual Date Of Transfer To IEPF" : "",
                    "Number of Shares": "",
                    "Nominal Value":""             
                }   
                print(item)
                self.list_of_scrapped_data.append(item)
        except Exception as e:
            pass

        #4
        try:
            xpath = sections_xpath[3]
            table = wait.until(EC.presence_of_element_located((By.XPATH, xpath)))
            rows = table.find_elements(By.XPATH, ".//tbody/tr[@class='RowDataAlt' or @class='RowData']")
            for row in rows:
                row_data = [td.text for td in row.find_elements(By.XPATH, ".//td")]
                item = {
                    "Given Name": given_name,
                    "Inversor Name" : investor_name,
                    "Address" : address,
                    "Type" : "Shares Refunded from IEPF account",
                    "Company Name": row_data[0],
                    "Investment Type": "",
                    "Amount Due": "",
                    "Proposed Date Of Transfer To IEPF" : "",
                    "Actual Date Of Transfer To IEPF" : row_data[3],
                    "Number of Shares": row_data[1],
                    "Nominal Value":row_data[2]            
                }   
                print(item)
                self.list_of_scrapped_data.append(item)
        except Exception as e:
            pass

        #5
        try:
            xpath = sections_xpath[4]
            table = wait.until(EC.presence_of_element_located((By.XPATH, xpath)))
            rows = table.find_elements(By.XPATH, ".//tbody/tr[@class='RowDataAlt' or @class='RowData']")
            for row in rows:
                row_data = [td.text for td in row.find_elements(By.XPATH, ".//td")]
                item = {
                    "Given Name": given_name,
                    "Inversor Name" : investor_name,
                    "Address" : address,
                    "Type" : "Amount Refunded from IEPF",
                    "Company Name": row_data[0],
                    "Investment Type": row_data[1],
                    "Amount Due": row_data[2],
                    "Proposed Date Of Transfer To IEPF" : row_data[3],
                    "Actual Date Of Transfer To IEPF" : "",
                    "Number of Shares": "",
                    "Nominal Value":""             
                }    
                print(item)
                self.list_of_scrapped_data.append(item)
        except Exception as e:
            pass


        filename = 'Data.csv'
        if os.path.exists(filename):
            existing_df = pd.read_csv(filename)
            df = pd.concat([existing_df, pd.DataFrame(self.list_of_scrapped_data)], ignore_index=True)
        else:
            df = pd.DataFrame(self.list_of_scrapped_data)
        
        # Save DataFrame to CSV
        df.to_csv(filename, index=False)
        print(f"Data saved to {filename}")
                    

    def reaching_data(self,first_name,middle_name,last_name):
        wait = WebDriverWait(self.driver, 5)

        try:
            wait.until(EC.presence_of_element_located((By.XPATH, "(//input[@class='txt'])[1]")))
        except:
            return
        
        first_name_input = wait.until(EC.presence_of_element_located((By.XPATH, "(//input[@class='txt'])[1]")))
        first_name_input.clear()
        first_name_input.send_keys(first_name)
        middle_name_input = wait.until(EC.presence_of_element_located((By.XPATH, "(//input[@class='txt'])[2]")))
        middle_name_input.clear()
        middle_name_input.send_keys(middle_name)
        last_name_input = wait.until(EC.presence_of_element_located((By.XPATH, "(//input[@class='txt'])[3]")))
        last_name_input.clear()        
        last_name_input.send_keys(last_name)
        time.sleep(0.5)

        wait.until(EC.presence_of_element_located((By.XPATH, "(//input[@name='method'])[1]"))).click()

        links = wait.until(EC.presence_of_all_elements_located((By.XPATH, "//tbody/tr/td/a")))

        for i in links[3:]:
            i.click()
            #entered checkbox page

            self.driver.switch_to.window(self.driver.window_handles[2])

            given_name = f"{first_name} {middle_name} {last_name}"
            investor_name = wait.until(EC.presence_of_element_located((By.XPATH, "(//tr[@class='RowDataAlt'])/td[2]"))).text
            address = wait.until(EC.presence_of_element_located((By.XPATH, "(//tr[@class='RowDataAlt'])/td[3]"))).text

            checkboxes = wait.until(EC.presence_of_all_elements_located((By.XPATH, "//tbody/tr/td/input[@type='checkbox']")))
            for checkbox in checkboxes:
                checkbox.click()

            wait.until(EC.presence_of_element_located((By.XPATH, "(//input[@class='button1'])[1]"))).click()
            self.driver.switch_to.window(self.driver.window_handles[3])
            #entered data page


            self.scrap_data(given_name=given_name, investor_name=investor_name, address=address)

            while True:
                try:
                    next_button = wait.until(EC.presence_of_element_located((By.XPATH, "//span/a[contains(text(), 'Next')]")))
                    self.driver.execute_script("arguments[0].scrollIntoView(true);", next_button)
                    time.sleep(1)
                    next_button.click()

                    self.scrap_data(given_name=given_name, investor_name=investor_name, address=address)
                except Exception as e: 
                    break


            #exiting data page
            close_btn = wait.until(EC.presence_of_element_located((By.XPATH, "//input[@type='button']")))
            self.driver.execute_script("arguments[0].scrollIntoView(true);", close_btn)
            time.sleep(1)
            close_btn.click()
            self.driver.switch_to.window(self.driver.window_handles[2])


            #at last(Checkbox page)
            close_btn = wait.until(EC.presence_of_element_located((By.XPATH, "(//input[@class='button1'])[2]")))
            self.driver.execute_script("arguments[0].scrollIntoView(true);", close_btn)
            time.sleep(1)
            close_btn.click()
            self.driver.switch_to.window(self.driver.window_handles[1])

        
    def save_data(self):
        filename = 'Data.csv'
        if os.path.exists(filename):
            existing_df = pd.read_csv(filename)
            df = pd.concat([existing_df, pd.DataFrame(self.list_of_scrapped_data)], ignore_index=True)
        else:
            df = pd.DataFrame(self.list_of_scrapped_data)
        
        df.to_csv(filename, index=False)
        print(f"Data saved to {filename}")


    def main_workflow(self):
        txt_file_name = input("Enter the txt file name (for eg 'majmundar_spello') : ")
        file_path = f"{txt_file_name}.txt"

        self.config_driver()
        wait = WebDriverWait(self.driver, 5)

        url =  "https://www.iepf.gov.in/IEPFWebProject/SearchInvestorAction.do?method=gotoSearchInvestor"
        self.driver.get(url)
        time.sleep(1)
        
        wait.until(EC.presence_of_element_located((By.XPATH, "//a[@title='Query Unpaid and Unclaimed Amounts'][1]"))).click()

        # Get handles of all currently open windows
        self.driver.switch_to.window(self.driver.window_handles[1])


        try:
            with open(file_path, 'r') as file:
                for line in file:
                    words = line.strip().split()
                    
                    if len(words) == 1:
                        first_name = words[0]
                        middle_name = ""
                        last_name = ""
                    elif len(words) == 2:
                        first_name = words[0]
                        middle_name = ""
                        last_name = words[1]
                    elif len(words) == 3:
                        first_name = words[0]
                        middle_name = words[1]
                        last_name = words[2]

                    self.reaching_data(first_name, middle_name, last_name)
                    

        except Exception as e:
            print(e)
        
        finally:
            # self.save_data()

            self.driver.close()



url =  "https://www.iepf.gov.in/IEPFWebProject/SearchInvestorAction.do?method=gotoSearchInvestor"

data_scraper = Scraper()
data_scraper.main_workflow()