"""
third_eye_nav.py

This module contains functions for navigating across and interacting with
thirdeye.com. It's purpose is to help automate tasks for General Agents 
Acceptance Coorporation.

Dependencies:
- Selenium
- Chrome Web Browser executbale on your system 
"""

"""
TODO
    1) add web scrape functionality (phone number)
    2) special memo
"""

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import StaleElementReferenceException
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
import os
import time
import sys

# time in seconds that the selenium driver will wait for a webpage to resolve
# this may need to be adjusted for slow internet connection 
error_wait_time = 10

# URL's
LOGIN_PAGE = "https://gaac.thirdeyesys.ca/insight/"
SEARCH_PAGE = "https://gaac.thirdeyesys.ca/insight/ControllerServlet?action=425&guid=1734391617932"
ACCOUNTING_REPORTS_PAGE = "https://gaac.thirdeyesys.ca/insight/ControllerServlet?action=293&guid=1733370000893"
MANAGEMENT_REPORTS_PAGE = "https://gaac.thirdeyesys.ca/insight/ControllerServlet?action=295&guid=1734375935562"

# download directory 
DOWNLOAD_DIR = os.path.join(os.path.expanduser('~'), 'Downloads')

class ThirdEyeNav:
    '''
    Description: 
        Contructor that initializes a chrome web driver and a web
        driver wait object. The selenium web driver grants access to web 
        elements and allows the user to interact programitcally with those web
        elems. 

    Input: 
        [headless] - a string indicating whether or not the web driver
            should be displayed to the screen
        [username] - Username for Third Eye account 
        [password] - Password for Third Eye account
    '''
    def __init__(self, headless, username, password):
        # set data members
        self.open_driver()
        self.download_dir = DOWNLOAD_DIR
        self.username = username
        self.password = password
        self.is_logged_in = False
        self.current_page=None 
    '''
    Description:
        Release the web browser driver and reset session information
    '''
    def __del__(self):
        # Cleanup when the object is destroyed
        self.close_driver()

    def close_driver(self):
        # Explicitly close the Selenium WebDriver
        if self.driver:
            self.driver.close()
            self.driver = None
        # Reset other attributes
        self.is_logged_in = False
        self.current_page = None
    
    def open_driver(self):
        self.driver = driver_factory('head')
        self.wait = WebDriverWait(self.driver, error_wait_time)

    def login(self):
        '''
        Description: 
            Populates the LoginID and LoginPassword input fields. Clicks the login
            element.

        Input:
            [username] - username as a string
            [password] - password as a string
        Output:
            Return true if successful. Otherwise, false
        '''
        # if user is already logged in, return
        if self.is_logged_in is True:
            return True

        try:
            # travel to the login page
            self.navigate_to("Login Page")

            # find login elements
            login_ID  = self.wait.until(EC.presence_of_element_located((By.NAME,"LoginId")))
            login_pwd = self.wait.until(EC.presence_of_element_located((By.NAME,"LoginPassword")))
        except TimeoutException:
            print("Timeout Error. Failed to login")
            return False
        except:
            print("Cannot find web elements... Are you on the login screen?")
            return False
        else:
            # send key strokes 
            login_ID.clear()
            login_pwd.clear()
            login_ID.send_keys(self.username)
            login_pwd.send_keys(self.password)
            self.wait.until(EC.presence_of_element_located((By.NAME,"login"))).click()
            
        # test for successful login
        try:
            self.wait.until(EC.presence_of_element_located((By.XPATH,"//*[contains(.,'Admin Options')]")))
        except TimeoutException:
            print("Unsuccessful Login")
            return False
        else:
            self.is_logged_in = True
        
        return True
    
    def navigate_to(self, dest):
        '''
        Description: 
            Helps navigate to certain pages in Third Eye. For some moves, 
            self.driver is expected to ALREADY be set to a certain page
            'Collection Page' - Needs to be on an account page
            'Memo Screen' - Needs to be on an account page
        Input:  
            [dest] - The destination web page to be travelled to.
        Output: 
            Sets self.driver to the new page and returns true if successful. 
            Otherwise, false
        TODO - check elems to make sure user is on correct page and that is loaded. 
        '''      
        try: 
            if dest == 'Login Page' and self.current_page != 'Login Page':
                self.driver.get(LOGIN_PAGE)
            
            elif dest == 'Search Page' and self.current_page != 'Search Page':
                self.driver.get(SEARCH_PAGE)
        
            elif dest == 'Memo Screen' and self.current_page != 'Memo Screen':
                self.wait.until(EC.presence_of_element_located((By.XPATH,"//body"))).send_keys(Keys.ALT, 'M')
                self.wait.until(EC.presence_of_element_located((By.ID,"memoFunctions"))).click()

            elif dest == 'Accounting Reports Page' and self.current_page != 'Accounting Reports Page':
                self.driver.get(ACCOUNTING_REPORTS_PAGE)
            
            elif dest == 'Management Reports Page' and self.current_page != 'Management Reports Page':
                self.driver.get(MANAGEMENT_REPORTS_PAGE)
            
            elif dest == 'Collection Page' and self.current_page != 'Collection Page':
                self.wait.until(EC.presence_of_element_located((By.ID,"top5"))).click()

            elif dest == 'Collection Memo Page' and self.current_page != 'Collection Memo Page':
                self.wait.until(EC.presence_of_element_located((By.ID,"top5"))).click()
                self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR,"input[title='Reminder Call']"))).click()
                
                 
        except TimeoutException:
            print("Failed to navigate to " + dest)
            return False
        except:
            print("Unknown Error in navigate_to")
            return False
        
        self.current_page = dest
        return True
           
    def memo_account(self, account_number, memo_subject, memo_body, insured_name = None):
        '''
        Description: 
            Memo's an account in Third Eye. 
        Input:  
            [account_numbers] - list of account numbers. MWF prefix is optional {MWF99999,99999}              
            [memo_subject]    - The subject as a string
            [memo_body]       - The body as a string
            [insured_name]    - Optional. Used to do an additional check 
        Output: 
            Returns true if all the account was memo'd succsesfully. Otherwise,
            return false.    
        '''
        # move to search page, if not already there
        if not self.navigate_to("Search Page"):
            return false
        
        # search for account and move to memo screen
        self.search_account(account_number)
        self.navigate_to('Memo Screen')
        
        # write memo
        try:
            memoSubject = self.wait.until(EC.presence_of_element_located((By.ID,"memoSubject")))
            memoBody = self.wait.until(EC.presence_of_element_located((By.ID,"memoBody")))
            memoSubject.clear()
            memoBody.clear()
            memoSubject.send_keys(memo_subject)
            memoBody.send_keys(memo_body)
        except TimeoutException:
            print("Could not find the memo's subject line/body")
            return False
        except:
            print("Unknown Error in memo_account")
            return False
        
        # save memo
        try:
            save_memo = self.wait.until(EC.presence_of_element_located((By.XPATH,'//*[@title="Save Memo"]')))
        except TimeoutException:
            print("Could not find 'save memo' button")
            return False
        except:
            print("Unknown Error")
            return False
        else:
            save_memo.click()
            #print("Save succussful")
            return True
        
    def memo_account_collection(self, account_number, memo_subject, date):
        # move to search page, if not already there
        if not self.login():
            return False

        if not self.navigate_to("Search Page"):
            return False
        
        # search for account and move to memo screen
        self.search_account(account_number)
        self.navigate_to('Collection Memo Page')

        try:
            memoSubject = self.wait.until(EC.presence_of_element_located((By.NAME,"ContractCollectionLetter_Memo0")))
            callDoneDate = self.wait.until(EC.presence_of_element_located((By.NAME,"ContractCollectionLetter_CallDoneDate0")))
            save = self.wait.until(EC.presence_of_element_located((By.XPATH,'//*[@title="Save"]')))

            callDoneDate.clear()
            memoSubject.clear()
            #callDoneDate.send_keys(date)
            memoSubject.send_keys(memo_subject)

            save.click()
        except TimeoutException:
            return False
        except Exception as e:
            print(e) # write to stdout to avoid program failure
            return False
        
        return True

    def download_report(self, report, l_date=None, r_date=None):
        '''
        Description:
            Downloads the specified report. The report should appear in the users
            download file. The naming convention for the files is 
            'ControllerServlet'. This module cannot guarantee the exact file name
            as it appears in the users filesystem
        Input:
            [report] - The name of the target file. So far this function supports
                1. Check Register
                2. Collection Calls
        Output:
            Returns the expected filepath of the download. The operating system may 
            suffix the name with (1), (2), etc. The calling function will need to 
            decipher the correct file name.
        '''
        try:
            if report == 'Check Register':
                if l_date or r_date:
                    raise ValueError('Dates are required for Check Register report')

                # travel to the accounts report page
                self.navigate_to('Accounting Reports Page')

                # locate the dropdown object
                select_elem = self.wait.until(EC.presence_of_element_located((By.NAME,'availableReportList')))
                dropdown = Select(select_elem)

                # make selection
                dropdown.select_by_visible_text('Check Register')

                # locate the from date, to date and generate report elements
                from_date = self.wait.until(EC.presence_of_element_located((By.ID,'FromDate')))
                to_date = self.wait.until(EC.presence_of_element_located((By.ID,'ToDate')))
                create_report = self.wait.until(EC.presence_of_element_located((By.XPATH,'//*[@title="Create Report"]')))

                # send data
                from_date.clear()
                to_date.clear()
                from_date.send_keys(l_date)
                to_date.send_keys(r_date)
                create_report.click()

            elif report == 'Collection Report':
                # travel to the management report page
                self.navigate_to('Management Reports Page')
                
                # locate the dropdown object
                select_elem = self.wait.until(EC.presence_of_element_located((By.NAME,'availableReportList')))
                dropdown = Select(select_elem)

                dropdown.select_by_visible_text('Collection Report')

                create_report = self.wait.until(EC.presence_of_element_located((By.XPATH,'//*[@title="Create Reports"]')))
                create_report.click()
                time.sleep(10)
            else:
                raise Exception("Invalid arguments in function download_report")
        except Exception as e:
            print(e)
            return ''

        return self.download_dir
    
    def search_account(self,account_number):
        '''
        Description:
            This function is used to simulate searching for a contract. It is designed
            to be a helper function and used in conjunction with other methods to
            either scrape or send data to a contract's data page
        In:
            [account_number] - The account number to search for
        Out:
            True if the contract was found, otherwise false.
        '''
        # check for empty string
        if not account_number:
            return False
        
        # move to the search page if necessary
        if not self.navigate_to('Search Page'):
            return False

        # input contract number into search bar and search
        try:
            contractSearch = self.wait.until(EC.presence_of_element_located((By.NAME,"VISIBLE_ContractNo")))
            contractSearch.clear()
            contractSearch.send_keys(account_number)
            self.wait.until(EC.presence_of_element_located((By.NAME,"quoteSearchContractByAll"))).click()
            self.current_page = "Account " + account_number

            # if the following element is resolved then the contract was found
            self.wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/form[2]/div/table[2]/tbody/tr/td/table/tbody/tr[3]/td[2]/table/tbody/tr[2]/td[2]/span[1]/div[1]/table/tbody/tr[1]/td[1]")))
        except TimeoutException:
            print("Could not find search bar or search failed")
            return False
        except:
            print("Unknown Error")
            return False
        return True

    def get_info(self,contract, desired_info):
        '''
        Description:
            This method is designed to web scrape insured (client) info
        Input:
            [contract] - The contract to search for
            [desired_info] - the desired information as a bool array [insured name, insured phone number, agent name, loan group, 
                next payment amount, default date, cancellation date, regular payment, late payment, next payment,
                current amt due, agent phone, insured mailing address]
        Output:
            returns the desired info as a dictionary
        '''
        if not self.search_account(contract):
            pass # TODO 

        map = {}
        try:
            # insured
            if desired_info[0] is True:
                elem = self.wait.until(EC.presence_of_element_located((By.XPATH,"/html/body/form[2]/div/table[2]/tbody/tr/td/table/tbody/tr[3]/td[2]/table/tbody/tr[2]/td[2]/span[1]/div[1]/table/tbody/tr[2]/td[3]")))
                map['insured'] = elem.text

            # phone number
            if desired_info[1] is True: 
                elem = self.wait.until(EC.presence_of_element_located((By.XPATH,"/html/body/form[2]/div/table[2]/tbody/tr/td/table/tbody/tr[3]/td[2]/table/tbody/tr[2]/td[2]/span[1]/div[1]/table/tbody/tr[3]/td[3]/span")))
                map['phone'] = (elem.text)[:13]
            
            # Agent
            if desired_info[2] is True:
                elem = self.wait.until(EC.presence_of_element_located((By.XPATH,"/html/body/form[2]/div/table[2]/tbody/tr/td/table/tbody/tr[3]/td[2]/table/tbody/tr[2]/td[2]/span[1]/div[3]/table/tbody/tr[1]/td[3]")))
                map['agent'] = (elem.text)
            
            # loan group
            if desired_info[3] is True:
                elem = self.wait.until(EC.presence_of_element_located((By.XPATH,"/html/body/form[2]/div/table[2]/tbody/tr/td/table/tbody/tr[3]/td[2]/table/tbody/tr[2]/td[2]/span[1]/div[3]/table/tbody/tr[3]/td[3]")))
                map['loangroup'] = (elem.text)

            # next payment due date 
            if desired_info[4] is True:
                elem = self.wait.until(EC.presence_of_element_located((By.XPATH,"/html/body/form[2]/div/table[2]/tbody/tr/td/table/tbody/tr[3]/td[2]/table/tbody/tr[2]/td[2]/span[1]/div[7]/div[2]/table/tbody/tr[2]/td[2]")))
                map['nextpaymentdate'] = (elem.text)

            # NOI (default) date 
            if desired_info[5] is True:
                elem = self.wait.until(EC.presence_of_element_located((By.XPATH,"/html/body/form[2]/div/table[2]/tbody/tr/td/table/tbody/tr[3]/td[2]/table/tbody/tr[2]/td[2]/span[1]/div[7]/div[2]/table/tbody/tr[3]/td[2]")))
                map['defaultdate'] = (elem.text)

            # cancellation date 
            if desired_info[6] is True:
                elem = self.wait.until(EC.presence_of_element_located((By.XPATH,"/html/body/form[2]/div/table[2]/tbody/tr/td/table/tbody/tr[3]/td[2]/table/tbody/tr[2]/td[2]/span[1]/div[7]/div[2]/table/tbody/tr[4]/td[2]")))
                map['canceldate'] = (elem.text)

            # regular payment amount 
            if desired_info[7] is True:
                elem = self.wait.until(EC.presence_of_element_located((By.XPATH,"/html/body/form[2]/div/table[2]/tbody/tr/td/table/tbody/tr[3]/td[2]/table/tbody/tr[2]/td[2]/span[1]/div[7]/div[2]/table/tbody/tr[6]/td[2]/b/span[1]")))
                map['payamt'] = (elem.text)

            # late payment amount 
            if desired_info[8] is True:
                elem = self.wait.until(EC.presence_of_element_located((By.XPATH,"/html/body/form[2]/div/table[2]/tbody/tr/td/table/tbody/tr[3]/td[2]/table/tbody/tr[2]/td[2]/span[1]/div[7]/div[2]/table/tbody/tr[7]/td[2]")))
                map['latepayamt'] = (elem.text)

            # next payment due amount 
            if desired_info[9] is True:
                elem = self.wait.until(EC.presence_of_element_located((By.XPATH,"/html/body/form[2]/div/table[2]/tbody/tr/td/table/tbody/tr[3]/td[2]/table/tbody/tr[2]/td[2]/span[1]/div[7]/div[2]/table/tbody/tr[9]/td[2]")))
                map['nextpaydate'] = (elem.text)

            # current due amount 
            if desired_info[10] is True:
                elem = self.wait.until(EC.presence_of_element_located((By.XPATH,"/html/body/form[2]/div/table[2]/tbody/tr/td/table/tbody/tr[3]/td[2]/table/tbody/tr[2]/td[2]/span[1]/div[7]/div[2]/table/tbody/tr[10]/td[2]")))
                map['due'] = (elem.text)

            # agent phone 
            if desired_info[11] is True:
                elem = self.wait.until(EC.presence_of_element_located((By.XPATH,"/html/body/form[2]/div/table[2]/tbody/tr/td/table/tbody/tr[3]/td[2]/table/tbody/tr[2]/td[2]/span[1]/div[3]/table/tbody/tr[2]/td[3]")))
                map['agentphone'] = (elem.text)[:14]

            # insured address
            if desired_info[12] is True:
                self.wait.until(EC.presence_of_element_located((By.ID,"top1"))).click()
                elem = self.wait.until(EC.presence_of_element_located((By.XPATH,"/html/body/form[2]/div/table[2]/tbody/tr/td/table/tbody/tr[3]/td[2]/table/tbody/tr[2]/td[2]/span[2]/table[1]/tbody/tr[3]/td[2]")))
                map['address'] = (elem.text)

        except IndexError:
            pass
        except Exception as e:
            return e

        return map

    def search_for_mail(self):
        '''
        Description: 
            Searches memos to see if anything has been mailed to a client Expects
            to be on an account screen.
        TODO NOT TESTED
        '''
        
        try:
            #go to notices tab
            self.wait.until(EC.presence_of_element_located((By.ID,"top6"))).click() 
            notices_table = self.wait.until(EC.presence_of_element_located((By.XPATH,"/html/body/form[2]/div/table[2]/tbody/tr/td/table/tbody/tr[3]/td[2]/table/tbody/tr[2]/td[2]/span[7]/table[1]")))           
        except TimeoutException:
            print("Could not find search bar")
            return False
        except:
            print("Unknown Error")
            return False
        
        lines = notices_table.text.split('\n')
        last_sent = lines[-1]
        index_of_date = last_sent.rfind(" ")
        return (last_sent[:index_of_date] + ',' + last_sent[index_of_date+1:])


def driver_factory(headless: str):
    # configure chrome web browser for automation
    chrome_options = webdriver.ChromeOptions()
    prefs = {
        "plugins.always_open_pdf_externally": True,
        "download.default_directory": DOWNLOAD_DIR,
        "download.prompt_for_download": False,  # Disable the prompt
        "download.directory_upgrade": True,  # Automatically overwrite existing files
        "safebrowsing.enabled": True,  # Enable safe browsing to avoid issues
    }
    chrome_options.add_experimental_option("prefs", prefs)
    if headless == 'headless':
        chrome_options.add_argument('--headless')

    # return web driver
    return webdriver.Chrome(options=chrome_options)