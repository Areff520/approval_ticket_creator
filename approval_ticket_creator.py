import datetime, time
import re
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
import os
import pandas as pd
import Blurbs

"""Takes vendor ID from the excel, takes the neccesarry details from smores-eu.corp
    then create approval request ticket via approvals.amazon.com"""


class Automation():

    def __init__(self):
        self.boktanlist = []

    def selenium(self):
        profile_name = 'Mail_Automation'
        profile_path = rf'{os.path.expanduser("~")}\Chrome Profiles\{profile_name}'
        chrome_options = webdriver.ChromeOptions()
        chrome_options.add_argument(f'user-data-dir={profile_path}')
        chrome_options.add_argument(f'--profile-directory={profile_name}')
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-gpu")
        chrome_options.add_experimental_option('excludeSwitches', ['enable-logging'])
        chrome_options.add_experimental_option("detach", True)
        browser = webdriver.Chrome(ChromeDriverManager().install(), options=chrome_options)

        return browser

    def get_vendorcode(self):
        df = pd.read_excel('RVR+Submitted+VC.xlsx')

        self.vendor_list = df.vendor_code.values.tolist()

    def check_scs_status(self,vendorID,browser):
        link = f'https://smores-eu.corp.amazon.com/vendor-onboarding/vendor-search-results?marketscope=TR&businessName=&owningBuyer=&vendorCode={vendorID}&createdBy=&createdFrom=&createdTo=&vendorGroupID=&globalSetupId='
        browser.get(link)
        WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.CLASS_NAME, "paginationHolder")))
        time.sleep(1)
        contents = browser.page_source
        contents = re.findall(f'(?:href)(.*?)(?:</a>)', contents)
        her_link = []
        the_link = []
        for value in contents:
            if 'https' in value:
                her_link.append(value)
        for value in her_link:
            correct_link = re.findall(f'(?:")(.*?)(?:")', value)
            correct_link = correct_link[0]
            the_link.append(correct_link)
        the_link = the_link[0]
        the_link = the_link.replace(';', '&')
        """
        Checking SCS
        """
        #the_link = 'https://smores-eu.corp.amazon.com/vendor-onboarding/setup/details?vendorGroupId=8088572&vendorSetupGroupId=63155812&realm=EU'
        browser.get(the_link)
        WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.CLASS_NAME, "vendorSetupModuleContent")))
        element = browser.find_elements(By.CLASS_NAME, 'vendorSetupModuleContent')

        for elements in element:
            if 'SCS (Purchasing Terms)' in elements.text:
                print('SCS contenti', elements.text)
                if 'Pending Amazon Action' in elements.text:
                    Continue=True
                    print('oh yeah its terrible')
                elif 'Completed' in elements.text:
                    Continue='Check Vendor Onboarding Approval '
                else:
                    Continue=False
        print(Continue)
        WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.XPATH, "//div[@class='vendorSetupDetailContent']//kat-table-cell[@role='cell']")))
        element = browser.find_element(By.XPATH,"//div[@class='vendorSetupDetailContent']//kat-table-cell[@role='cell']")
        scs_link = element.find_element(By.TAG_NAME, "a")
        scs_link = scs_link.get_attribute('href')

        """ Getting Primary group id and vendor code ID and Owning Buyer"""
        vendor_groupID=browser.find_element(By.XPATH,"//span[contains(text(),'Vendor Group ID:')]").find_element(By.XPATH,'..')
        vendor_groupID=vendor_groupID.find_element(By.TAG_NAME,'a').text

        GL = browser.find_element(By.XPATH, "//span[contains(text(),'Primary Business Group:')]").find_element(By.XPATH, '..')
        GL=GL.text.replace("Primary Business Group: ",'')

        OB= browser.find_element(By.XPATH, "//span[contains(text(),'Owning Buyer:')]").find_element(By.XPATH,'..').find_element(By.TAG_NAME,'a').text.replace('','')
        print(OB)

        return Continue,scs_link,vendor_groupID,GL,OB

    def check_vendor_onboarding(self,browser):
        try:
            vendor_onboarding_status=browser.find_element(By.XPATH,"//span[contains(text(),'Vendor Onboarding Approval')]")
            vendor_onboarding_status=vendor_onboarding_status.find_element(By.XPATH,"..")
            l7_alias=vendor_onboarding_status.find_element(By.XPATH,"..")
            if 'Pending Approval' in vendor_onboarding_status.text:
                status='continue'
            else:
                status='skip'

            actions = ActionChains(browser)
            l7_alias=browser.find_element(By.ID,'approvalOverview')
            actions.move_to_element(l7_alias).perform()
            time.sleep(2)
            l7_alias=browser.find_element(By.ID,'approvalOverview')
            l7_alias=l7_alias.find_element(By.TAG_NAME,'a').text
            #l7_alias=l7_alias.get_attribute('href')
            l7_alias=re.findall('users/\w+',l7_alias)
            l7_alias=l7_alias[0].replace('users/','')
            print(len(l7_alias))
            print(l7_alias)
        except Exception as e:
            status='skip'
            l7_alias=None
            print('Excaption at check vendor onboarding;',e)
        return status,l7_alias



    def get_alias(self, browser, scs_link):
        browser.get(scs_link)
        WebDriverWait(browser, 20).until(EC.presence_of_element_located((By.CLASS_NAME, "ng-scope")))
        #time.sleep(10)
        # Şuraya for looplu try ekle
        aliasa_dogru = browser.find_element(By.CLASS_NAME, "ng-scope")

        WebDriverWait(aliasa_dogru, 10).until(EC.presence_of_element_located((By.XPATH, "//li[@class='nav-item ng-scope']")))
        aliasa_dogru = aliasa_dogru.find_element(By.XPATH, "//li[@class='nav-item ng-scope']")
        aliasa_dogru.click()

        WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.CLASS_NAME, 'list-group')))
        time.sleep(2)
        alias = browser.find_element(By.CLASS_NAME, 'list-group')
        alias = alias.text.split()
        alias = alias[0]
        print(alias)
        return alias

    def approval_page(self, browser, alias, link, vendorID, gl, ob, summary, vendorgroupid):

        """
        Approvals page, creates the ticket

        """


        scs_title=Blurbs.scs_title
        scs_title=scs_title.replace("vendor_id",vendorID).replace("gl",gl).replace("ob",ob).replace('\n','')
        summary_is=summary.replace("vendorid",vendorID).replace("GL_IS",gl).replace("scs_link",link).replace('vendor_groupid',vendorgroupid)

        link = 'https://approvals.amazon.com/approval/create'
        browser.get(link)
        title = browser.find_element(By.ID, 'txtTitle')
        title.send_keys(scs_title)
        """Switching to iframe for summary"""

        b=browser.find_element(By.XPATH,'//iframe[@title="Rich Text Editor, editor1"]')
        browser.switch_to.frame(b)
        summary = browser.find_element(By.XPATH, '//body[@contenteditable="true"]')
        summary.send_keys(summary_is)

        date_approval = datetime.datetime.now().date() + datetime.timedelta(days=7)
        date_approval = date_approval.strftime('%m/%d/%Y')
        browser.switch_to.default_content()
        date=browser.find_element(By.XPATH,'//input[@placeholder="MM/DD/YYYY"]')
        date.send_keys(Keys.CONTROL,"a", Keys.DELETE)
        date.send_keys(date_approval)

        #priority=browser.find_element(By.XPATH,'//span[@id="awsui-select-0-textbox"]//span[@class="awsui-select-trigger-label"]')
        priority=browser.find_element(By.XPATH,'//div[@awsui-form-section-region="content"]//div[@class="awsui-grid awsui-form-field-controls"]//div[@role="button"]')

        priority.click()

        browser.find_element(By.XPATH,"//span[contains(text(),'Medium')]").click()

        button = browser.find_element(By.XPATH, '//div[@class="awsui-util-container"]//span[@awsui-button-region="text"]')
        button.click()

        """ Approvals ID seçme """
        WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.XPATH, '//awsui-input[@id="searchText_input"]')))
        approval = browser.find_element(By.XPATH, '//awsui-input[@id="searchText_input"]')
        approval = approval.find_element(By.ID, 'awsui-input-2')
        actions = ActionChains(browser)
        actions.move_to_element(approval).perform()
        approval.send_keys('ogyilmap')

        """ add """
        WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.XPATH, '//awsui-form-section[@class="approvalSequenceContainer"]//div[@class="searchResultContainer"]')))
        time.sleep(2)
        add = browser.find_element(By.XPATH,'//awsui-form-section[@class="approvalSequenceContainer"]//div[@class="searchResultContainer"]')
        add = add.find_element(By.TAG_NAME, 'a')
        actions.move_to_element(add).perform()
        add.click()
        add = browser.find_element(By.ID, 'group-link')
        add.click()

        """Clicking No"""

        no_element=browser.find_element(By.ID,'awsui-radio-button-2-label')
        no_element.click()

        """Clicking Submit"""
        time.sleep(2)
        submit_button=browser.find_element(By.XPATH,'//section[@id="btnSubmitApproval"]//button[@type="submit"]')
        actions.move_to_element(submit_button).perform()
        submit_button.click()


    def process_automation(self):
        self.get_vendorcode()
        browser = self.selenium()

        for vendorID in self.vendor_list:
            print('Working on vendor ID ;',vendorID)
            Continue, scs_link, vendor_groupID, GL, OB=self.check_scs_status(vendorID,browser)
            if Continue==True:
                summary = Blurbs.scs_summary
                VM_ID = self.get_alias(browser,scs_link)
                self.approval_page(browser,VM_ID,scs_link,vendorID,GL,OB,summary,vendor_groupID)
                print('DONE_1')
            elif Continue=='Check Vendor Onboarding Approval ':
                summary=Blurbs.onboarding_summary
                status, l7_alias = self.check_vendor_onboarding(browser)
                link=browser.current_url
                self.approval_page(browser,l7_alias,link,vendorID,GL,OB,summary,vendor_groupID)
            else:
                print(f'Skipping {vendorID}')
            time.sleep(10)



if __name__ == "__main__":
    obj = Automation()
    print("Starting Andon Flip Automation at " + str(datetime.datetime.now()))
    obj.process_automation()
    print("Ending Andon Flip Automation at " + str(datetime.datetime.now()))
