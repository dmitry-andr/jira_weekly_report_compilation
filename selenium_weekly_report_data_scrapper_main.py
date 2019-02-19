import time
import xlsxwriter
import os

from datetime import datetime

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException
from nltk import app
from _tracemalloc import start

#custom modules imports
import utils



def check_exists_by_xpath(browser, xpath):
    try:
        browser.find_element_by_xpath(xpath)
    except NoSuchElementException:
        return False
    return True

def check_exists_by_css_selector(browser, cssSelector):

    #DEBUG data print
    #print("browser", browser)
    #print("cssSelector", cssSelector)

    try:
        browser.find_element_by_css_selector(cssSelector)
    except NoSuchElementException:
        return False
    return True



def parsePageData(browser):
    print("Parsing page data started ...")
    tableRows = browser.find_elements_by_css_selector("table.aui tbody tr")
    pageData = []
    for row in tableRows:
        #print("processing : " + str(row.get_attribute('innerHTML')))
        issueType = row.find_element_by_css_selector(".issue-link img").get_attribute('alt')
        issueID = row.find_element_by_css_selector(".issuekey a").get_attribute('innerHTML')
        issueStatus = row.find_element_by_css_selector(".status span").get_attribute('innerHTML')
        issueCreatedDate = row.find_element_by_css_selector("td.created time").get_attribute('innerHTML')
        issueUpdatedDate = row.find_element_by_css_selector("td.updated time").get_attribute('innerHTML')
        print("issueType : " + issueType + " ; " + "issueID : " + issueID + " ; " + "issueStatus : " + issueStatus + " ; " + "issueCreatedDate : " + issueCreatedDate + " ; " + "issueUpdatedDate : " + issueUpdatedDate)
        dataRow = [issueID, issueType, issueStatus, issueCreatedDate, issueUpdatedDate]
        pageData.append(dataRow)
    print("Parsing page data - Complete")

    return pageData



def runJQLstatement(browser, jqlStatement):
    adv_search_selector = "#advanced-search"
    if(check_exists_by_css_selector(browser, adv_search_selector)):
        print("Advanced search input enabled")
    else:
        print("Jira in 'Basic' search mode. Switching to advanced")
        selector = ".switcher-item"
        enableAdvancedModeElement = browser.find_element_by_css_selector(selector)
        enableAdvancedModeElement.click()
        time.sleep(utils.TIMER_INTERVAL_SHORT)
        print("Jira in 'Advanced' search mode.")
        

    if(check_exists_by_css_selector(browser, adv_search_selector)):
        print("JQL request execution started >>> " + jqlStatement)
        advSearchInputElement = browser.find_element_by_css_selector(adv_search_selector)
        advSearchInputElement.clear()
        advSearchInputElement.send_keys(jqlStatement)
        advSearchInputElement.send_keys(Keys.ENTER)
        time.sleep(utils.TIMER_INTERVAL)

        #Read number of items returned for provided JQL. If nothing found - return empty data structure
        num_of_items_selector = "div.aui-item span.results-count-total"
        if(check_exists_by_css_selector(browser, num_of_items_selector)):
            totalNumOfItemsElement = browser.find_element_by_css_selector(num_of_items_selector)
            numOfItems = totalNumOfItemsElement.get_attribute('innerHTML')
            print("Number of records returned for provided JQL : " + str(numOfItems))
        else:
            print("No data returned for provided JQL")
            return None


        jqlResultData = parsePageData(browser)
        #check if many items and pagination exists
        pagination_selector = "div.pagination"
        if(check_exists_by_css_selector(browser, pagination_selector)):
            next_page_elemement_selector = "div.pagination a.icon-next"
            while check_exists_by_css_selector(browser, next_page_elemement_selector):
                print("Many items - need to process pagination")
                print("Moving to next page")
                nextPageLinkElement = browser.find_element_by_css_selector(next_page_elemement_selector)
                nextPageLinkElement.click()
                time.sleep(utils.TIMER_INTERVAL)
                jqlResultData.extend(parsePageData(browser))

        else:
            print("ERR !!!: Jira remains in 'Basic' search mode. Script can't be used for reporting generation")
        

    else:
        print(MSG_ELEMENT_NOT_FOUND)
    
    print("runJQLstatement completed")
    return jqlResultData


def writeDataToReportFile(dataToWrite):
    reportTextFile = utils.REPORT_DATA_FOLDER + "\\" + utils.REPORT_DATA_FILE
    print("Writing to file : " + reportTextFile)
    with open(reportTextFile, "a", encoding='utf-8') as text_file:
        print(str(dataToWrite), file=text_file)

    return 1



# MAIN PROGRAM
print("Starting Selenium script to gather weekly report data from Jira")
utils.initSettings()

options = webdriver.ChromeOptions()
options.add_argument("--lang=en")

browser_jira = webdriver.Chrome(utils.WEBDRIVER_PATH, chrome_options=options)

# Make request - load page
browser_jira.get(utils.JIRA_LOGIN_URL)

#Login
print("Attempting to login ...")
time.sleep(utils.TIMER_INTERVAL)
user_name_selector = "#username"
if(check_exists_by_css_selector(browser_jira, user_name_selector)):
    loginFormElement = browser_jira.find_element_by_css_selector(user_name_selector)
    print("Login controls exist, entering account params")
    loginFormElement.send_keys(utils.USER_NAME)
    loginFormElement.send_keys(Keys.ENTER)
    time.sleep(utils.TIMER_INTERVAL_SHORT)
    passw_selector = "#password"
    loginFormElement = browser_jira.find_element_by_css_selector(passw_selector)
    loginFormElement.send_keys(utils.PASSW)
    loginFormElement.send_keys(Keys.ENTER)
else:
    print(utils.MSG_ELEMENT_NOT_FOUND)
    print("NOTE!!! If no errors reported  and program continues running, ignore this message as user already logged in.")
print("Login complete")
time.sleep(utils.TIMER_INTERVAL)
print("Change browser url to issues search")
browser_jira.get(utils.JIRA_ISSUES_SEARCH_URL)
time.sleep(utils.TIMER_INTERVAL)


'''
print("*** DEBUG DATA ***")
print(runJQLstatement(browser_jira, "project=projName and type not in (Sub-task, Test) and status not in(Canceled)"))
print("*** END OF DEBUG DATA ***")
'''

print("Report compilation starts")



print("Scope summary(# of items) - Project scope for sprint")
writeDataToReportFile("Scope summary(# of items) - Project scope for sprint")
for project in utils.PROJECTS_LIST:
    scope_summary_JQL = "Sprint in (%s) AND type not in(Sub-task, Test) and project='%s'" #"Sprint in (288, 280) AND type not in(Sub-task, Test) and project=ProjName"
    jqlRunOutput = runJQLstatement(browser_jira, (scope_summary_JQL % (utils.SPRINT_IDS_LIST, project)))
    print(jqlRunOutput)
    reportVal = len(jqlRunOutput) if (jqlRunOutput != None) else 0
    writeDataToReportFile(project + " : " + str(reportVal))


print("Delivered(closed) items per week")
writeDataToReportFile("Delivered(closed) items per week")
for project in utils.PROJECTS_LIST:
    delivered_ites_per_week_JQL = "project = '%s' AND status in (Closed, Done) and type not in (Sub-task, Test) and (updatedDate >= '%s' AND updatedDate <= '%s')" #project = "ProjName" AND status in (Closed, Done) and type not in (Sub-task, Test) and (updatedDate >= "2019/01/28" AND updatedDate <= "2019/02/01")
    jqlRunOutput = runJQLstatement(browser_jira, (delivered_ites_per_week_JQL % (project, utils.START_DATE_OF_WEEK, utils.END_DATE_OF_WEEK)))
    print(jqlRunOutput)
    reportVal = len(jqlRunOutput) if (jqlRunOutput != None) else 0
    writeDataToReportFile(project + " : " + str(reportVal))
    if (jqlRunOutput != None):
        timeToCloseList = []
        for jqlRow in jqlRunOutput or []:
            timeToCloseList.append(utils.periodForDatesInDays(jqlRow[3], jqlRow[4], utils.SCRAPPED_DATES_FORMAT))
        averTimeToClose = sum(timeToCloseList)/len(timeToCloseList)
        print("Close time(days) : " + str(averTimeToClose) + " (" + str(min(timeToCloseList)) + " , " + str(max(timeToCloseList)) + ")")
        writeDataToReportFile(project + " : Close time Average(min, max) (days) : " + " : " + str(averTimeToClose) + " (" + str(min(timeToCloseList)) + " , " + str(max(timeToCloseList)) + ")")
    else:
        writeDataToReportFile("Close time (days) : 0")

'''Commented out as must be run only in prod mode due to possible size of returning data.
print("Backlog size")
writeDataToReportFile("Backlog size")
for project in utils.PROJECTS_LIST:
    backlog_size_JQL = "project = '%s' AND status in ('To Do', 'In Analysis', 'In Review', 'On Hold', 'Open Issue', Open, Backlog) AND type not in (Sub-task, Test)" #project = ProjName AND status in ("To Do", "In Analysis", "In Review", "On Hold", "Open Issue", Backlog) AND type not in (Sub-task, Test)
    jqlRunOutput = runJQLstatement(browser_jira, (backlog_size_JQL % (project)))
    print(jqlRunOutput)
    reportVal = len(jqlRunOutput) if (jqlRunOutput != None) else 0
    writeDataToReportFile(project + " : " + str(reportVal))
'''




browser_jira.close()

print("Selenium script - FINISHED")
