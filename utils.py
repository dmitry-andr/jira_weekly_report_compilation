from datetime import datetime
from time import gmtime, strftime
import configparser
import os

#JIRA connection params
JIRA_LOGIN_URL = None
JIRA_ISSUES_SEARCH_URL = None

USER_NAME = None
PASSW = None
#**********************

#Report JQL params
START_DATE_OF_WEEK = None
END_DATE_OF_WEEK = None
SPRINT_IDS_LIST = None
PROJECTS_LIST = []
TEAM_LIST = []
#*****************

#Script runtime params
WEBDRIVER_PATH = None
MSG_ELEMENT_NOT_FOUND = "ERR : Element not found"
TIMER_INTERVAL = None
TIMER_INTERVAL_SHORT = None

REPORT_TYPES_CMD_LINE_ARGS = ["all-in development!!!", "mbrprojstat", "scpsumsprint", "itmsweek", "backlogsize"]

DATES_FORMAT = "%Y-%m-%d"  #2018-08-24
TIME_FORMAT = "%H:%M:%S"
SCRAPPED_DATES_FORMAT = "%d/%b/%y"#13/Aug/18
REPORT_DATA_FOLDER = "generated_data"
TEAM_MEMBERS_REPORT_FOLDER = "team_members_statistics"
REPORT_DATA_FILE = "weekly_report_data.txt"
DATE_TIMESTAMP = None
TIME_TIMESTAMP = None

def initSettings():
    initFolder(REPORT_DATA_FOLDER)
    
    global DATE_TIMESTAMP, TIME_TIMESTAMP
    DATE_TIMESTAMP = strftime(DATES_FORMAT, gmtime())
    TIME_TIMESTAMP = strftime(TIME_FORMAT, gmtime())

    config = configparser.ConfigParser()
    config.read('settings.ini')
    print("Reading settings from properties file")
    print(config.get("JIRA Connection", "jira_login_url"))
    print(config.get("JIRA Connection", "jira_issues_search_url"))

    global JIRA_LOGIN_URL
    JIRA_LOGIN_URL = config.get("JIRA Connection", "jira_login_url")
    global JIRA_ISSUES_SEARCH_URL
    JIRA_ISSUES_SEARCH_URL = config.get("JIRA Connection", "jira_issues_search_url")
    global USER_NAME
    USER_NAME = config.get("JIRA User", "user_name")
    global PASSW
    PASSW = config.get("JIRA User", "passw")
    global START_DATE_OF_WEEK
    START_DATE_OF_WEEK = config.get("JIRA JQL params", "start_date_of_week")
    global END_DATE_OF_WEEK
    END_DATE_OF_WEEK = config.get("JIRA JQL params", "end_date_if_week")
    global SPRINT_IDS_LIST
    SPRINT_IDS_LIST = config.get("JIRA JQL params", "sprint_ids_list")
    global PROJECTS_LIST
    PROJECTS_LIST = [e.strip() for e in config.get("JIRA JQL params", "projects_list").split(';')]
    global TEAM_LIST
    TEAM_LIST = [e.strip() for e in config.get("JIRA JQL params", "team_list").split(';')]
    global TIMER_INTERVAL
    TIMER_INTERVAL = int(config.get("Script Params", "timer_interval"))
    global TIMER_INTERVAL_SHORT
    TIMER_INTERVAL_SHORT = int(config.get("Script Params", "timer_interval_short"))
    global WEBDRIVER_PATH
    WEBDRIVER_PATH = config.get("Script Params", "webdriver_path")

    print("Reading settings from properties file - DONE")
    return 1

def initFolder(folderPath):
    if (os.path.isdir(folderPath)):
        print(folderPath + " folder exists - OK")
    else:
        print(folderPath + " folder doesn't exist and will be created")
        os.makedirs(folderPath)
        print(folderPath + " folder was created - OK")

def periodForDatesInDays(startDate, endDate, datesFormat):
    createdDate = datetime.strptime(startDate, datesFormat)
    completedDate = datetime.strptime(endDate, datesFormat)
    period = completedDate - createdDate
    return period.days

def normalizeDateFormat(dateToNormalize, originalFromatDescriptor):
    objDate = datetime.strptime(dateToNormalize, originalFromatDescriptor)
    normalizedDate = objDate.strftime(DATES_FORMAT)
    return normalizedDate

def reformatPercentageValue(percentageValue):
    formattedValue = percentageValue.replace("%","")
    
    return formattedValue

def cleanHtmlInnerContent(contentToClean):
    contentToClean = contentToClean.strip()
    contentToClean = contentToClean.replace("<p>", "\n")
    contentToClean = contentToClean.replace("<br>", "\n")
    contentToClean = contentToClean.replace("</p>", "")
    contentToClean = contentToClean.replace("&amp;", "&")
    contentToClean = contentToClean.replace("<i>", "'")
    contentToClean = contentToClean.replace("</i>", "'")
    contentToClean = contentToClean.replace("<b>", "")
    contentToClean = contentToClean.replace("</b>", "")
    contentToClean = contentToClean.replace("<u>", "")
    contentToClean = contentToClean.replace("</u>", "")
    contentToClean = contentToClean.replace("\n", "")
    
    return contentToClean

def appendDataToReportFile(dataToWrite):
    reportTextFile = REPORT_DATA_FOLDER + "\\" + DATE_TIMESTAMP + "_" + REPORT_DATA_FILE
    print("Writing to file : " + reportTextFile)
    with open(reportTextFile, "a", encoding='utf-8') as text_file:
        print(str(dataToWrite), file=text_file)

    return 1

def appendDataToTeamMemberStatsFile(team_member, dataToWrite):
    memberFolderPath = REPORT_DATA_FOLDER + "\\" + TEAM_MEMBERS_REPORT_FOLDER
    initFolder(memberFolderPath)
    reportTextFile = memberFolderPath + "\\" + team_member.lower() + ".txt"
    print("Writing to file : " + reportTextFile)
    with open(reportTextFile, "a", encoding='utf-8') as text_file:
        print(str(dataToWrite), file=text_file)

    return 1
