import requests
import ssl
import os
from pprint import pprint
import getpass
import xlsxwriter
import requests
from requests.auth import HTTPBasicAuth
import json

####################################### GLOBAL VARIABLES ############################################################

############### This section calls Atlassian cloud and returns a string JSON response with a page of JIRA issues from all JIRA projects.  The string is converts to a JSON object called jsonToPython
#auth = HTTPBasicAuth("kingstonn@hotmail.com", "73vpU61KFiplAyCka3Ut829A")
auth = HTTPBasicAuth("neil.kingston@intelematics.com", "QeZIvdCM0sdHYcdL3BV3B75D")
headers = {
   "Accept": "application/json"
}
pageIndex = 0
totalIssues = 0
counter = 0

######################################### FUNCTIONS ####################################################################

def writeIssues (issues_dict,worksheet_name, counter,worksheet):
    # The workbook object is then used to add new
    # worksheet via the add_worksheet() method.
#    print ("Entering writeIssues() worksheet_name={} counter={}".format(worksheet_name,counter))
#    if counter == 0:
#        worksheet = workbook.add_worksheet(worksheet_name)

    fields_dict = {
                    "issueid": "A",
                    "key": "B",
                    "summary": "C",
                    "projectkey": "D",
                    "projectname": "E",
                    "issuetype": "F",
                    "description": "G",
                    "resolution": "H",
                    "assignee": "I",
                    "linkup" : "J",
                    "linkupsummary": "K",
                    "linkupstatus": "L",
                    "linkupissuetype": "M"
    }
    for key, value in fields_dict.items():
        cell = "{}1".format(value)
        worksheet.write(cell, key)

    for s in issues_dict:
        themes_values = {"issueid": "",
                         "key": "",
                         "summary": "",
                         "projectkey": "",
                         "projectname": "",
                         "issuetype": "",
                         "description": "",
                         "resolution": "",
                         "assignee": "",
                         "linkup":"",
                         "linkupsummary": "",
                         "linkupstatus": "",
                         "linkupissuetype": ""
        }

            singleIssue = s
            themes_values["issueid"] = singleIssue["id"]
            themes_values["key"] = singleIssue["key"]
            themes_values["summary"] = singleIssue["fields"]["summary"]
            themes_values["projectkey"] = singleIssue["fields"]["project"]["key"]
            themes_values["projectname"] = singleIssue["fields"]["project"]["name"]
            themes_values["issuetype"] = singleIssue["fields"]["issuetype"]["name"]
            themes_values["description"] = singleIssue["fields"]["status"]["description"]
            if worksheet_name == "Stories":
                themes_values["linkup"] = singleIssue["fields"]["customfield_10014"]
                themes_values["linkupissuetype"] = "Epic"

            issuelinks = singleIssue["fields"]["issuelinks"]
            if len(singleIssue ["fields"]["issuelinks"]) > 0:
                for thisissuelink in issuelinks:
                    issuelinkstype = thisissuelink["type"]["name"] # Values can be Blocks, Relates, Cloners, Duplicate, Problem/Incident, Treatment


                    if thisissuelink["type"]["name"] == "Treatment":
                        if (worksheet_name == "Initiatives" or worksheet_name == "Epics") and "outwardIssue" in thisissuelink:
                            themes_values["linkup"] = thisissuelink["outwardIssue"]["key"]
                            themes_values["linkupsummary"] = thisissuelink["outwardIssue"]["fields"]["summary"]
                            themes_values["linkupstatus"] = thisissuelink["outwardIssue"]["fields"]["status"]["name"]
                            themes_values["linkupissuetype"] = thisissuelink["outwardIssue"]["fields"]["issuetype"]["name"]

    #               issuelinksoutwardkey = thisissuelink["outwardIssue"]["key"]
    #               issuelinksinwardkey = thisissuelink["inwardIssue"]["key"]

    #               issuelinkstype = singleIssue ["fields"]["issuelinks"]["type"]["name"]
    #               issuelinksoutwardkey = singleIssue ["fields"]["issuelinks"]["outwardIssue"]
    #               issuelinksinwardkey = singleIssue["fields"]["issuelinks"]["inwardIssue"]
    #               issuelinksoutwardkey = singleIssue ["fields"]["issuelinks"]["outwardIssue"]["key"]
    #               issuelinksinwardkey = singleIssue["fields"]["issuelinks"]["inwardIssue"]["key"]
    #                print("counter={} issuelinks issuelinkstype={} issuetype={}".format(counter, issuelinkstype,issuetype))
            else:
                issuelinks = "emptylist"

    # Handle exceptions.  We get dirty data in the assignee and resolution fields which are meant to be strings, but sometimes come through as dictionaries of crap
            if type(singleIssue["fields"]["assignee"]) == dict:
                assignee = "dictionary"
                themes_values["assignee"] = "dictionary"
            else:
                assignee = singleIssue["fields"]["assignee"]
                themes_values["assignee"] = singleIssue["fields"]["assignee"]

            if type(singleIssue["fields"]["resolution"]) == dict:
                resolution = "dictionary"
                themes_values["resolution"] = "dictionary"
            else:
                resolution = singleIssue["fields"]["resolution"]
                themes_values["resolution"] = singleIssue["fields"]["resolution"]

        counter = counter + 1
#        print("Counter={} totalIssue={}".format(counter,totalIssues))

        for key, value in fields_dict.items():
            cell = "{}{}".format(value, counter + 1)
            worksheet.write(cell, themes_values[key])

    return counter

#######################################################################################################################

# Workbook() takes one, non-optional, argument
# which is the filename that we want to create.
workbook = xlsxwriter.Workbook('JIRA.xlsx')

url_dict = {
            "Themes": "&jql=project in (CLOUD, SEC) AND issuetype = Theme ORDER BY key ASC&",
            "Initiatives": "&jql=project in (CLOUD, SEC) AND issuetype = Initiative ORDER BY key ASC&",
            "Epics": "&jql=project in (CLOUD, SEC) AND issuetype = Epic ORDER BY key ASC&",
            "Stories": "&jql=project in (CLOUD, SEC) AND issuetype = Story ORDER BY key ASC&"
            }

for key, value in url_dict.items():
    counter = 0
    totalIssues = 0
    worksheet = workbook.add_worksheet(key)
    while counter < totalIssues or counter == 0:
        url = "https://intelehub.atlassian.net/rest/api/3/search?startAt=" + str(counter) + value
        response = requests.request(
            "GET",
            url,
            headers=headers,
            auth=auth
            )
        jsonToPython = response.json()
        totalIssues = jsonToPython["total"]
        issues = jsonToPython["issues"]
        counter = writeIssues(issues,key,counter,worksheet)

 # Finally, close the Excel file
 # via the close() method.
workbook.close()

print('Complete excel file creation')
#########################################



###############crap below ignore ###############
def unpackDictionaryFields(**kwargs):
    counter = currentIndex
    for key, value in kwargs.items():
        counter = counter + 1
#        print ("Dictionary counter={}".format(counter))
        print ("Dictionary counter={}   Key={}: Value={}".format(counter, key, value))
        print (" ")

def loo(inputList):
    counter = currentIndex
    for i in inputList:
#        print ("listcounter=".format(counter))
#        print (type(inputList[counter]))
#        print ("list counter={}   {}:{}".format(counter, inputList[counter]))
#        print (" ")
        counter = counter + 1