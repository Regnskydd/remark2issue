#!/usr/bin/env python3

def die(message):
    print(message)
    import sys
    sys.exit(1)

from remark import Remark
import optparse
import json

try:
    import requests
except ImportError:
    die("Please install restkit")

try:
    from openpyxl import load_workbook
except ImportError:
    die("Please install openpyxl")

# CONSTANTS
IR_SHEET = 2
ID_COL = 1
AUTHOR_COL = 5
COMMENT_COL = 6
DECISION_COL = 9
DECISION_COMMENT_COL = 10
ACTION_DESCRIPTION_COL = 11
STATUS_COL = 12
BEARER_TOKEN = "ADD TOKEN"

def parse_args():
    parser = optparse.OptionParser()
    parser.add_option('-j', '--jira', dest='jira_url', default='http://localhost:8080', help='JIRA Base URL')
    parser.add_option('-f', '--file', dest='filename', help='Filename for excel workbook')
    parser.add_option('-k', '--key', dest='key', help='Project key to jira')

    return parser.parse_args()

def fetch_open_remarks(workbook):
    sheet = workbook[workbook.sheetnames[IR_SHEET]]
    remarks = []
    
    for row_index in range(sheet.max_row)[1:]:
        # Check if the remark have been accepted or postponed
        if sheet.cell(row_index,DECISION_COL).value == 'A' or sheet.cell(row_index,DECISION_COL).value == 'P':
            #Check if the remark is open
            if not sheet.cell(row_index,STATUS_COL).value:
                remarks.append(Remark(row_index,
                                      sheet.cell(row_index,AUTHOR_COL).value,
                                      sheet.cell(row_index,COMMENT_COL).value,
                                      sheet.cell(row_index,DECISION_COL).value,
                                      sheet.cell(row_index,DECISION_COMMENT_COL).value,
                                      sheet.cell(row_index,ACTION_DESCRIPTION_COL).value,
                                      sheet.cell(row_index,STATUS_COL).value))
    return remarks

# Fetches the status for an issue
def fetch_issue_status_from_jira(options):
    headers = {"Content-Type": "application/json", "Authorization": BEARER_TOKEN}
    issues = []
    # Only fetch status on issues that is still open
    sheet = workbook[workbook.sheetnames[IR_SHEET]]
    for row_index in range(sheet.max_row)[1:]:
        # Check if the remark have been accepted or postponed
        if sheet.cell(row_index,DECISION_COL).value == 'A' or sheet.cell(row_index,DECISION_COL).value == 'P':
            #Check if the remark is open
            if not sheet.cell(row_index,STATUS_COL).value:
                action_description = sheet.cell(row_index,ACTION_DESCRIPTION_COL).value
                # This could be empty if no issues have been created with the script previously or if the user have manually removed them. So double check to be certain. 
                if action_description != None:
                    request_url = options.jira_url + "/rest/api/latest/issue/" + action_description.split(' ', 1)[0]
                    #response = requests.get(url=request_url,headers=headers,data=json.dumps(payload))

                    issue = json.loads(response)

                    if issue
                    
                    if response.status_code == 200:
                        issues.append(json.loads(response))
                    else:
                        die(response)

def write_issue_key_on_remark(remark,options,response):
    workbook = load_workbook(options.filename)
    sheet = workbook[workbook.sheetnames[IR_SHEET]]

    sheet.cell(remark.get_identifier(),ACTION_DESCRIPTION_COL).value = json.loads(response)['key'] + " have been created, but not yet completed. This cell will be updated when task is completed."
    workbook.save(options.filename)

def create_issue(remark, options):
    #request_url = options.jira_url + "/rest/api/latest/issues/"
    #payload = {"fields": {"project": {"key": options.key},"summary": remark.get_identifier()}}
    #headers = {"Content-Type": "application/json", "Authorization": BEARER_TOKEN}
    #response = requests.post(url=request_url,headers=headers,data=json.dumps(payload))

    #if response.status_code == 200:
    #    print("TODO Add issue number to excel row")
    #else:
    #    die(response)
    #TEST DATA REMOVE ME
    response = '{"id":"39001","key":"TEST-102"}'
    write_issue_key_on_remark(remark,options,response)
if __name__ == '__main__':
    (options, args) = parse_args()

    # Exit if no file was supplied to program
    if not options.filename:
        parser.error('Filename to workbook not given')

    # Exit if no project key was supplied to program
    if not options.key:
        parser.error('Project key to JIRA not given')

    workbook = load_workbook(options.filename)
    
    remarks = fetch_open_remarks(workbook)

    for remark in remarks:
        print(remark.get_identifier(), remark.get_decision(),remark.get_comment())
        create_issue(remark,options)

    fetch_issue_status_from_jira(options)
