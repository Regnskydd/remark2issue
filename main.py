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
    die("Please install restkit (in debian: python3-requests)")

try:
	from xlrd import open_workbook,empty_cell
except ImportError:
    die("Please install xlrd (in debian: python3-xlrd)")
	
# CONSTANTS
IR_SHEET = 2
ID_COL = 0
AUTHOR_COL = 1
COMMENT_COL = 1
DECISION_COL = 1
DECISION_COMMENT_COL = 1
ACTION_DESCRIPTION_COL = 1
STATUS_COL = 1
BEARER_TOKEN = "ADD TOKEN"

def parse_args():
	parser = optparse.OptionParser()
	parser.add_option('-j', '--jira', dest='jira_url', default='http://localhost:8080', help='JIRA Base URL')
	parser.add_option('-f', '--file', dest='filename', help='Filename to write image to')
	parser.add_option('-k', '--key', dest='key', help='Project key to jira')
	
	return parser.parse_args()
	
def fetch_open_remarks(workbook):
	sheet = workbook.sheet_by_index(IR_SHEET)
	remarks = []
	for row_index in range(sheet.nrows):
		# Check if the remark have been accepted or postponed
		if sheet.cell(row_index,DECISION_COL).value == 'A' or sheet.cell(row_index,DECISION_COL).value == 'P':
			# Check if the remark is open
			if sheet.cell(row_index,STATUS_COL).value is empty_cell:
				remarks.append(Remark(sheet.cell(row_index,ID_COL).value,
				sheet.cell(row_index,AUTHOR_COL).value,
				sheet.cell(row_index,COMMENT_COL).value,
				sheet.cell(row_index,DECISION_COL).value,
				sheet.cell(row_index,DECISION_COMMENT_COL).value,
				sheet.cell(row_index,ACTION_DESCRIPTION_COL).value,
				sheet.cell(row_index,STATUS_COL).value))				
	return remarks

def create_issue(remark, options):
	request_url = options.jira_url + "/rest/api/latest/issue/"
	payload = {"fields": {"project": {"key": options.key},"summary": remark.get_identifier}}
	headers = {"Content-Type": "application/json", 'Authorization': BEARER_TOKEN}
	response = requests.post(url=request_url,auth=auth,headers=headers,data=json.dumps(payload))
		
	if response.status_code != 200:
		die(response)

if __name__ == '__main__':
	(options, args) = parse_args()
	
	# Exit if no file was supplied to program
	if not options.filename:
		parser.error('Filename not given')
		
	# Exit if no project key was supplied to program
	if not options.key:
		parser.error('Key to JIRA not given')
	
	workbook = open_workbook(options.filename)
	
	remarks = fetch_open_remarks(workbook)
	
	for remark in remarks:
		create_issue(remark,options)
	
