#evaluate.py
#Justin Selig, June 15, 2014
#Cornell eRulemaking Initiative

from constants import *
import xlrd

#Opens excel spreadsheet containing manually parsed comments
MANUAL_FILE = raw_input("Enter Manual-Parsed File: ").replace('"', '') #quotes must be removed to read file
MANUAL_WORKBOOK = xlrd.open_workbook(MANUAL_FILE)
MANUAL_SHEET = MANUAL_WORKBOOK.sheet_by_index(SHEET1)

#Opens excel spreadsheet containing auto-parsed comments
AUTO_FILE = raw_input("Enter Auto-Parsed File: ").replace('"', '')
AUTO_WORKBOOK = xlrd.open_workbook(AUTO_FILE)
AUTO_SHEET = AUTO_WORKBOOK.sheet_by_index(SHEET1)


"""Method returns a list of starting positions (rows indexed at 0) 
for every comment.
Precondition: sheet_type is MANUAL_SHEET or AUTO_SHEET."""
def find_comments(sheet_type):
	counter = 0
	list = []
	for row in range(sheet_type.nrows):	#rows and cols indexed starting at 0
		datum = sheet_type.cell_value(row, COL1)
		try:
			datum = int(datum)
			counter += 1
			list.append(row)
		except:
			continue
	return list
	
"""Returns a tuple containing start and end rows of a comment
Precondition: start is a row number (int), sheet_type is MANUAL_SHEET
or AUTO_SHEET"""
def get_one_comment(start, sheet_type):
	row = start + 1
	found = False
	while not found:
		datum = sheet_type.cell_value(row, COL1)
		try:	#if number, then next comment after current one was found
			datum = int(datum)
			found = True #should break out of while
		except:	
			row += 1
			if row == sheet_type.nrows:	#at end of file
				break
	comment_bounds = (start, row)
	return comment_bounds
	
"""Returns a list of comments as tuples given a sheet type
Precondition: sheet_type is MANUAL_SHEET or AUTO_SHEET"""
def list_comments(sheet_type):
	list = []
	for row in find_comments(sheet_type):
		list.append(get_one_comment(row, sheet_type))
	return list
		
########################### COMMENTS FOUND ####################################

"""Handles comparison of excel sheets by iterating over comments (tuples).
Precondition: This method assumes that the comments being compared 
come in precisely the same order between spreadsheets."""
def main():
	autoComments = list_comments(AUTO_SHEET)
	manualComments = list_comments(MANUAL_SHEET)
	print "Comments found, performing comparison...\n"
	commentNum = 0
	for comment in list_comments(AUTO_SHEET):	#either sheet is acceptable
		verify(autoComments[commentNum], manualComments[commentNum])
		commentNum += 1

"""Method converts auto and manual comments to lists and equates them
to ensure that identical comments won't be compared; (This reduces error).
Precondition: autoComment and manualComment are tuples delineating the
start and end of an auto and manual comment."""
def verify(autoComment, manualComment):
	autoStart = autoComment[0]+1 #avoids comment number at index 0
	autoEnd = autoComment[1]
	manualStart = manualComment[0]+1
	manualEnd = manualComment[1]
	autoList = []
	manList = []
	for autoRow in range(autoStart, autoEnd):
		autoDatum = AUTO_SHEET.cell_value(autoRow, COL1).replace(u'\xa0', ' ')
		autoList.append(autoDatum)
	for manualRow in range(manualStart, manualEnd):
		manualDatum = MANUAL_SHEET.cell_value(manualRow, COL1).replace(u'\xa0', ' ')
		manList.append(manualDatum)
	if autoList != manList:
		compare(autoComment, manualComment)
		
"""Method performs a line-by-line comparison between comments.
Precondition: autoComment and manualComment are tuples delineating the
start and end of an auto and manual comment."""
def compare(autoComment, manualComment):
	autoStart = autoComment[0]+1
	autoEnd = autoComment[1]
	manualStart = manualComment[0]+1
	manualEnd = manualComment[1]
	errors = []
	for autoRow in range(autoStart, autoEnd):
		autoDatum = AUTO_SHEET.cell_value(autoRow, COL1).replace(u'\xa0', ' ')	#string containing row info
		for manualRow in range(manualStart, manualEnd):
			manualDatum = MANUAL_SHEET.cell_value(manualRow, COL1).replace(u'\xa0', ' ')
			#this is where the magic happens
			if (autoDatum in manualDatum) or (manualDatum in autoDatum):
				#string comparisons
				if autoDatum == manualDatum:
					break
				elif len(autoDatum) > len(manualDatum):
					error = "Comment " + `AUTO_SHEET.cell_value(autoComment[0], COL1)` + ", Row " + `autoRow+1` + ": missing slice, " + `manualDatum`
					errors.append(error)
				elif len(autoDatum) < len(manualDatum):
					error = "Comment " + `AUTO_SHEET.cell_value(autoComment[0], COL1)` + ", Row " + `autoRow+1` + ": extra slice" #row is spreadsheet row
					errors.append(error)
					break
			else:
				continue
	if errors:
		format_output(autoComment, manualComment, errors)

"""Outputs data (optional) and error message.
Precondition: autoComment and manualComment are tuples, and errors
is a list of strings representing error messages."""
def format_output(autoComment, manualComment, errors):
	autoStart = autoComment[0]+1
	autoEnd = autoComment[1]
	manualStart = manualComment[0]+1
	manualEnd = manualComment[1]
	
	######## Uncomment region to view slices ##################################
	#print "Manually sliced comment " + `MANUAL_SHEET.cell_value(manualComment[0], COL1)` + ":"
	#for manualRow in range(manualStart, manualEnd):
	#	print "Comment: " + `MANUAL_SHEET.cell_value(manualComment[0], COL1)` + ", Row " + `manualRow+1` + ": " + `MANUAL_SHEET.cell_value(manualRow, COL1)`
	#print "Auto sliced comment " + `AUTO_SHEET.cell_value(autoComment[0], COL1)` + ":"
	#for autoRow in range(autoStart, autoEnd):
	#	print "Comment: " + `AUTO_SHEET.cell_value(autoComment[0], COL1)` + ", Row " + `autoRow+1` + ": " + `AUTO_SHEET.cell_value(autoRow, COL1)`
	#print "Errors to be recorded:" 
	###########################################################################
	
	for error in errors:
		print error
	print "\n"