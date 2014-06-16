#constants.py
#Justin Selig, June 15, 2014
#Cornell eRulemaking Initiative

"""xlrd requies that you specify a sheet number. I have added two types of
sheets here, but only SHEET1 is ever used. Since the sheets are indexed at 0,
change SHEET1 to 0 if spreadsheets being compared come first."""

MACRO1 = 0
SHEET1 = 1

"""This column number is used by cell_value and remains the same because 
of the format of the excel sheets given."""
COL1 = 1

"""Possible sheet types include: MANUAL_SHEET, AUTO_SHEET. These are given as
raw input by the user and are instantiated in evaluate.py"""