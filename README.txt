To run the program change your directory to the file containing the program, and type in your terminal/command line:
	
	python evaluate

When prompted, specify the file path of the excel spreadsheet containing your manual and automatically parsed comments.
If you'd like to view the parsed comments before each error is displayed, open up evaluate.py in a text editor and uncomment the region in the method, format_output(). The comments appear as unicode instead of strings to account for any foreign characters that the user may input within a comment.

Then, watch the magic happen...