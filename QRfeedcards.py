"""
	Author: Liana Piedra
	Created: March 2015
	A program that generates feed cards for Claremont HVZ 
	with a feed code and its corresponding feeding URL. 


	For pyqrcode: https://pypi.python.org/pypi/PyQRCode
	Necessary to create the neat QR code image.

	For xlsxwriter: http://xlsxwriter.readthedocs.org/en/latest/index.html
	This website was a huge source, literally the coolest thing ever.
	Does practically anything you would want to do in an Excel sheet.
"""

import pyqrcode
import xlsxwriter
import random
import string

charSet = 'ACELNOPSTWXZ' # currently accepted letters for QR code

numCodes = int(raw_input("How many feed cards would you like to create? [Integer] "))
semester = str.capitalize(raw_input("What semester, Fall [F] or Spring [S]? "))
year = raw_input("What year? [Last two digits] ")

# Create a new Excel file and add a worksheet
workbook = xlsxwriter.Workbook('FeedCardSheets.xlsx')
worksheet = workbook.add_worksheet()

# These measurements are very specific to the paper being used,
# and will accomodate 2" x 3.5" perforated cards.
worksheet.set_margins(left=.75,right=.75,top=.5,bottom=.5)
worksheet.set_column('A:A', 31.3)
worksheet.set_column('B:B', 12.0)
worksheet.set_column('C:C', 43.3)

# Counter variables
rows = 0
cards = 1
fileNum = 1

# This giant loop will create the necessary formatting for each card
for number in range(1, numCodes+1): 
	# Create QR code images specific to each feed code
	num = str(number)
	feed = ''.join(random.choice(charSet) for i in range(5)) #range is how long the feed code is
	url = pyqrcode.create('http://www.claremonthvz.org/login/?next=/eat/'+feed, 
		error='L', version=5, mode='binary')
	url.png('hvz'+num+'.png', scale=8) # I suggest not changing the version or scale numbers

	# Set rows and columns to specific sizes for the paper
	worksheet.set_row(rows, 68) #row for name and QR code
	worksheet.set_row(rows+1, 80) #row for feed code
	rows += 2

	# Setup for image insertion
	num = str(cards)
	num2 = str(cards+1)
	
	# Format the name and QR image boxes of the card
	name_format = workbook.add_format({'align':'left', 'valign':'top', 'italic': True, 
		'bold': True, 'size':10, 'font':'Rockwell', 'border':1})
	nameCell = 'A'+num
	qrCell = 'B'+num
	worksheet.write(nameCell, 'Name:', name_format)
	worksheet.write(qrCell, '', name_format)
	worksheet.insert_image(qrCell, 'hvz'+str(fileNum)+'.png', {'x_offset': 5, 
		'y_offset': 5, 'x_scale': 0.26, 'y_scale': 0.26})

	# Insert each HVZ logo image
	logoCell = 'C'+num
	worksheet.insert_image(logoCell, 'HVZlogo.png', {'x_scale': 0.29, 
		'y_scale': 0.29, 'x_offset': 24, 'y_offset':15})

	# Merge the cells where the feed code goes
	mergeCell = 'A'+num2+':B'+num2
	code_format = workbook.add_format({'align':'center', 'valign':'vcenter', 
		'size':55, 'font':'Rockwell', 'bold': True, 'border':1})
	worksheet.merge_range(mergeCell, feed, code_format)

	#Merge the cells where the logo goes
	logoMerge = 'C'+num+':C'+num2
	info_format = workbook.add_format({'align':'center', 'valign':'bottom', 
		'size':11.5, 'font':'Rockwell', 'bold': True, 'border':1})
	worksheet.merge_range(logoMerge, "claremonthvz.org   "+semester+" '"+year
		+"  (909) 525-4551", info_format)

	cards += 2
	fileNum += 1

workbook.close()

