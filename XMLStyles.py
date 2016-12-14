#!/usr/bin/python
# -*- coding: utf-8 -*-

from lxml import etree as ET
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import re
from collections import Counter
from docx.enum.text import WD_BREAK
import sys
import logging
import time

#Gets the Current Page Number
def getCurrentPageNum(a_format, pages_element):
	page_num = 1
	pg1 = list(a_format.iterancestors(pages_element))
	pg2 = pg1[0]
	for previous_pages in pg2.itersiblings(preceding=True):
		page_num += 1
	print("Current Page Num is ", page_num)
	logging.debug("Current Page Num is " + page_num)
	return page_num

#Insert the page number for each page tag on the new xml file
def insertPageNumbers(root_element, pages_element, formattings_element,num_of_pages,bottom3,top3):
	pagenum_location = checkPageNumbering(top3, bottom3, num_of_pages)
	if(pagenum_location == "TOP"):
		dict_1 = getPageNumbering(top3,num_of_pages)
		pagenum_detected = 1
	elif(pagenum_location == "BOTTOM"):
		dict_1 = getPageNumbering(bottom3,num_of_pages)
	elif (pagenum_location == "NONE"):
		dict_1 = {}
	else:
		dict_1 = {}
	counter = 1
	if (pagenum_location is "NONE"):
		return
	for a_page in root_element.iter(pages_element):
		a_page.attrib['PAGENUMBERINDEX'] = str(counter)
		a_page.attrib['PAGENUMBERSCANNED'] = str(dict_1[counter])
		counter += 1

#Insert an inline status into formatting tag to show if style in text is on the same line as previous text
def getInlineStatus(root_element, pages_element, formattings_element,num_of_pages,bottom3,top3):

	llb_previous = 0;
	for a_format in root_element.iter(formattings_element):
		llb_current = a_format.getparent().get("baseline", 99999)
		if (llb_current == llb_previous ):
			inline_1 = True
		elif (llb_current != llb_previous ):
			inline_1 = False
		llb_previous = llb_current
		a_format.attrib['INLINE'] = str(inline_1)

#Strip Digits from the beginning and end of first three lines, to check if there are headers plus page/section numbering
def stripDigits(string2a):
	string2a = str.lstrip(string2a, ' 0123456789 ')
	string2a = str.rstrip(string2a, ' 0123456789 ')
	return string2a

#Strips Non Digits from String to Search for Page Numbering
#ENHANCE BY SEARCHING FOR PAGE, PG, P. FIRST...
def stripNONDigits_3(string1a, pgNum):
	string2a = str(string1a)
	newString = ""
	pg_1 = 0
	for char_1 in string2a:
		if not(char_1.isdigit()):
			if (pg_1):
				break # Added			
		else:
			newString+=char_1;
			pg_1 = 1
	if((newString == "")or(int(newString) > (pgNum * 10))):
		return ""
	elif(int(newString) > 0):
		return newString

#Create a matrix of number of pages X 3 for ease of header comparison and checks
def create2DimArray(headerArray, numOfPages):
	pageLineMatrix = [[0 for i in range(3)] for i in range(numOfPages)]
	row = 0		
	ccount = 0
	for page_1 in range(numOfPages):
		col = 0
		for line_1 in range(3):
			pageLineMatrix[row][col] = headerArray[ccount]
			ccount += 1
			col += 1
		row += 1
	return pageLineMatrix

#Take a 2 dim array of page and 3 top/bottom lines and return a dictionary of common lines across pages	
def compare3TopBottomLines(pageLinesMatrix, pageCount):
	foundHeaderAlready = 0
	headerList = list()
	commonLines_dict = dict()
	row_1 = 0
	col_1 = 0
	headerCount_1 = 0 #To Check if First Line is Header
	headerCount2_1 = 0
	headerCount3_1 = 0		
	for ccount_1 in range(pageCount):
		if(pageLinesMatrix[row_1][col_1] == pageLinesMatrix[row_1 + 2][col_1]):
			headerCount_1 += 1
			commonLines_dict[ccount_1] = pageLinesMatrix[row_1][col_1]
		if(pageLinesMatrix[row_1][col_1 + 1] == pageLinesMatrix[row_1 + 2][col_1 + 1]):
			headerCount2_1 += 1
			commonLines_dict[ccount_1] = pageLinesMatrix[row_1][col_1 + 1]
		if(pageLinesMatrix[row_1][col_1 + 2] == pageLinesMatrix[row_1 + 2][col_1 + 2]):
			headerCount3_1 += 1
			commonLines_dict[ccount_1] = pageLinesMatrix[row_1][col_1 + 2]
		if(row_1 >= (pageCount - 3)): #Check to only compare 3 lines else array overflow
			break
		else:
			row_1 += 1
	return commonLines_dict

# Checker for presence of Headers
#In Progresss
def getHeader (header3, pages):
	header33 = list() #Aim is not to modify the list of headers
	header33 = header3
	if(len(header3) != (3*pages)):
		print("Header Extraction Incorrect ")
		logging.debug("Header Extraction Incorrect ")
		return
	if(pages < 2):	
		print("Only Single Page ")
		logging.debug("Only Single Page ")
		return
	#Strip digits at start and end of top 3 lines	
	else:
		headerIndex = 0
		for eachLine in header33:						
			header33[headerIndex] = stripDigits(eachLine)
			print("IE", header33[headerIndex])
			logging.debug("IE", header33[headerIndex])
			headerIndex += 1
		pageMatrix = create2DimArray(header33, pages)
		dictOfHeaders = compare3TopBottomLines(pageMatrix, pages)
		print("DICT OF HEADERS IS ****    ", dictOfHeaders)
		logging.debug("DICT OF HEADERS IS ****    ", dictOfHeaders)
	return dictOfHeaders

#Check for the Presence of Footers
# In Progress
def getFooter (footer4, pages):	
	if(len(footer4) != (3*pages)):
		print("Footer Extraction Incorrect ")
		logging.debug("Footer Extraction Incorrect ")
		return
	if(pages < 2):	
		print("Only Single Page ")
		logging.debug("Only Single Page ")
		return
	#Strip digits at start and end of top 3 lines	
	else:
		footerIndex = 0
		for eachLine in footer4:
			footer4[footerIndex] = stripDigits(eachLine)
			print ("THAT WAS NEW FOOTER FUNCTION ", footer4[footerIndex])
			logging.debug("THAT WAS NEW FOOTER FUNCTION " + footer4[footerIndex])
			footerIndex += 1
		pageMatrix = create2DimArray(footer4, pages)
		dictOfFooters = compare3TopBottomLines(pageMatrix, pages)
		print("DICT OF FOOTERS IS ****    ", dictOfFooters)
		logging.debug("DICT OF FOOTERS IS ****    " + dictOfFooters)
	return dictOfFooters

#Determines if Page Numbering is at Top, Bottom or None
def checkPageNumbering(topLines, bottomLines, num_of_pages):
	topLines_1 = list()
	bottomLines_1 = list()
	i = 0

	for tpline in topLines:
		topLines_1.append(stripNONDigits_3(tpline, num_of_pages))
		i += 1
	topLines_1 = list(filter(None, topLines_1))
	topLines_1 = list(set(topLines_1))	
	i = 0

	for bmline in bottomLines:
		bottomLines_1.append(stripNONDigits_3(bmline, num_of_pages))
		i += 1
	bottomLines_1 = list(filter(None, bottomLines_1))
	bottomLines_1 = list(set(bottomLines_1))


	if((len(topLines_1) <= 0) and (len(bottomLines_1) <= 0)):
		print("PAGENUMBERING TOP AND BOTTOM LINES = 0 - NO PAGE NUMBER DETECTED")
		logging.debug("PAGENUMBERING TOP AND BOTTOM LINES = 0 - NO PAGE NUMBER DETECTED")
		return "NONE"
	elif(len(topLines_1) > len(bottomLines_1)):
		return "TOP"
	elif(len(topLines_1) < len(bottomLines_1)):
		return "BOTTOM"
	else:
		return "NONE"

#This returns a Dictionary of Page Numbering in the format (a,b) where a = page index, b = parsed page number
def getPageNumbering(lines_array, num_of_pages):
	i = 0
	page_dict = dict()
	lines_1 = list()
	for lines in lines_array:
		lines_1.append(stripNONDigits_3(lines_array[i], num_of_pages))
		i += 1
	mat_lines = create2DimArray(lines_1, num_of_pages)
	i = 0
	j = 0
	for i in range(0, len(mat_lines)):
		for j in range(0, len(mat_lines[i])):
			if (mat_lines[i][j] is None):
				mat_lines[i][j] = ""

			if (mat_lines[i][j].isdigit()):
				page_dict[i + 1] = mat_lines[i][j]
				break;
			page_dict[i + 1] = ""
	return page_dict

def filterNumericStrings_3(string1a):
	string2a = str(string1a)
	if(string2a.isdigit()):
		return string2a
	else:
		return ""
		
def areDigitsUnique(digits_array):
	counter = Counter(digits_array)
	for iterator in counter.elements():
		if (iterator.isdigit() == False):
			return False
	mc = counter.most_common(1)
	print(mc)
	if (mc[0][1]) > 1:
		return False
	else:
		return True

#Extract Top 3 Lines from the XML Using Elements - formatting, 
def getTop3Line(root_element, pages_element, formattings_element, charParams_element):
	num1 = getNumberOfPages(root_element, pages_element, formattings_element) * 3
	top3linesandPage = [" "] * num1 
	numOfPagesCounted = 0
	for a_page_element in root_element.iter(pages_element):
		count3Lines_1 = 0		
		for a_format in a_page_element.iter(formattings_element):		
			if (count3Lines_1 > 2):
				break
			elif (count3Lines_1 <= 2):
				top3linesandPage[numOfPagesCounted * 3 + count3Lines_1] = buildString(a_format, charParams_element)
			count3Lines_1 = count3Lines_1 + 1			
		numOfPagesCounted += 1
	return top3linesandPage


def buildString(formattings_element, charParams_element):	
	strBuilt = ""
	bb = len(formattings_element.findall(charParams_element))	
	if ( bb < 1):		
		return formattings_element.text
	else:
		for a_charParams in formattings_element.iter(charParams_element):
			strBuilt = strBuilt + a_charParams.text
		return strBuilt

#Get Bottom three lines from each Page
def getBottom3Line(root_element, pages_element, formattings_element, charParams_element):
	page_num = 0
	num1 = getNumberOfPages(root_element, pages_element, formattings_element) * 3
	bottom3linesandPage = [" "] * num1 
	for a_page33 in root_element.iter(pages_element):
		countPageLines = 0		
		for a_format33 in a_page33.iter(formattings_element):		
			countPageLines = countPageLines + 1
		bottom3Counter = 0
		hcounter = 0
		for a_format333 in a_page33.iter(formattings_element):
			if((bottom3Counter > countPageLines - 4)and(bottom3Counter < countPageLines + 1)):
				bottom3linesandPage[page_num * 3 + hcounter] = buildString(a_format333, charParams_element)
				hcounter += 1
			bottom3Counter += 1
		page_num += 1
	return bottom3linesandPage

# Get the number of Pages on the XML file
def getNumberOfPages(root_element, pages_element, formattings_element):
	numOfPagesCounted = 0
	for a_page_element in root_element.iter(pages_element):
		numOfPagesCounted += 1
	return numOfPagesCounted

#END OF FUNCTIONS,
def main():
	







	LOG_FILENAME = 'XMLStyles.log'
	logging.basicConfig(filename=LOG_FILENAME,level=logging.DEBUG, filemode = 'w', format='%(message)s')
	logging.debug('*********************************')



	#Define XML File ParametersS
	print("Checking Input and Output Files ...")
	xml_file_name = ""
	if(sys.argv[1].rfind(".xml") == -1):
		xml_file_name = sys.argv[1] + ".xml"
	else:
		xml_file_name = sys.argv[1]
	print("Input File is ...", xml_file_name)
	logging.debug("Input File is ..." + xml_file_name)
	print("*********************************")

	tree = ET.parse(xml_file_name)
	root = tree.getroot()
	nameSpace = "{http://www.abbyy.com/FineReader_xml/FineReader10-schema-v1.xml}"
	blocks = nameSpace + "block"
	formattings = nameSpace + "formatting"
	charParams = nameSpace + "charParams"
	pars = nameSpace + "par"
	lines = nameSpace + "line" 
	pages = nameSpace + "page"
	texts = nameSpace + "text"
	
	numofpages = getNumberOfPages(root, pages, formattings)
	toplines = getTop3Line(root, pages, formattings, charParams)
	bottomlines = getBottom3Line(root, pages, formattings, charParams)

	root.attrib['NUMBEROFPAGES'] = str(numofpages)
	root.attrib['PAGENUMBERINGLOCATION'] = str(checkPageNumbering(toplines, bottomlines, numofpages))

	insertPageNumbers(root, pages, formattings, numofpages, bottomlines,toplines)
	getInlineStatus(root, pages, formattings, numofpages, bottomlines,toplines)
	
	created_file = ""
	if(sys.argv[2].rfind(".xml") == -1):
		created_file = sys.argv[2] + ".xml"
	else:
		created_file = sys.argv[2]

	print("Output File is ...", created_file)
	print("*********************************")
	logging.debug("Output File is ..." + created_file)

	print("Processing ...")

	new_xml_file = open(created_file, "wb")
	new_xml_file.write(ET.tostring(tree))
	new_xml_file.close()
	logging.debug("END OF RUN")
	logging.debug("********************************")
	logging.debug("EXECUTION DURATION IS  " + str((time.time() - start_time)))

start_time = time.time()	
main()



