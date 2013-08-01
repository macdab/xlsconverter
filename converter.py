#!/usr/bin/python

# read and write xls
import xlrd
import csv

import sys
import argparse

import re
import os

import xlsproperties

columns = ['file','team','jiratask','est_BA',
'est_DEV','est_QA','est_INT','est_REP','est_ANA','est_ASE',
'act_BA','act_DEV','act_QA','act_INT','act_REP','act_ANA','act_ASE',
'pen_BA','pen_DEV','pen_QA','pen_INT','pen_REP','pen_ANA','pen_ASE',
'uof_BA','uof_DEV','uof_QA','uof_INT','uof_REP','uof_ANA', 'uof_ASE']

#write data (dict) to csv file
def writeCSV(dict, ofile):
	writer = csv.writer(open(ofile, 'wb'))

	#write header row
	writer.writerow(columns)
	for key, value in dict.items():
		writer.writerow(value)
#end writeCSV

def readxlsx(ifile, ofile):
	workbook = xlrd.open_workbook(ifile)
	sheet = workbook.sheet_by_name(xlsproperties.SHEETNAME)
	
	print "reading file:", ifile

	mergedXlxs = {}

	#read relevant columns from the xlsx
	for i in range(sheet.nrows-2):
		row = []
		# keep the data source file
		row.append(ifile)

		# read team name and spec number (get only JIRA num)
		row.append(sheet.cell_value(i+2, 1).encode('utf-8'))
		if sheet.cell_value(i+2,2):
			row.append(sheet.cell_value(i+2, 2).encode('utf-8').split()[0])

		#read extimates, actuals, residuals, under/over flags (7 each)
		for k in range(12,40):	
 			row.append(sheet.cell_value(i+2, k))
		
		#add only if the record is not empty (jira number exists)
		if row[2]:
			mergedXlxs[row[2]] = row
		
	return mergedXlxs
#end readxlxs

def main(argv):

	#default inputs
	ifile = xlsproperties.IFILE
	idir = xlsproperties.IDIR

	#default output
	ofile = xlsproperties.OFILE

	files = []

	# parse arguments
	parser = argparse.ArgumentParser(prog='converter', description = 'Read a folder with xls(x) files and create a task data matrix in the output file.')
	parser.add_argument('-o', '--output_file')
	parser.add_argument('-d', '--directory')
	parser.add_argument('-f', '--file')
	args = parser.parse_args()

	if not args.output_file:
		sys.exit("Provide output file!")
	else:
		ofile = args.output_file
		print "using", ofile, "as output"

		if args.directory:
			idir = args.directory
			print "reading directory:", idir

			#read excel files - ~ ignores opened files
			for r,d,f in os.walk(idir):
				for file in f:
					if (file.endswith(".xlsx") or file.endswith(".xls")) and not file.startswith("~"):
						#split/group by letters/numbers
						words = re.findall(r"[A-Za-z0-9]+", file)
						files.append([os.path.join(r,file), int(words[next(i for i,v in enumerate(words) if v.lower() == 'sprint') + 1])])

			#sort asc by sprint number
			files = sorted(files, key=lambda sprint: sprint[1])
			
			completeData = {}
			
			#update the complete dataset with the latest file
			for file in files:
				completeData.update(readxlsx(file[0], ofile))

		if args.file:
			ifile = args.file
			print "reading file:", ifile
		
			completeData = readxlsx(ifile, ofile)

		#write data to output file
		writeCSV(completeData, ofile)
#end main

# execute the script
if __name__ == '__main__':
	main(sys.argv[1:])
