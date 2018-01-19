'''
convergence.py version 1. 
'''

import os 
import xlwt # for writing to excel sheets 
import xlrd # for reading from excel sheets 
import re
from clean import findSpeakers

def createSectionDicts(filename, convergence_path):

	# find number of words in text 
	stream = os.popen("antiword -f ./samples/%s" % filename + '.doc').read()
	(speaker, legendText) = findSpeakers(stream)
	text = stream[stream.find(legendText) + len(legendText):].strip()
	wordCount = len(text.split())
	
	# read convergence data 
	testbook = xlrd.open_workbook(convergence_path + filename + '.xls')
	testsheet = testbook.sheet_by_name('Results')

	curr = 0
	val = 1
	count = 0
	sectionNumber = 3 
	sectionDicts = []
	dict1 = {}
	codes = []
	maxCount = wordCount / sectionNumber 

	# read speaker names 
	for i in range(len(speaker)):
		speakerName = testsheet.cell(i, 0).value.encode("ascii", "ignore")
		codes.append(speakerName) 

	# initialize keys for section dictionaries 
	for i in range(sectionNumber):
		tmp = {}
		for j in codes: 
			tmp[j] = ""
		sectionDicts.append(tmp) 

	# build dictionaries 
	for i in range(sectionNumber):
		while count < maxCount: 
			try:
				newText = testsheet.cell(curr, val).value.encode("ascii", "ignore")
			except Exception as e:
				break 
			sectionDicts[i][codes[curr]] += newText
			if curr == len(codes) - 1:
				val += 1 
			curr = (curr + 1) % len(codes)
			count += len(newText.split())
		count = 0 

	return sectionDicts 

if __name__ == '__main__':
	convergence_path = './convergence_outputs/'
	input_path = './samples/'
	
	for filename in os.listdir(input_path):
		if filename.endswith('.doc'):
			sectionLists = createSectionDicts(filename[:filename.find('.doc')], convergence_path)


	# testing code 
	#filename = 'A004_Clean'



	

