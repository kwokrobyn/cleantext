'''

convergence.py version 1. 

'''

import os 
import xlwt # for writing to excel sheets 
import xlrd # for reading from excel sheets 
import re
import numpy
import matplotlib.pyplot as plt
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

def calculateScores(data): 
	
	PromotionFocus = ["accomplish", "achiev", "aspire", "aspiration", "advance", "attain", "desire", "earn", "gain", "hope", "hoping", "ideal", "improve", "momentum", "obtain", "optimist", "promote", "promoti", "speed", "swift", "toward", "wish"]
	PreventionFocus = ["accura", "afraid", "anxi", "avoid", "careful", "conservative", "defen", "duty", "escape", "escaping", "evade", "fail", "fear", "loss", "obligation", "ought", "pain", "prevent", "protect", "responsible", "risk", "safe", "secur", "threat", "vigilan"]

	scoredict = {}

	for entry in data: 
		for key in entry: 

			promscore = 0 
			prevscore = 0 

			for num in range(len(PromotionFocus)):
				promscore += entry[key].lower().count(PromotionFocus[num])

			for num in range(len(PreventionFocus)):
				prevscore += entry[key].lower().count(PreventionFocus[num])

			if key not in scoredict: 
				scoredict[key] = [(promscore/22.0, prevscore/25.0)]

			else: 
				scoredict[key].append((promscore/22.0,prevscore/25.0))

	return scoredict

def initScoreStructure():

	masterScore = []

	tmp1 = {'speaker':'MD', 'wordset':'Promotion', 'fileList': [],'parts': 3, 'scores': [[],[],[]]}
	tmp2 = {'speaker':'PT', 'wordset':'Promotion', 'fileList': [],'parts': 3, 'scores': [[],[],[]]}
	tmp3 = {'speaker':'MD', 'wordset':'Prevention', 'fileList': [],'parts': 3, 'scores': [[],[],[]]}
	tmp4 = {'speaker':'PT', 'wordset':'Prevention', 'fileList': [],'parts': 3, 'scores': [[],[],[]]}

	masterScore.append(tmp1)
	masterScore.append(tmp2)
	masterScore.append(tmp3)
	masterScore.append(tmp4)

	return masterScore

def inputIntoMasterScore(masterScore, DictScores):

	for scoreSet in masterScore: 
				# add filename to set 
				scoreSet['fileList'].append(filename)
				# add MD Promotion scores
				if scoreSet['speaker'] == 'MD' and scoreSet['wordset'] == 'Promotion':	
					for x in range(len(scoreSet['scores'])):
						scoreSet['scores'][x].append(DictScores['MD:'][x][0])

				# add MD Prevention scores 
				if scoreSet['speaker'] == 'MD' and scoreSet['wordset'] == 'Prevention':	
					for x in range(len(scoreSet['scores'])):
						scoreSet['scores'][x].append(DictScores['MD:'][x][1])

				# add PT Promotion scores 
				if scoreSet['speaker'] == 'PT' and scoreSet['wordset'] == 'Promotion':	
					for x in range(len(scoreSet['scores'])):
						scoreSet['scores'][x].append(DictScores['PT:'][x][0])

				# add PT Prevention scores 
				if scoreSet['speaker'] == 'PT' and scoreSet['wordset'] == 'Prevention':	
					for x in range(len(scoreSet['scores'])):
						scoreSet['scores'][x].append(DictScores['PT:'][x][1])

def plotMeanOfScores(masterScore):

	for scoreSet in masterScore:
		meanList = []
		for x in scoreSet['scores']:
			meanList.append(sum(x)/len(x))

		for x in range(len(meanList)):
			meanList[x] = round(meanList[x], 6)
		print 'Scores for %s %s' % (scoreSet['speaker'], scoreSet['wordset'])
		print 'List of Means:', str(meanList)
		print 'SD:', round(numpy.std(meanList), 6)
		
		plt.plot([1,2,3], meanList, marker='o')
		plt.ylabel('Means for %s %s' % (scoreSet['speaker'], scoreSet['wordset']))
		plt.show()

if __name__ == '__main__':

	convergence_path = './convergence_outputs/'
	input_path = './samples/'
	masterScore = initScoreStructure()
	test = []
	for filename in os.listdir(input_path):

		if filename.endswith('.doc'):
			sectionLists = createSectionDicts(filename[:filename.find('.doc')], convergence_path)
			DictScores = calculateScores(sectionLists) #list of scores for each file

			# place calculated scores for this file into master data structure 
			inputIntoMasterScore(masterScore, DictScores)

	plotMeanOfScores(masterScore)

		# testing code 

		#filename = 'A004_Clean'