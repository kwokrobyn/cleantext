'''
clean.py version 3. Supports doc format input 
'''

import os
import re
import xlwt 

def findSpeakers(stream):

	# find speaker legend from file 

	speakers = [] 
	speakerNames = []

	legendIndex = stream.find("LEGEND:")
	endIndex = stream.find("\n\n",legendIndex)
	legendText = stream[legendIndex:endIndex]

	# speakers not detected, assign default 
	if legendIndex == -1:
		speakers = ["MD2:", "PT2:"]
		speakerNames = ["Physician", "Participant"]

	else: 
		legend = legendText.split()
		for s in legend:
			if "=" in s: 
				tmp = s.split("=")
				tmp = [x.strip() for x in tmp]
				tmp = [re.sub(r"\W", "", x) for x in tmp] # remove all non-acsii characters]
				speakers.append(tmp[0] + ":")
				speakerNames.append(tmp[1])

		print speakers
		print speakerNames 

	return (speakers, legendText)

def checkMatch(curr, speaker, stream, pos, insert):
	'''
	Return index of new speaker if new speaker detected 
	If current speaker stays the same, return -1 
	Also returns whether new speaker is in an insert 
	'''

	new = -1

	# create list of other speaker indexes 
	otherSpeakers= []
	for i in range(len(speaker)):
		if speaker[i] != curr:
			otherSpeakers.append(i)

	# check regular 
	if not insert:
		for i in otherSpeakers: 
			if stream[pos:pos+len(speaker[i])] == speaker[i]:
				new = i 

	# check insert 
	for i in otherSpeakers: 
		if stream[pos:pos+len(speaker[i])+1] == "[" + speaker[i]:
			new, insert = i, True 

	return (new, insert)

def parseText(speaker, stream):
	'''
	Parses text into separate speakers 
	'''
	# preliminary cleaning 
	stream = re.sub(r"[^\x00-\x7F]+", "", stream) # remove all non-acsii characters
	stream = re.sub(r"\(\([^\(\(\)\)]*\)\)", "", stream) # remove all text within (())
	stream = re.sub(r"\([^\(\)]*\)", "", stream) # remove all text within ()
	stream = re.sub(r"\{[^\(\)]*\}", "", stream) # remove all text within {}
	
	# stream = re.sub("\n", " ", stream) # remove all \n

	#print stream

	prev = -1
	curr = -1
	currText = ""
	recording, insert = False, False
	text = {}

	# init text dictionary. key is speaker code, val is list of text 
	for i in speaker: 
		text[i] = []

	for i in range(len(stream)):
		(new, insert) = checkMatch(curr, speaker, stream, i, insert)
		#print 'i', i
		if new != -1:
			print (new, insert)

		# start recording 
		if not recording and new != -1: 
			curr, recording = new, True 

		# recording, new speaker not detected 
		if recording and new == -1 and not insert: 
			currText += stream[i]

		# recording, new speaker detected 
		if recording and new != -1:
			text[speaker[curr]].append(currText)
			#print 'the i is:', stream[i:i+3]
			#print 'before switch', speaker[prev], speaker[curr]
			prev = curr 
			curr = new 
			#print 'after switch', speaker[prev], speaker[curr]
			currText = ""

		if recording and insert and stream[i] != "]":
			currText += stream[i]

		if recording and insert and stream[i] == "]":
			#print 'new, insert', new, insert 
			#print stream[i:i+30]
			#print 'curr, prev', curr, prev 
			text[speaker[curr]].append(currText)

			prev, curr = curr, prev
			currText = ""
			insert = False 

		#print 

	# insert remaining text 
	if currText != "":
		text[speaker[curr]].append(currText)
		currText = ""

	# further cleaning 
	for i in text:
		for j in range(len(text[i])):
			if text[i][j] != 0:
				if ":" in text[i][j]:
					text[i][j] = ':'.join(text[i][j].split(":")[1:])
				text[i][j] = re.sub( '\s+', " ", text[i][j]) 
				text[i][j] = re.sub( '\n', " ", text[i][j]) 
				text[i][j] = re.sub( '\r', " ", text[i][j]) 

	return text
def inputSheetConvergence(text, filename, speaker):
	'''
	Input separate conversation into excel sheet with the same name 
	'''
	# input into excel sheet 
	wb = xlwt.Workbook()
	ws = wb.add_sheet('Results')

	# wrap text 
	style = xlwt.XFStyle()
	style.alignment.wrap = 1

	# set width
	for i in range(len(max(text.values()))):
		ws.col(i).width = 256 * 30

	# write legend in first col 
	for i in range(len(text.keys())):
		ws.write(i, 0, text.keys()[i])

	# fill in content 
	count = 0 
	for i in text.keys():
		for j in range(len(text[i])):  
			ws.write(count, j+1, text[i][j], style)
		count += 1 

	wb.save('./convergence_outputs/' + filename[:filename.find('.doc')]+'.xls')

def inputSheetCollated(text, filename, speaker, collatedSheet, count):

	'''
	Input into single sheet 
	'''

	# wrap text 
	style = xlwt.XFStyle()
	style.alignment.wrap = 1 

	collatedSheet.write(count, 0, filename[:filename.find('.doc')], style)

	j = 1
	for i in text:
		collatedSheet.write(count, j, i, style)
		collatedSheet.write(count+1, j, ' '.join(text[i]), style)
		j += 1

	wb.save('collated_output.xls')


if __name__ == '__main__':

	# testing code 
	# testpath = 'A041_Clean.doc'

	# stream = os.popen("antiword -f ./samples3/%s" % testpath).read()

	# speaker = findSpeakers(stream)
	# text = parseText(speaker, stream)

	# inputSheet(text, testpath, speaker)

	# full code 

	count = 0
	wb = xlwt.Workbook()
	collatedSheet = wb.add_sheet('Results')

	for i in range(5):
		collatedSheet.col(i).width = 256 * 50

	path = './samples/'
	for filename in os.listdir(path):
		if filename.endswith(".doc"):
			stream = os.popen("antiword -f ./samples/%s" % filename).read()

			(speaker, legend) = findSpeakers(stream)
			text = parseText(speaker, stream)
			inputSheetConvergence(text, filename, speaker)
			inputSheetCollated(text, filename, speaker, collatedSheet, count)

			count += 2

