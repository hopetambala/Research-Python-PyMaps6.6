#!/usr/bin/python

from xlwings import *
#version 6.6

def callALL(a):
	#a is the Sheet |  format would be ~ 'Sheet1',
	#useless now
	#below is what used to be right below wb in MapPy.py (in case weird stuff happens)
	'''
	p = Sheet(1).name
	o = Sheet(2).name

	x = callALL(p)
	y = callALL(o)
	'''
	x = Range(a,'A5:A16').value + Range(a,'B5:B16').value + Range(a,'C5:C16').value + Range(a,'D5:D16').value + Range(a,'E5:E16').value + Range(a,'G5:G16').value + Range(a,'H5:H16').value  + Range(a,'I5:I16').value+ Range(a,'J5:J16').value + Range(a,'K5:K16').value + Range(a,'L5:L16').value + Range(a,'M5:M16').value+ Range(a,'N5:N16').value + Range(a,'O5:O16').value + Range(a,'P5:P16').value + Range(a,'Q5:Q16').value + Range(a,'R5:R16').value + Range(a,'T5:T16').value + Range(a,'U5:U16').value + Range(a,'V5:V16').value + Range(a,'W5:W16').value + Range(a,'X5:X16').value
 	return x;

def conditionsTable():
	Range('Y2').value = 'Male'
	Range('Y3').value = 'Female'
	Range('Z1').value = 'Groupwork'
	Range('AA1').value = 'Non-Groupwork'
	Range('AB1').value = 'Percent Similarity Between Group Sheets'

	Range('Y5').value = 'Conditions'
	Range('Y6').value = 'Condition 1'
	Range('Y7').value = 'Condition 2'
	Range('Y8').value = 'Condition 3'
	Range('Y9').value = 'Condition 4'

	Range('Z5').value = 'Sum'
	Range('Z6').value = '=SUMIF(A18:X28,1)'
	Range('Z7').value = '=SUMIF(A18:X28,2)/2'
	Range('Z8').value = '=SUMIF(A18:X28,3)/3'
	Range('Z9').value = '=SUMIF(A18:X28,4)/4'

	Range('AA5').value = 'Percentage'
	Range('AA6').value = str((Range('Z6').value / (Range('Z6').value + Range('Z7').value + Range('Z8').value + Range('Z9').value) *100)) + '%'
	Range('AA7').value = str((Range('Z7').value	/ (Range('Z6').value + Range('Z7').value + Range('Z8').value + Range('Z9').value) *100)) + '%'
	Range('AA8').value = str((Range('Z8').value	/ (Range('Z6').value + Range('Z7').value + Range('Z8').value + Range('Z9').value) *100)) + '%'
	Range('AA9').value = str((Range('Z9').value	/ (Range('Z6').value + Range('Z7').value + Range('Z8').value + Range('Z9').value) *100)) + '%'

	Range('AB5').value = 'Description'
	Range('AB6').value = 'Groupwork (no isolated)'
	Range('AB7').value = 'Groupwork (isolated on left/right)'
	Range('AB8').value = 'No Groupwork (no isolation)'
	Range('AB9').value = 'No Groupwork (isolated on left/right)'

def header(noc,workbook):
	if Range('AE1').value is None:
			Range('AE1').value = "_"
	if Range('AE2').value is None:
			Range('AE2').value = "_"
	if Range('AE3').value is None:
			Range('AE3').value = "_"
	if Range('AE4').value is None:
			Range('AE4').value = "_"
	if Range('AE5').value is None:
			Range('AE5').value = "_"

	Range('AE1').value = Range(noc,'D1',wkb=workbook).value
	Range('AE2').value = Range(noc,'D2',wkb=workbook).value

	Range('AE3').value = Range(noc,'D3',wkb=workbook).value

	Range('AE4').value = Range(noc,'Q1',wkb=workbook).value
	Range('AE5').value = Range(noc,'Q2',wkb=workbook).value

	Range('AD1').value = "Date"
	Range('AD2').value = "Class"
	Range('AD3').value = "Section/Time"
	Range('AD4').value = "Instructor"
	Range('AD5').value = "Location"

def continuityTable():
	Range('Y18').value = 'OBSERVER ONE'
	Range('Y19').value = 'Size of Group'
	Range('Y20').value = '1'
	Range('Y21').value = '2'
	Range('Y22').value = '3'
	Range('Y23').value = '4'
	Range('Y24').value = '5'

	Range('Y26').value = 'OBSERVER TWO'
	Range('Y27').value = 'Size of Group'
	Range('Y28').value = '1'
	Range('Y29').value = '2'
	Range('Y30').value = '3'
	Range('Y31').value = '4'
	Range('Y32').value = '5'

	Range('Y13').value = 'OBSERVER ONE'
	Range('Y14').value = "OBSERVER TWO"
	Range('Z12').value = 'Groups All Male'
	Range('AA12').value = 'Groups All Female'
	Range('AB12').value = 'Groups Mixed'

	Range('Z19').value = 'Sum(Continuous)'
	Range('Z20').value = '=-SUMIF(A5:X16,-1)'
	Range('Z21').value = ''
	Range('Z22').value = ''
	Range('Z23').value = ''
	Range('Z24').value = ''

	Range('Z27').value = 'Sum(Continuous)'
	Range('Z28').value = '=-SUMIF(A5:X16,-1)'
	Range('Z29').value = ''
	Range('Z30').value = ''
	Range('Z31').value = ''
	Range('Z32').value = ''


	Range('AA19').value = 'Sum(Discontinuous)'
	Range('AA20').value = ''
	Range('AA21').value = ''
	Range('AA22').value = ''
	Range('AA23').value = ''
	Range('AA24').value = ''

	Range('AA27').value = 'Sum(Discontinuous)'
	Range('AA32').value = ''
	Range('AA28').value = ''
	Range('AA29').value = ''
	Range('AA30').value = ''
	Range('AA31').value = ''

	Range('AB19').value = 'Percent Continuous'
	Range('AB20').value = '=SUM(Z20)/SUM(Z20:AA20)'

	Range('AB27').value	= 'Percent Continuous'
	Range('AB28').value = '=SUM(Z20/SUM(Z20:AA20)'
	
	Range('AC19').value = 'Percent Discontinuous'
	Range('AC20').value = '=SUM(AA20)/SUM(Z20:AA20)'

	Range('AC27').value = 'Percent Discontinuous'
	Range('AC28').value = '=SUM(AA20)/SUM(Z20:AA20)'

def generateLeftrow(noc):
	
	hehe = []
	for y in range(5,16):
		hehe.append('=IF(AND(' + noc +'!' + 'A' + str(y) + '=-1,ISBLANK(' + noc +'!' + 'B' + str(y) + ')),4,IF(AND(' + noc +'!' + 'A' + str(y) + '=1,ISBLANK(' + noc +'!' + 'B' + str(y) + ')),2,IF(AND(' + noc +'!' + 'A' + str(y) + '=-1, OR(ISNUMBER(SEARCH(1,' + 'B' + str(y) + ')),)),3,IF(AND(' + noc +'!' + 'A' + str(y) + '=1, OR(ISNUMBER(SEARCH(1,' + 'B' + str(y) + ')),)),1,0))))') #the returns of your own function

	i=0
	for x in range(18,29): #loop through a range
			Range('A' + str(x)).value = hehe[i] #the returns of your own function
			i += 1

def generateMap(B,C,A,noc):
	#noc is "Name of Sheet" but I'm an idiot and used a "c" instead of an "s" for the acronym 
	B=B
	C=C
	A=A
	
	hehe = []
	for y in range(5,16):
		#hehe.append('=IF(AND(Group!' + 'A' + str(y) + '=-1,ISBLANK(Group!' + 'B' + str(y) + ')),4,IF(AND(Group!' + 'A' + str(y) + '=1,ISBLANK(Group!' + 'B' + str(y) + ')),2,IF(AND(Group!' + 'A' + str(y) + '=-1, OR(ISNUMBER(SEARCH(1,' + 'A' + str(y) + ')),)),3,IF(AND(Group!' + 'A' + str(y) + '=1, OR(ISNUMBER(SEARCH(1,' + 'B' + str(y) + ')),)),1,0))))') #the returns of your own function
		hehe.append('=IF(AND(' + noc +'!' + str(B) + str(y) + '=-1, ISBLANK(' + noc +'!' + str(C) +  str(y) + '),ISBLANK(' + noc +'!'+ str(A) + str(y) + ')),4,IF(AND(' + noc +'!' + str(B) + str(y) + '=1, ISBLANK(' + noc +'!' + str(C) + str(y) + '),ISBLANK(' + noc +'!' + str(A) + str(y) + ')),2,IF(AND(' + noc +'!' + str(B) + str(y) + '=-1, OR(ISNUMBER(SEARCH(1,'+str(A) + str(y) + ')), ISNUMBER(SEARCH(1,'+ str(C) + str(y) + ')))),3,IF(AND(' + noc +'!' + str(B) + str(y) + '=1, OR(ISNUMBER(SEARCH(1,' + str(A) + str(y) + ')), ISNUMBER(SEARCH(1,' + str(C) + str(y) + ')))),1,0))))')
	i=0
	for x in range(18,29): #loop through a range
			Range(str(B) + str(x)).value = hehe[i] #the returns of your own function
			i += 1

def generateRightrow(noc):
	
	hehe = []
	for y in range(5,16):
		hehe.append('=IF(AND(' + noc +'!' + 'X' + str(y) + '=-1,ISBLANK(' + noc +'!' + 'W' + str(y) + ')),4,IF(AND(' + noc +'!' + 'X' + str(y) + '=1,ISBLANK(' + noc +'!' + 'W' + str(y) + ')),2,IF(AND(' + noc +'!' + 'X' + str(y) + '=-1, OR(ISNUMBER(SEARCH(1,' + 'W' + str(y) + ')),)),3,IF(AND(' + noc +'!' + 'X' + str(y) + '=1, OR(ISNUMBER(SEARCH(1,' + 'X' + str(y) + ')),)),1,0))))') #the returns of your own function

	i=0
	for x in range(18,29): #loop through a range
			Range('X' + str(x)).value = hehe[i] #the returns of your own function
			i += 1

def genderate(noc,noc1,noc2,workbook,workbook1,workbook2):
	#combines gender sheets of both observers and removes disparity in the final generated sheet
	#noc is  "Gender" of Observer 1
	#noc1 is "Gender" of Observer 2
	#noc2 is "Analysis"

	#workbook is Observer 1
	#workbook1 is Observer 2
	#workbook2 is Individual Analysis


	haha = ['A', 'B', 'C', 'D', 'E','G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'T', 'U', 'V', 'W', 'X']

	for i in range(0,22):
		for y in range(5,16):
			if Range(noc, haha[i] + str(y),wkb=workbook).value == -1 or Range(noc1, haha[i] + str(y),wkb=workbook1).value == -1:
				Range(noc2, haha[i] + str(y+26),wkb=workbook2).value = -1

	for i in range(0,22):
		for y in range(5,16):
			if Range(noc, haha[i] + str(y),wkb=workbook).value == 1:
				Range(noc2, haha[i] + str(y+26),wkb=workbook2).value = 1

	for i in range(0,22):
		for y in range(5,16):
			if Range(noc1, haha[i] + str(y),wkb=workbook1).value == 1:
				Range(noc2, haha[i] + str(y+26),wkb=workbook2).value = 1

	for i in range(0,22):
		for y in range(5,16):
			if Range(noc, haha[i] + str(y),wkb=workbook).value == 1 or Range(noc1, haha[i] + str(y),wkb=workbook1).value == 1:
				Range(noc2, haha[i] + str(y+26),wkb=workbook2).value = 1

def conditionsGenderate():
	#conditions for gender
	haha = ['A', 'B', 'C', 'D', 'E','G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'T', 'U', 'V', 'W', 'X']
	'''
	Conditions
	1 - Isolated
	2 - Sitting next to one female
	3 - Sitting next to one male
	4 - Sitting next to a male and female
	5 - Sitting next to all females
	6 - Sitting next to all males
	'''
	
	#A42 to A53
	#MainRows
	for i in range(1,4):
		for y in range(31,42):
			Range(haha[i] + str(y+12)).value = '=IF(AND(' + haha[i] + str(y) + '=-1,ISBLANK(' +haha[i+1] + str(y) +'),ISBLANK(' + haha[i-1] + str(y)+ ')),-1,IF(AND(' + haha[i] + str(y) + '=-1,OR(AND(ISBLANK(' + haha[i-1] + str(y)+ '),' +haha[i+1] + str(y) +'=-1),AND(ISBLANK(' +haha[i+1] + str(y) +'),' + haha[i-1] + str(y)+ '=-1))),-2,IF(AND(' + haha[i] + str(y) + '=-1,OR(AND(ISBLANK(' + haha[i-1] + str(y)+ '),' +haha[i+1] + str(y) +'=1),AND(ISBLANK(' +haha[i+1] + str(y) +'),' + haha[i-1] + str(y)+ '=1))),-3,IF(AND(' + haha[i] + str(y) + '=-1,OR(AND(' +haha[i+1] + str(y) +'=-1,' + haha[i-1] + str(y)+ '=1),AND(' +haha[i+1] + str(y) +'=1,' + haha[i-1] + str(y)+ '=-1))),-4,IF(AND(' + haha[i] + str(y) + '=-1,' +haha[i+1] + str(y) +'=-1,' + haha[i-1] + str(y)+ '=-1),-5,IF(AND(' + haha[i] + str(y) + '=-1,' +haha[i+1] + str(y) +'=1,' + haha[i-1] + str(y)+ '=1),-6,IF(AND(' + haha[i] + str(y) + '=1,ISBLANK(' +haha[i+1] + str(y) +'),ISBLANK(' + haha[i-1] + str(y)+ ')),1,IF(AND(' + haha[i] + str(y) + '=1,OR(AND(ISBLANK(' + haha[i-1] + str(y)+ '),' +haha[i+1] + str(y) +'=-1),AND(ISBLANK(' +haha[i+1] + str(y) +'),' + haha[i-1] + str(y)+ '=-1))),2,IF(AND(' + haha[i] + str(y) + '=1,OR(AND(ISBLANK(' + haha[i-1] + str(y)+ '),' +haha[i+1] + str(y) +'=1),AND(ISBLANK(' +haha[i+1] + str(y) +'),' + haha[i-1] + str(y)+ '=1))),3,IF(AND(' + haha[i] + str(y) + '=1,OR(AND(' +haha[i+1] + str(y) +'=-1,' + haha[i-1] + str(y)+ '=1),AND(' +haha[i+1] + str(y) +'=1,' + haha[i-1] + str(y)+ '=-1))),4,IF(AND(' + haha[i] + str(y) + '=1,' +haha[i+1] + str(y) +'=-1,' + haha[i-1] + str(y)+ '=-1),5,IF(AND(' + haha[i] + str(y) + '=1,' +haha[i+1] + str(y) +'=1,' + haha[i-1] + str(y)+ '=1),6,0))))))))))))'

	for i in range(6,16):
		for y in range(31,42):
			Range(haha[i] + str(y+12)).value = '=IF(AND(' + haha[i] + str(y) + '=-1,ISBLANK(' +haha[i+1] + str(y) +'),ISBLANK(' + haha[i-1] + str(y)+ ')),-1,IF(AND(' + haha[i] + str(y) + '=-1,OR(AND(ISBLANK(' + haha[i-1] + str(y)+ '),' +haha[i+1] + str(y) +'=-1),AND(ISBLANK(' +haha[i+1] + str(y) +'),' + haha[i-1] + str(y)+ '=-1))),-2,IF(AND(' + haha[i] + str(y) + '=-1,OR(AND(ISBLANK(' + haha[i-1] + str(y)+ '),' +haha[i+1] + str(y) +'=1),AND(ISBLANK(' +haha[i+1] + str(y) +'),' + haha[i-1] + str(y)+ '=1))),-3,IF(AND(' + haha[i] + str(y) + '=-1,OR(AND(' +haha[i+1] + str(y) +'=-1,' + haha[i-1] + str(y)+ '=1),AND(' +haha[i+1] + str(y) +'=1,' + haha[i-1] + str(y)+ '=-1))),-4,IF(AND(' + haha[i] + str(y) + '=-1,' +haha[i+1] + str(y) +'=-1,' + haha[i-1] + str(y)+ '=-1),-5,IF(AND(' + haha[i] + str(y) + '=-1,' +haha[i+1] + str(y) +'=1,' + haha[i-1] + str(y)+ '=1),-6,IF(AND(' + haha[i] + str(y) + '=1,ISBLANK(' +haha[i+1] + str(y) +'),ISBLANK(' + haha[i-1] + str(y)+ ')),1,IF(AND(' + haha[i] + str(y) + '=1,OR(AND(ISBLANK(' + haha[i-1] + str(y)+ '),' +haha[i+1] + str(y) +'=-1),AND(ISBLANK(' +haha[i+1] + str(y) +'),' + haha[i-1] + str(y)+ '=-1))),2,IF(AND(' + haha[i] + str(y) + '=1,OR(AND(ISBLANK(' + haha[i-1] + str(y)+ '),' +haha[i+1] + str(y) +'=1),AND(ISBLANK(' +haha[i+1] + str(y) +'),' + haha[i-1] + str(y)+ '=1))),3,IF(AND(' + haha[i] + str(y) + '=1,OR(AND(' +haha[i+1] + str(y) +'=-1,' + haha[i-1] + str(y)+ '=1),AND(' +haha[i+1] + str(y) +'=1,' + haha[i-1] + str(y)+ '=-1))),4,IF(AND(' + haha[i] + str(y) + '=1,' +haha[i+1] + str(y) +'=-1,' + haha[i-1] + str(y)+ '=-1),5,IF(AND(' + haha[i] + str(y) + '=1,' +haha[i+1] + str(y) +'=1,' + haha[i-1] + str(y)+ '=1),6,0))))))))))))'

	for i in range(18,21):
		for y in range(31,42):
			Range(haha[i] + str(y+12)).value = '=IF(AND(' + haha[i] + str(y) + '=-1,ISBLANK(' +haha[i+1] + str(y) +'),ISBLANK(' + haha[i-1] + str(y)+ ')),-1,IF(AND(' + haha[i] + str(y) + '=-1,OR(AND(ISBLANK(' + haha[i-1] + str(y)+ '),' +haha[i+1] + str(y) +'=-1),AND(ISBLANK(' +haha[i+1] + str(y) +'),' + haha[i-1] + str(y)+ '=-1))),-2,IF(AND(' + haha[i] + str(y) + '=-1,OR(AND(ISBLANK(' + haha[i-1] + str(y)+ '),' +haha[i+1] + str(y) +'=1),AND(ISBLANK(' +haha[i+1] + str(y) +'),' + haha[i-1] + str(y)+ '=1))),-3,IF(AND(' + haha[i] + str(y) + '=-1,OR(AND(' +haha[i+1] + str(y) +'=-1,' + haha[i-1] + str(y)+ '=1),AND(' +haha[i+1] + str(y) +'=1,' + haha[i-1] + str(y)+ '=-1))),-4,IF(AND(' + haha[i] + str(y) + '=-1,' +haha[i+1] + str(y) +'=-1,' + haha[i-1] + str(y)+ '=-1),-5,IF(AND(' + haha[i] + str(y) + '=-1,' +haha[i+1] + str(y) +'=1,' + haha[i-1] + str(y)+ '=1),-6,IF(AND(' + haha[i] + str(y) + '=1,ISBLANK(' +haha[i+1] + str(y) +'),ISBLANK(' + haha[i-1] + str(y)+ ')),1,IF(AND(' + haha[i] + str(y) + '=1,OR(AND(ISBLANK(' + haha[i-1] + str(y)+ '),' +haha[i+1] + str(y) +'=-1),AND(ISBLANK(' +haha[i+1] + str(y) +'),' + haha[i-1] + str(y)+ '=-1))),2,IF(AND(' + haha[i] + str(y) + '=1,OR(AND(ISBLANK(' + haha[i-1] + str(y)+ '),' +haha[i+1] + str(y) +'=1),AND(ISBLANK(' +haha[i+1] + str(y) +'),' + haha[i-1] + str(y)+ '=1))),3,IF(AND(' + haha[i] + str(y) + '=1,OR(AND(' +haha[i+1] + str(y) +'=-1,' + haha[i-1] + str(y)+ '=1),AND(' +haha[i+1] + str(y) +'=1,' + haha[i-1] + str(y)+ '=-1))),4,IF(AND(' + haha[i] + str(y) + '=1,' +haha[i+1] + str(y) +'=-1,' + haha[i-1] + str(y)+ '=-1),5,IF(AND(' + haha[i] + str(y) + '=1,' +haha[i+1] + str(y) +'=1,' + haha[i-1] + str(y)+ '=1),6,0))))))))))))'

	#LeftSideofRows
	for y in range(31,42):
			Range('A' + str(y+12)).value = 	'=IF(AND(A' +str(y) + '=-1,ISBLANK(B'+ str(y)+')),-1,IF(AND(A' +str(y) + '=-1,B'+ str(y)+'=-1),-5,IF(AND(A' +str(y) + '=-1,B'+ str(y)+'=1),-6,IF(AND(A' +str(y) + '=1,ISBLANK(B'+ str(y)+')),1,IF(AND(A' +str(y) + '=1,B'+ str(y)+'=-1),5,IF(AND(A' +str(y) + '=1,B'+ str(y)+'=1),6,0))))))'
	for y in range(31,42):
			Range('G' + str(y+12)).value =	'=IF(AND(G' +str(y) + '=-1,ISBLANK(H'+ str(y)+')),-1,IF(AND(G' +str(y) + '=-1,H'+ str(y)+'=-1),-5,IF(AND(G' +str(y) + '=-1,H'+ str(y)+'=1),-6,IF(AND(G' +str(y) + '=1,ISBLANK(H'+ str(y)+')),1,IF(AND(G' +str(y) + '=1,H'+ str(y)+'=-1),5,IF(AND(G' +str(y) + '=1,H'+ str(y)+'=1),6,0))))))'
	for y in range(31,42):
			Range('T' + str(y+12)).value = 	'=IF(AND(T' +str(y) + '=-1,ISBLANK(U'+ str(y)+')),-1,IF(AND(T' +str(y) + '=-1,U'+ str(y)+'=-1),-5,IF(AND(T' +str(y) + '=-1,U'+ str(y)+'=1),-6,IF(AND(T' +str(y) + '=1,ISBLANK(U'+ str(y)+')),1,IF(AND(T' +str(y) + '=1,U'+ str(y)+'=-1),5,IF(AND(T' +str(y) + '=1,U'+ str(y)+'=1),6,0))))))'

	#RightSideofRows
	for y in range(31,42):
			Range('E' + str(y+12)).value = '=IF(AND(E' +str(y) + '=-1, ISBLANK(D' +str(y) + ')),-1,IF(AND(E' +str(y) + '=-1, D' +str(y) + '=-1),-5,IF(AND(E' +str(y) + '=-1, D' +str(y) + '=1),-6,IF(AND(E' +str(y) + '=1, ISBLANK(D' +str(y) + ')),1,IF(AND(E' +str(y) + '=1,D' +str(y) + '=-1),5,IF(AND(E' +str(y) + '=1,D' +str(y) + '=1),6,0))))))'
	for y in range(31,42):
			Range('R' + str(y+12)).value = '=IF(AND(R' +str(y) + '=-1, ISBLANK(S' +str(y) + ')),-1,IF(AND(R' +str(y) + '=-1, S' +str(y) + '=-1),-5,IF(AND(R' +str(y) + '=-1, S' +str(y) + '=1),-6,IF(AND(R' +str(y) + '=1, ISBLANK(S' +str(y) + ')),1,IF(AND(R' +str(y) + '=1,S' +str(y) + '=-1),5,IF(AND(R' +str(y) + '=1,S' +str(y) + '=1),6,0))))))'
	for y in range(31,42):
			Range('X' + str(y+12)).value = '=IF(AND(X' +str(y) + '=-1, ISBLANK(W' +str(y) + ')),-1,IF(AND(X' +str(y) + '=-1, W' +str(y) + '=-1),-5,IF(AND(X' +str(y) + '=-1, W' +str(y) + '=1),-6,IF(AND(X' +str(y) + '=1, ISBLANK(W' +str(y) + ')),1,IF(AND(X' +str(y) + '=1,W' +str(y) + '=-1),5,IF(AND(X' +str(y) + '=1,W' +str(y) + '=1),6,0))))))'

	######################################
	Range('Y42').value = 'Male'
	Range('Y43').value = '=COUNTIF(A43:X54,1)'
	Range('Y44').value = '=COUNTIF(A43:X54,2)'
	Range('Y45').value = '=COUNTIF(A43:X54,3)'
	Range('Y46').value = '=COUNTIF(A43:X54,4)'
	Range('Y47').value = '=COUNTIF(A43:X54,5)'
	Range('Y48').value = '=COUNTIF(A43:X54,6)'

	Range('Z42').value = 'Male Percentage'
	Range('Z43').value = '=Y43/(SUM(Y43:Y48))'
	Range('Z44').value = '=Y44/(SUM(Y43:Y48))'
	Range('Z45').value = '=Y45/(SUM(Y43:Y48))'
	Range('Z46').value = '=Y46/(SUM(Y43:Y48))'
	Range('Z47').value = '=Y47/(SUM(Y43:Y48))'
	Range('Z48').value = '=Y48/(SUM(Y43:Y48))'

	Range('AA42').value = 'Female'
	Range('AA43').value = '=COUNTIF(A43:X54,-1)'
	Range('AA44').value = '=COUNTIF(A43:X54,-2)'
	Range('AA45').value = '=COUNTIF(A43:X54,-3)'
	Range('AA46').value = '=COUNTIF(A43:X54,-4)'
	Range('AA47').value = '=COUNTIF(A43:X54,-5)'
	Range('AA48').value = '=COUNTIF(A43:X54,-6)'

	Range('AB42').value = 'Female Percentage'
	Range('AB43').value = '=AA43/(SUM(AA43:AA48))'
	Range('AB44').value = '=AA44/(SUM(AA43:AA48))'
	Range('AB45').value = '=AA45/(SUM(AA43:AA48))'
	Range('AB46').value = '=AA46/(SUM(AA43:AA48))'
	Range('AB47').value = '=AA47/(SUM(AA43:AA48))'
	Range('AB48').value = '=AA48/(SUM(AA43:AA48))'

	Range('AC42').value = 'Condition'
	Range('AC43').value = 1
	Range('AC44').value = 2
	Range('AC45').value = 3
	Range('AC46').value = 4
	Range('AC47').value = 5
	Range('AC48').value = 6

	Range('AD42').value = 'Percent Male Condition'
	Range('AD43').value = '=Y43/(Y43+AA43)'
	Range('AD44').value = '=Y44/(Y44+AA44)'
	Range('AD45').value = '=Y45/(Y45+AA45)'
	Range('AD46').value = '=Y46/(Y46+AA46)'
	Range('AD47').value = '=Y47/(Y47+AA47)'
	Range('AD48').value = '=Y48/(Y48+AA48)'

	Range('AE42').value = "Meaning"
	Range('AE43').value = 'Isolated'
	Range('AE44').value = 'Sitting Next to One Female'
	Range('AE45').value = 'Sitting Next to One Male'
	Range('AE46').value = 'Sitting Next to One Female and One Male'
	Range('AE47').value = 'Sitting Next to All Females'
	Range('AE48').value = 'Sitting Next to All Males'

def cleanStuff(noc1,noc2,noc3,workbook,workbook1):
	#combines the groupwork sheets of the two observes into a third (already open) workbook and removes disparity in final sheet
	#done during Individual Analysis
	Sheet.add(noc3,before='Sheet1')
	haha = ['A', 'B', 'C', 'D', 'E','G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'T', 'U', 'V', 'W', 'X']
	i=0
	
	for i in range(0,22):
		for y in range(5,16):
			if Range(noc1, haha[i] + str(y),wkb=workbook).value == -1 or Range(noc2, haha[i] + str(y),wkb=workbook1).value == -1:
				Range(noc3, haha[i] + str(y)).value = -1
			
	#######################

	for i in range(0,22):
		for y in range(5,16):
			if Range(noc1, haha[i] + str(y),wkb=workbook).value > 0:
					Range(noc3, haha[i] + str(y)).value = 1
	for i in range(0,22):
		for y in range(5,16):
			if Range(noc2, haha[i] + str(y),wkb=workbook1).value > 0:
					Range(noc3, haha[i] + str(y)).value = 1

	for i in range(0,22):
		for y in range(5,16):
			if Range(noc1, haha[i] + str(y),wkb=workbook).value > 0 or Range(noc2, haha[i] + str(y),wkb=workbook1).value > 0:
					Range(noc3, haha[i] + str(y)).value = 1

def groupWork(noc1, noc2,workbook,workbook1):

	haha = ['A', 'B', 'C', 'D', 'E','G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'T', 'U', 'V', 'W', 'X']
	a = 0
	b = 0
	c = 0
	d = 0
	#noc1 is group
	#noc2 is gender
	##noc2 ,which is gender, is only choosing one of the sheets to compare groupwork and gender (second file)
	#done during Individual Analysis
	for i in range(0,22):
		for y in range(5,16):
			if Range(noc1, haha[i] + str(y),wkb=workbook).value > 0 and Range(noc2, haha[i] + str(y),wkb=workbook1).value == 1:
				a += 1
	Range('Z2').value = a

	for i in range(0,22):
		for y in range(5,16):
			if Range(noc1, haha[i] + str(y),wkb=workbook).value > 0 and Range(noc2, haha[i] + str(y),wkb=workbook1).value == -1:
				b += 1
	Range('Z3').value = b

	for i in range(0,22):
		for y in range(5,16):
			if Range(noc1, haha[i] + str(y),wkb=workbook).value == -1 and Range(noc2, haha[i] + str(y),wkb=workbook1).value == 1:
				c += 1
	Range('AA2').value = c

	for i in range(0,22):
		for y in range(5,16):
			if Range(noc1, haha[i] + str(y),wkb=workbook).value < 0 and Range(noc2, haha[i] + str(y),wkb=workbook1).value == -1:
				d += 1
	Range('AA3').value = d

def comparison(noc1,noc2,workbook,workbook1):
	#percent similiarity
	#compares between two group sheets ONLY FOR THE MIDDLE PART OF ROOM
	#all possible changes are added into numerator (which equals a)
	#all possible combinations are added into denomenator (which equals b)
	
	a=0
	b=0

	haha = ['G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R']
	
	for i in range(0,11):
		for y in range(5,16):
			'''
			-blanks are similar 
			-blanks and something marked similar



			'''
			#possible changes
			if Range(noc1, haha[i] + str(y),wkb=workbook).value == None and Range(noc2, haha[i] + str(y),wkb=workbook1).value == 1:
				a += 1
				b += 1
			
			if Range(noc1, haha[i] + str(y),wkb=workbook).value == 1 and Range(noc2, haha[i] + str(y),wkb=workbook1).value == None:
				a += 1
				b += 1
			
			if Range(noc1, haha[i] + str(y),wkb=workbook).value == None and Range(noc2, haha[i] + str(y),wkb=workbook1).value == -1:
				a += 1
				b += 1

			if Range(noc1, haha[i] + str(y),wkb=workbook).value == -1 and Range(noc2, haha[i] + str(y),wkb=workbook1).value == None:
				a += 1
				b += 1

			if Range(noc1, haha[i] + str(y),wkb=workbook).value ==1 and Range(noc2, haha[i] + str(y),wkb=workbook1).value == -1:
				a += 1
				b += 1

	
			if Range(noc1, haha[i] + str(y),wkb=workbook).value == -1 and Range(noc2, haha[i] + str(y),wkb=workbook1).value == 1:
				a += 1
				b += 1

			######
			if Range(noc1, haha[i] + str(y),wkb=workbook).value == 1 and Range(noc2, haha[i] + str(y),wkb=workbook1).value == 1:
				b += 1

	
			if Range(noc1, haha[i] + str(y),wkb=workbook).value == -1 and Range(noc2, haha[i] + str(y),wkb=workbook1).value == -1:
				b += 1
			
			if Range(noc1, haha[i] + str(y),wkb=workbook).value == None and Range(noc2, haha[i] + str(y),wkb=workbook1).value == None:
				b += 1
	#Range('AB2').value = str((float(b-a)/b)*100)  + "%"
	Range('AB2').value = (float(b-a)/b)*100

def genderANDgroupAnalysis(noc1,noc2,workbook,workbook1):
	#noc1 = from Individual Analysis and uses the Tables sheet
	#noc2 = from Overall Analysis
	#workbook = Individual Analysis
	#workbook1 = Overall Analysis

	#this work is done during Final Analysis
	haha = ['A', 'B', 'C', 'D', 'E','G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'T', 'U', 'V', 'W', 'X']
	haha1 = ['Z', 'AA', 'AB', 'AC', 'AD', 'AF', 'AG','AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN', 'AO','AP', 'AQ', 'AS','AT','AU','AV','AW']

	#Male Groupwork
	for i in range(0,22):
		for y in range(5,16):
			if Range(noc1, haha[i] + str(y),wkb=workbook).value > 0:
				Range(noc2, haha[i] + str(y)).value +=1
	
	#Male Non-Groupwork
	for i in range(0,22):
		for y in range(5,16):
			if Range(noc1, haha[i] + str(y+14),wkb=workbook).value == -1:
				Range(noc2, haha[i] + str(y+14)).value += 1

	#Female Groupwork
	for i in range(0,22):
		for y in range(5,16):
			if Range(noc1, haha1[i] + str(y),wkb=workbook).value > 0:
				Range(noc2, haha1[i] + str(y)).value +=1
	#Female Non-Groupwork
	for i in range(0,22):
		for y in range(5,16):
			if Range(noc1, haha1[i] + str(y+14),wkb=workbook).value == -1:
				Range(noc2, haha1[i] + str(y+14)).value += 1

	Sheet(noc2).autofit()

def trendData(noc1, noc2,noc3, workbook, workbook1):
	#add new sheet for raw gender and group data
	#do right before closing file(s)
	#noc1 is Groupwork
	#noc2 is Gender
	#noc3 is Tables (in the Individual Analysis File)
	#workbook 1 is first partners workbook
	#workbook 2 is second partners workbook



	Sheet.add(noc3,before='Sheet1')
	haha = ['A', 'B', 'C', 'D', 'E','G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'T', 'U', 'V', 'W', 'X']
	haha1 = ['Z', 'AA', 'AB', 'AC', 'AD', 'AF', 'AG','AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN', 'AO','AP', 'AQ', 'AS','AT','AU','AV','AW']

	#Male Groupwork
	Range('Tables','A4').value = "M-G"
	for i in range(0,22):
		for y in range(5,16):
			if Range(noc1, haha[i] + str(y),wkb=workbook).value > 0 and Range(noc2, haha[i] + str(y),wkb=workbook).value == 1:
				Range(noc3, haha[i] + str(y)).value = 1
	for i in range(0,22):
		for y in range(5,16):
			if Range(noc1, haha[i] + str(y),wkb=workbook1).value > 0 and Range(noc2, haha[i] + str(y),wkb=workbook1).value == 1:
				Range(noc3, haha[i] + str(y)).value = 1
	'''
	for i in range(0,22):
		for y in range(5,16):
			if Range(noc3, haha[i] + str(y)).value is None:
				Range(noc3, haha[i] + str(y)).value = 0
	'''

	#Male Non-Groupwork
	Range('Tables','A18').value = "M-NG"
	for i in range(0,22):
		for y in range(5,16):
			if Range(noc1, haha[i] + str(y),wkb=workbook).value == -1 and Range(noc2, haha[i] + str(y),wkb=workbook).value == 1:
				Range(noc3, haha[i] + str(y+14)).value = -1
	for i in range(0,22):
		for y in range(5,16):
			if Range(noc1, haha[i] + str(y),wkb=workbook1).value == -1 and Range(noc2, haha[i] + str(y),wkb=workbook1).value == 1:
				Range(noc3, haha[i] + str(y+14)).value = -1
	'''
	for i in range(0,22):
		for y in range(5,16):
			if Range(noc3, haha[i] + str(y+14)).value is None:
				Range(noc3, haha[i] + str(y+14)).value = 0
	'''

	#Female Groupwork
	Range('Tables','Z4').value = "F-G"
	for i in range(0,22):
		for y in range(5,16):
			if Range(noc1, haha[i] + str(y),wkb=workbook).value > 0 and Range(noc2, haha[i] + str(y),wkb=workbook).value == -1:
				Range(noc3, haha1[i] + str(y)).value = 1
	for i in range(0,22):
		for y in range(5,16):
			if Range(noc1, haha[i] + str(y),wkb=workbook1).value > 0 and Range(noc2, haha[i] + str(y),wkb=workbook1).value == -1:
				Range(noc3, haha1[i] + str(y)).value = 1
	'''
	for i in range(0,22):
		for y in range(5,16):
			if Range(noc3, haha1[i] + str(y)).value is None:
				Range(noc3, haha1[i] + str(y)).value = 0			
	'''
	#Female Non-Groupwork
	Range('Tables','Z18').value = "F-NG"
	for i in range(0,22):
		for y in range(5,16):
			if Range(noc1, haha[i] + str(y),wkb=workbook).value == -1 and Range(noc2, haha[i] + str(y),wkb=workbook).value == -1:
				Range(noc3, haha1[i] + str(y+14)).value = -1
	for i in range(0,22):
		for y in range(5,16):
			if Range(noc1, haha[i] + str(y),wkb=workbook1).value == -1 and Range(noc2, haha[i] + str(y),wkb=workbook1).value == -1:
				Range(noc3, haha1[i] + str(y+14)).value = -1

def genderORgroupAnalysis(noc1,noc2,workbook,workbook1):
	#noc1 = Combined Data ==Final
	#noc2 = Individual Data ==Final 2
	#workbook = Overall Analysis
	#workbook1 = Overall Analysis

	#Works because you don't need PHYS sheets to run it (independently works within Final Analysis workbook)
	haha = ['A', 'B', 'C', 'D', 'E','G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'T', 'U', 'V', 'W', 'X']
	haha1 = ['Z', 'AA', 'AB', 'AC', 'AD', 'AF', 'AG','AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN', 'AO','AP', 'AQ', 'AS','AT','AU','AV','AW']

	#Total Males
	for i in range(0,22):
		for y in range(5,16):
			if Range(noc1, haha[i] + str(y),wkb=workbook).value:
				Range(noc2, haha[i] + str(y)).value += Range(noc1, haha[i] + str(y),wkb=workbook).value

	for i in range(0,22):
		for y in range(5,16):
			if Range(noc1, haha[i] + str(y+14),wkb=workbook).value:
				Range(noc2, haha[i] + str(y)).value += Range(noc1, haha[i] + str(y+14),wkb=workbook).value

	#Total Female
	for i in range(0,22):
		for y in range(5,16):
			if Range(noc1, haha1[i] + str(y),wkb=workbook).value:
				Range(noc2, haha[i] + str(y+14)).value +=Range(noc1, haha1[i] + str(y),wkb=workbook).value
	for i in range(0,22):
		for y in range(5,16):
			if Range(noc1, haha1[i] + str(y+14),wkb=workbook).value:
				Range(noc2, haha[i] + str(y+14)).value +=Range(noc1, haha1[i] + str(y+14),wkb=workbook).value

	#Total Groupwork
	for i in range(0,22):
		for y in range(5,16):
			if Range(noc1, haha[i] + str(y),wkb=workbook).value:
				Range(noc2, haha1[i] + str(y)).value += Range(noc1, haha[i] + str(y),wkb=workbook).value
	for i in range(0,22):
		for y in range(5,16):
			if Range(noc1, haha1[i] + str(y),wkb=workbook).value:
				Range(noc2, haha1[i] + str(y)).value +=Range(noc1, haha1[i] + str(y),wkb=workbook).value
	#Total NonGroupwork
	for i in range(0,22):
		for y in range(5,16):
			if Range(noc1, haha[i] + str(y+14),wkb=workbook).value:
				Range(noc2, haha1[i] + str(y+14)).value +=Range(noc1, haha[i] + str(y+14),wkb=workbook).value
	for i in range(0,22):
		for y in range(5,16):
			if Range(noc1, haha1[i] + str(y+14),wkb=workbook).value:
				Range(noc2, haha1[i] + str(y+14)).value +=Range(noc1, haha1[i] + str(y+14),wkb=workbook).value					

	Sheet(noc2).autofit()

def conditions(noc1,noc2,workbook,workbook1):
	#noc1 = Individual Data ==Analysis
	#noc2 = Combined Data ==Final3
	#workbook = Individual Analysis == PHYS###A
	#workbook1 = Overall Analysis == Final Analysis

	#Male Groupwork
	Range(noc2,"B2", wkb=workbook1).value += Range(noc1, "Z2",wkb=workbook).value
	#Female Groupwork
	Range(noc2,"B3", wkb=workbook1).value += Range(noc1, "Z3",wkb=workbook).value
	#Male Non-Groupwork
	Range(noc2,"C2", wkb=workbook1).value += Range(noc1, "AA2",wkb=workbook).value
	#Female Non-Groupwork
	Range(noc2,"C3", wkb=workbook1).value += Range(noc1, "AA3",wkb=workbook).value

	###########
	#Conditions
	Range(noc2,"B6", wkb=workbook1).value += Range(noc1, "Z6",wkb=workbook).value
	Range(noc2,"B7", wkb=workbook1).value += Range(noc1, "Z7",wkb=workbook).value
	Range(noc2,"B8", wkb=workbook1).value += Range(noc1, "Z8",wkb=workbook).value
	Range(noc2,"B9", wkb=workbook1).value += Range(noc1, "Z9",wkb=workbook).value

	Range(noc2,"C6", wkb=workbook1).value = "=B6/(SUM(B6:B9))"
	Range(noc2,"C7", wkb=workbook1).value = "=B7/(SUM(B6:B9))"
	Range(noc2,"C8", wkb=workbook1).value = "=B8/(SUM(B6:B9))"
	Range(noc2,"C9", wkb=workbook1).value = "=B9/(SUM(B6:B9))"

	Range(noc2,"E2", wkb=workbook1).value +=1 

def conditionsGender(noc1,noc2,noc3,workbook,workbook1,workbook2):
	#noc1 is first partner  ~will be using gender
	#noc2 is second partner ~will be using gender
	#noc3 is combined individual analysis ~will be using conditions 
	haha = ['A', 'B', 'C', 'D', 'E','G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'T', 'U', 'V', 'W', 'X']
	i=0
	
	Range(noc3, "Y12", wkb=workbook2).value = "Male Condition 1"
	Range(noc3, "Z12", wkb=workbook2).value = "Female Condition 1"

	Range(noc3, "Y14", wkb=workbook2).value = "Male Condition 2"
	Range(noc3, "Z14", wkb=workbook2).value = "Female Condition 2"

	Range(noc3, "Y16", wkb=workbook2).value = "Male Condition 3"
	Range(noc3, "Z16", wkb=workbook2).value = "Female Condition 3"

	Range(noc3, "Y18", wkb=workbook2).value = "Male Condition 4"
	Range(noc3, "Z18", wkb=workbook2).value = "Female Condition 4"

	#Male Condition 1
	Range(noc3, "Y13", wkb=workbook2).value = 0
	for i in range(0,22):
		for y in range(5,16):
			if Range(noc3, haha[i] + str(y+13),wkb=workbook2).value ==1 and (Range(noc1, haha[i] + str(y),wkb=workbook).value == 1 or Range(noc2, haha[i] + str(y),wkb=workbook1).value == 1):
				Range(noc3, "Y13", wkb=workbook2).value +=1
			

	#Female Condition 1
	Range(noc3, "Z13", wkb=workbook2).value = 0
	for i in range(0,22):
		for y in range(5,16):
			if Range(noc3, haha[i] + str(y+13),wkb=workbook2).value ==1 and (Range(noc1, haha[i] + str(y),wkb=workbook).value == -1 or Range(noc2, haha[i] + str(y),wkb=workbook1).value == -1):
				Range(noc3, "Z13", wkb=workbook2).value +=1

	#Male Condition 2
	Range(noc3, "Y15", wkb=workbook2).value = 0
	for i in range(0,22):
		for y in range(5,16):
			if Range(noc3, haha[i] + str(y+13),wkb=workbook2).value ==2 and (Range(noc1, haha[i] + str(y),wkb=workbook).value == 1 or Range(noc2, haha[i] + str(y),wkb=workbook1).value == 1):
				Range(noc3, "Y15", wkb=workbook2).value +=1
			

	#Female Condition 2
	Range(noc3, "Z15", wkb=workbook2).value = 0
	for i in range(0,22):
		for y in range(5,16):
			if Range(noc3, haha[i] + str(y+13),wkb=workbook2).value ==2 and (Range(noc1, haha[i] + str(y),wkb=workbook).value == -1 or Range(noc2, haha[i] + str(y),wkb=workbook1).value == -1):
				Range(noc3, "Z15", wkb=workbook2).value +=1

	#Male Condition 3
	Range(noc3, "Y17", wkb=workbook2).value = 0
	for i in range(0,22):
		for y in range(5,16):
			if Range(noc3, haha[i] + str(y+13),wkb=workbook2).value ==3 and (Range(noc1, haha[i] + str(y),wkb=workbook).value == 1 or Range(noc2, haha[i] + str(y),wkb=workbook1).value == 1):
				Range(noc3, "Y17", wkb=workbook2).value +=1
			

	#Female Condition 3
	Range(noc3, "Z17", wkb=workbook2).value = 0
	for i in range(0,22):
		for y in range(5,16):
			if Range(noc3, haha[i] + str(y+13),wkb=workbook2).value ==3 and (Range(noc1, haha[i] + str(y),wkb=workbook).value == -1 or Range(noc2, haha[i] + str(y),wkb=workbook1).value == -1):
				Range(noc3, "Z17", wkb=workbook2).value +=1

	#Male Condition 4
	Range(noc3, "Y19", wkb=workbook2).value = 0
	for i in range(0,22):
		for y in range(5,16):
			if Range(noc3, haha[i] + str(y+13),wkb=workbook2).value ==4 and (Range(noc1, haha[i] + str(y),wkb=workbook).value == 1 or Range(noc2, haha[i] + str(y),wkb=workbook1).value == 1):
				Range(noc3, "Y19", wkb=workbook2).value +=1
			

	#Female Condition 1
	Range(noc3, "Z19", wkb=workbook2).value = 0
	for i in range(0,22):
		for y in range(5,16):
			if Range(noc3, haha[i] + str(y+13),wkb=workbook2).value ==4 and (Range(noc1, haha[i] + str(y),wkb=workbook).value == -1 or Range(noc2, haha[i] + str(y),wkb=workbook1).value == -1):
				Range(noc3, "Z19", wkb=workbook2).value +=1			

def conditionsGenderclass(noc1, noc2, workbook, workbook1):
	#noc1 is "Analysis" of Individual Class Analysis
	#noc2 is "Final3" of Whole Class Analysis
	#workbook is Individual Class Analysis
	#workbook1 is Whole Class Analysis workbook

	#"Male Condition 1"
	Range(noc2, "A22", wkb=workbook1).value += Range(noc1, "Y13", wkb=workbook).value
	#"Female Condition 1"
	Range(noc2, "B22", wkb=workbook1).value += Range(noc1, "Z13", wkb=workbook).value

	#"Male Condition 2"
	Range(noc2, "A24", wkb=workbook1).value	+= Range(noc1, "Y15", wkb=workbook).value
	#"Female Condition 2"
	Range(noc2, "B24", wkb=workbook1).value	+= Range(noc1, "Z15", wkb=workbook).value

	#"Male Condition 3"
	Range(noc2, "A26", wkb=workbook1).value += Range(noc1, "Y17", wkb=workbook).value
	#"Female Condition 3"
	Range(noc2, "B26", wkb=workbook1).value	+= Range(noc1, "Z17", wkb=workbook).value

	#"Male Condition 4"
	Range(noc2, "A28", wkb=workbook1).value	+= Range(noc1, "Y19", wkb=workbook).value
	#"Female Condition 4"
	Range(noc2, "B28", wkb=workbook1).value += Range(noc1, "Z19", wkb=workbook).value

def conditionsFrequency(noc1,noc2, workbook,workbook1):
	#noc1 is "Analysis" of Individual Class Analysis
	#noc2 is "Final4" of Whole Class Analysis
	#workbook is Individual Class Analysis
	#workbook1 is Whole Class Analysis workbook

	haha = ['A', 'B', 'C', 'D', 'E','G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'T', 'U', 'V', 'W', 'X']
	haha1 = ['Z', 'AA', 'AB', 'AC', 'AD', 'AF', 'AG','AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN', 'AO','AP', 'AQ', 'AS','AT','AU','AV','AW']


	#Condition 1
	for i in range(0,22):
		for y in range(5,16):
			if Range(noc1, haha[i] + str(y+13),wkb=workbook).value == 1:
				Range(noc2, haha[i] + str(y)).value +=1
	
	#Condition 2
	for i in range(0,22):
		for y in range(5,16):
			if Range(noc1, haha[i] + str(y+13),wkb=workbook).value == 2:
				Range(noc2, haha[i] + str(y+14)).value += 1

	#Condition 3
	for i in range(0,22):
		for y in range(5,16):
			if Range(noc1, haha[i] + str(y+13),wkb=workbook).value == 3:
				Range(noc2, haha1[i] + str(y)).value +=1

	#Condition 4
	for i in range(0,22):
		for y in range(5,16):
			if Range(noc1, haha[i] + str(y+13),wkb=workbook).value == 4:
				Range(noc2, haha1[i] + str(y+14)).value += 1

def conditionsGendermap(noc, noc1,noc2,noc3, workbook,workbook1):
	#noc is from "Tables" of Individual Class Analysis and uses gender
	#noc1 is from "Analysis" of Individual Class Analysis
	#noc2 is from "Final5" of Whole Class Analysis (Male Conditions)
	#noc3 is from "Final6" of Whole Class Analysis (Female Conditions)

	#workbook is Individual Class Analysis
	#workbook1 is Whole Class Analysis workbook
	
	haha = ['A', 'B', 'C', 'D', 'E','G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P','Q', 'R', 'T', 'U', 'V', 'W', 'X']
	haha1 = ['Z', 'AA', 'AB', 'AC', 'AD', 'AF', 'AG','AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN', 'AO','AP', 'AQ', 'AS','AT','AU','AV','AW']

	
	#Male Condition 1
	for i in range(0,22):
		for y in range(5,16):
			if Range(noc1, haha[i] + str(y+13),wkb=workbook).value == 1 and Range(noc, haha[i] + str(y),wkb=workbook).value == 1:
				Range(noc2, haha[i] + str(y),wkb=workbook1).value +=1
	
	#Male Condition 2
	for i in range(0,22):
		for y in range(5,16):
			if Range(noc1, haha[i] + str(y+13),wkb=workbook).value == 2 and Range(noc, haha[i] + str(y),wkb=workbook).value == 1:
				Range(noc2, haha[i] + str(y+14),wkb=workbook1).value += 1

	#Male Condition 3
	for i in range(0,22):
		for y in range(5,16):
			if Range(noc1, haha[i] + str(y+13),wkb=workbook).value == 3 and Range(noc, haha[i] + str(y+14),wkb=workbook).value == -1:
				Range(noc2, haha1[i] + str(y),wkb=workbook1).value +=1

	#Male Condition 4
	for i in range(0,22):
		for y in range(5,16):
			if Range(noc1, haha[i] + str(y+13),wkb=workbook).value == 4 and Range(noc, haha[i] + str(y+14),wkb=workbook).value == -1:
				Range(noc2, haha1[i] + str(y+14),wkb=workbook1).value += 1

	########

	#Female Condition 1
	for i in range(0,22):
		for y in range(5,16):
			if Range(noc1, haha[i] + str(y+13),wkb=workbook).value == 1 and Range(noc, haha1[i] + str(y),wkb=workbook).value == 1:
				Range(noc3, haha[i] + str(y),wkb=workbook1).value +=1
	
	#Female Condition 2
	for i in range(0,22):
		for y in range(5,16):
			if Range(noc1, haha[i] + str(y+13),wkb=workbook).value == 2 and Range(noc, haha1[i] + str(y),wkb=workbook).value == 1:
				Range(noc3, haha[i] + str(y+14),wkb=workbook1).value += 1

	#Female Condition 3
	for i in range(0,22):
		for y in range(5,16):
			if Range(noc1, haha[i] + str(y+13),wkb=workbook).value == 3 and Range(noc, haha1[i] + str(y+14),wkb=workbook).value == -1:
				Range(noc3, haha1[i] + str(y),wkb=workbook1).value +=1

	#Female Condition 4
	for i in range(0,22):
		for y in range(5,16):
			if Range(noc1, haha[i] + str(y+13),wkb=workbook).value == 4 and Range(noc, haha1[i] + str(y+14),wkb=workbook).value == -1:
				Range(noc3, haha1[i] + str(y+14),wkb=workbook1).value += 1




