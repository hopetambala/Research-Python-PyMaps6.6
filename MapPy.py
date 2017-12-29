#!/usr/bin/python

from numpy import *
from xlwings import *
from MapPyfuncs import *
import os
#version 6.6

print "Class Number"
input20 = raw_input()

Folder = ['A', 'B', 'C', 'D', 'E','F','G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R','S', 'T', 'U', 'V', 'W', 'X', 'Y','Z']
u = 0

#Analysis for Each Individual File
for i in range(len(Folder)):
	input10 = Folder[u]
	
	#Retrieves the directory path of the folder containing the files and appends the A1/A2 (etc) + file extension
	if os.path.exists(str(os.getcwd()) + "/" + input10 + str(1) + '.xlsx') and os.path.exists(str(os.getcwd()) + "/" + input10 + str(2) + '.xlsx'):
		#A1, B1, C1 etc
		wb = Workbook(r'' + str(os.getcwd()) + "/" + input10 + str(1) + '.xlsx', app_visible=None)
		#A2, B2, C2, etc
		wb1 = Workbook(r'' + str(os.getcwd()) + "/" + input10 + str(2) + '.xlsx', app_visible=None)

		input1 = "Group"

		input2 = "Group"

		input4 = "Gender"

		input3 = "Analysis"

		input5 = "Tables"
		noc = input3

		#PHYS[Class]A/B/C
		wb5 = Workbook(app_visible=None)
		
		
		cleanStuff(input1,input2,noc,wb,wb1)
		Sheet(noc).autofit()

		generateLeftrow(noc)
		generateMap('B','C','A', noc)
		generateMap('C','D','B', noc)
		generateMap('D','E','C', noc)
		generateMap('E','F','D', noc)

		generateMap('G','H','F', noc)
		generateMap('H','I','G', noc)
		generateMap('I','J','H', noc)
		generateMap('J','K','I', noc)
		generateMap('K','L','J', noc)
		generateMap('L','M','K', noc)
		generateMap('M','N','L', noc)
		generateMap('N','O','M', noc)
		generateMap('O','P','N', noc)
		generateMap('P','Q','O', noc)
		generateMap('Q','R','P', noc)
		generateMap('R','S','Q', noc)

		generateMap('T','U','S', noc)
		generateMap('U','V','T', noc)
		generateMap('V','W','U', noc)
		generateMap('W','X','V', noc)
		generateRightrow(noc)

		genderate("Gender","Gender","Analysis", wb,wb1,wb5)
		conditionsGenderate()

		
		conditionsTable()
		

		groupWork(input3,input4,wb,wb1)
		comparison(input1,input2,wb,wb1)


		conditionsGender(input4,input4,noc,wb,wb1,wb5)
		

		header(input1,wb1)

		Sheet(noc).autofit()
		
		trendData(input1, input4, input5, wb, wb1)

		Sheet(input5).autofit()

		#title = str(Range('Analysis','AE2', wkb=wb5).value) + str(Folder[u])
		title = "PHYS" + str(input20) + str(Folder[u])

		wb5.save(r'' + str(os.getcwd()) + "/" + "cleaned/" + title + '.xlsx')
		wb.close()
		wb1.close()

		u +=1
	
	elif os.path.exists(str(os.getcwd()) + "/" + input10 + str(1) + '.xlsx') is True and os.path.exists(str(os.getcwd()) + "/" + input10 + str(2) + '.xlsx') is False:
		u+=1

	elif os.path.exists(str(os.getcwd()) + "/" + input10 + str(1) + '.xlsx') and os.path.exists(str(os.getcwd()) + "/" + input10 + str(2) + '.xlsx') is False:
		break
	
	else:
		u +=1

#Analysis for All the Files
u = 0
c1 = []
c2 = []
c3 = []
c4 = []

MGW =[]
MNGW =[]
FGW =[]
FNGW =[]

PS = []
for i in range(len(Folder)):
	#The letter for the file name
	#wb = PHYS###
	#wb1 = Final Analysis File
	

	input10 = Folder[u]
	wb = Workbook(r'' + str(os.getcwd()) + "/cleaned/PHYS" + str(input20) + input10 + '.xlsx', app_visible=None)
	wb1 = Workbook(r'' + str(os.getcwd()) + "/cleaned/Final Analysis.xlsx", app_visible=None)

	if os.path.exists(str(os.getcwd()) + "/cleaned/PHYS" + str(input20) + input10 + '.xlsx'):
		

		genderANDgroupAnalysis("Tables","Final",wb,wb1)

		conditions("Analysis", "Final3", wb, wb1)
		conditionsGenderclass("Analysis", "Final3", wb, wb1)
		conditionsFrequency("Analysis","Final4", wb, wb1) 
		conditionsGendermap("Tables", "Analysis","Final5","Final6", wb,wb1)

		c1.append(Range("Analysis","Z6", wkb=wb).value)
		c2.append(Range("Analysis","Z7", wkb=wb).value)
		c3.append(Range("Analysis","Z8", wkb=wb).value)
		c4.append(Range("Analysis","Z9", wkb=wb).value)

		MGW.append(Range("Analysis","Z2", wkb=wb).value)
		MNGW.append(Range("Analysis","AA2", wkb=wb).value)
		FGW.append(Range("Analysis","Z3", wkb=wb).value)
		FNGW.append(Range("Analysis","AA3", wkb=wb).value)

		PS.append(Range("Analysis","AB2", wkb=wb).value)

		#########

		Range("Final3","B16", wkb=wb1).value = sum(c1) / float(len(c1))
		Range("Final3","B17", wkb=wb1).value = sum(c2) / float(len(c2))
		Range("Final3","B18", wkb=wb1).value = sum(c3) / float(len(c3))
		Range("Final3","B19", wkb=wb1).value = sum(c4) / float(len(c4))

		Range("Final3","B12", wkb=wb1).value = sum(MGW) / float(len(MGW))
		Range("Final3","B13", wkb=wb1).value = sum(FGW) / float(len(FGW))
		Range("Final3","C12", wkb=wb1).value = sum(MNGW) / float(len(MNGW))
		Range("Final3","C13", wkb=wb1).value = sum(FNGW) / float(len(FNGW))
		Range("Final3","D2", wkb=wb1).value = sum(PS) / float(len(PS))





		wb.close()
		#wb1.save()

		u+=1;
	elif os.path.exists(str(os.getcwd()) + "/cleaned/PHYS" + str(input20) + input10 + '.xlsx') is False:
		genderORgroupAnalysis("Final","Final2",wb1,wb1)

		break




print "Voila!"