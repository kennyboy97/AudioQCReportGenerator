import sys, os
from os.path import basename
from openpyxl import load_workbook
import Tkinter as tk
import tkFileDialog
from datetime import date

#dictionary of most common tracks to be QCed
QCTracks = {"51c":"5.1 Comp",
            "20c":"2.0 Comp",
            "5me":"5.1 M&E",
            "2me":"2.0 M&E",
            "7me":"5.1 & 2.0 M&E",
            "5dm":"5.1 DME",
            "2dm":"2.0 DME",
            "50d":"5.0 DX",
            "20d":"2.0 DX",
            "51m":"5.1 MX",
            "20m":"2.0 MX",
            "51f":"5.1 FX",
            "20f":"2.0 FX",
            "5fm":"5.1 FME",
            "2fm":"2.0 FME",
            "5ff":"5.1 FFX",
            "2ff":"2.0 FFX",
            "all":"All Tracks",
            "e51":"ENG 5.1",
            "eds":"ENG DS",
            "g51":"GER 5.1",
            "fds":"FRP DS"}

today = str(date.today())
QCdate = today[5:7]+"/"+today[8:]+"/"+today[:4]  #for QC date field inside report
QCRdate = today[:4]+today[5:7]+today[8:]  #for QC date added to final file name

print "Select the PT Marker TXT file"
txt = tkFileDialog.askopenfilename()
marker_source = open(txt, 'r')

print "Select the QC Report template"
xlsx = tkFileDialog.askopenfilename()
QCReport = load_workbook(xlsx)
New_File = basename(xlsx)  #file name of QC template selected
directory = os.path.dirname(xlsx)  #directory of QC template selected

QCStatus = raw_input("Does the QC Pass or Fail? [p/f]")

#will insert generic content listed below on line 128 into specific general notes section of QC template
General_Notes = raw_input("All general notes good? [y/n]")

TCList =[]
NotesList = []
RatingsList = []

#Ignore first 12 lines of the source TXT.
#Add LOCATION time code to TCList, NAME notes up to comma to Notes List, and NAME notes after comma to Ratings List
#If asterisk is in COMMENTS of TXT, ammend to Notes List

for i in xrange(12):
    marker_source.next()
for line in marker_source:
    TCList.append(line[5:16])
    c=line.find(",")
    ast=line.find("*")
    if ast > 0:
        NotesList.append(line[48:c]+" "+line[ast+1:-1])
    else:
        NotesList.append(line[48:c])
    RatingsList.append(line[c+2:c+5])
0
marker_source.close()

#Fill in new XLSX QC report from lists starting with row 30 of XLSX sheet.
#If Notes List item ends with any 3 charecter 'code' from the QCTracks dictionary, Channels column gets the appropriate entry
#If Notes List item starts with an X, Rating column gets an X added.
#If Notes List has Commercial black, Notes column gets TC of that list itme plus TC of next list item (TC Out) added
# then TC Out list item is ignored.

r=30
#range = len(TCList)
for i in range(len(TCList)):
    if r == 66:
        r=r+5
    TCc="A"+str(r)
    NLc="B"+str(r)
    RLc="W"+str(r)
    Fixed="Y"+str(r)
    Channel="U"+str(r)
    ws1 = QCReport.active
    if NotesList[i][0] == 'X' and NotesList[i][-3:] in QCTracks:
        ws1[TCc] = TCList[i]
        ws1[NLc] = NotesList[i][2:-3]
        ws1[Channel] = QCTracks[NotesList[i][-3:]]
        ws1[Fixed] = 'X'
        ws1[RLc] = RatingsList[i]
    elif NotesList[i][0] == 'X':
        ws1[TCc] = TCList[i]
        ws1[NLc] = NotesList[i][2:]
        ws1[Fixed] = 'X'
        ws1[RLc] = RatingsList[i]
    elif NotesList[i][-3:] in QCTracks:
        ws1[TCc] = TCList[i]
        ws1[NLc] = NotesList[i][:-3]
        ws1[Channel] = QCTracks[NotesList[i][-3:]]
        ws1[RLc] = RatingsList[i]
    elif "Commercial Black" in NotesList[i] or "Burn-In" in NotesList[i]:
        ws1[TCc] = TCList[i]
        ws1[NLc] = NotesList[i] + "  " + TCList[i] + " - " + TCList[i+1]
        ws1[RLc] = RatingsList[i]
    elif "TC Out" in NotesList[i]:
        continue
    else:
        ws1[TCc] = TCList[i]
        ws1[NLc] = NotesList[i]
        ws1[RLc] = RatingsList[i]

    r=r+1

#Fill in top portion of QC report based on user inputs above; my initials, Pass or Fail, General Notes if selected, QC date

ws1['R7'] = 'KH'
if QCStatus.lower()=='p':
    ws1['F6'] = 'X'
elif QCStatus.lower()=='f':
    ws1['R6'] = 'X'
else:
    print "Neither p nor f entered, QC Status will be left blank"
if General_Notes.lower()=='y':
    ws1['O18'] = 'No missing Dx'
    ws1['O19'] = 'Very good sync'
    ws1['O20'] = 'Depth of fill matches English guide'
    ws1['O21'] = 'Foreign Dx level matches English guide'
    ws1['O22'] = 'Pitch matches guide'
    ws1['O23'] = 'Matches guide'
elif General_Notes.lower()=='n':
    print "General notes not filled out."
else:
    print "Neither y nor n entered. General notes not filled out."
ws1['V11'] = QCdate

#Output final XLSX file with the QC date added to the end of the file name
# output to the same directory as the QC template

QCFile = New_File[:-5]
QCReport.save(os.path.join(directory, QCFile+"_"+QCRdate+".xlsx"))
