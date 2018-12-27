import os
import json
import sys
import xlsxwriter
from pathlib import Path

os.remove('songSheet.xlsx')

workbook = xlsxwriter.Workbook('songSheet.xlsx')
worksheet = workbook.add_worksheet()

#Setting row and column
row = 1
col = 0

worksheet.write(0, 0, 'Song Name')
worksheet.write(0, 1, 'Song Sub Name')
worksheet.write(0, 2, 'Author')
worksheet.write(0, 3, 'BPM')

for root, dirs, files in os.walk(os.path.abspath(".")):
    for file in files:
        if file == "info.json":
            with open(os.path.join(root, file)) as f:
                try:
                    datastore = json.load(f)
                    #print os.path.join(root, file)
                    if datastore["songSubName"] == "":
                        songName = datastore["songName"]
                        authorName = datastore["authorName"]
                        bpm = datastore['beatsPerMinute']
                        song = datastore["songName"] + ' - ' + datastore["authorName"]
                        worksheet.write(row, col, songName)
                        worksheet.write(row, col + 1, 'None')
                        worksheet.write(row, col + 2, authorName)
                        worksheet.write(row, col + 3, bpm)
                        row += 1
                        print(song + ' added to the sheet.')
                    else:
                        songName = datastore["songName"]
                        songSubName = datastore["songSubName"]
                        authorName = datastore["authorName"]
                        bpm = datastore['beatsPerMinute']
                        worksheet.write(row, col, songName)
                        worksheet.write(row, col + 1, songSubName)
                        worksheet.write(row, col + 2, authorName)
                        worksheet.write(row, col + 3, bpm)
                        row += 1
                        song = datastore["songName"] + ' - ' + datastore["songSubName"] + ' - ' + datastore["authorName"]
                        print(song + ' added to the sheet.')
                        #file.write(datastore["songName"] + ' - ' + datastore["songSubName"] + ' - ' + datastore["authorName"])
                except:
                    pass

workbook.close()
