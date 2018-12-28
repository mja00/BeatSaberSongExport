import os
import json
import sys
import xlsxwriter
import time
from pathlib import Path

folders = 0
for _, dirnames, filenames in os.walk("."):
    folders += len(dirnames)

print('Detected '+ str(folders) +' folders. Calculating how long this will take.')
timeToComplete = folders * .02
print('This will take: ~' + str(timeToComplete) + ' seconds.')
time.sleep(1)

try:
    print('Removing the old song sheet if it\'s there.')
    time.sleep(2)
    os.remove('songSheet.xlsx')
    print('Success. Starting export.')
    time.sleep(1)
except:
    print('No file detected. Starting export.')
    time.sleep(1)

workbook = xlsxwriter.Workbook('songSheet.xlsx')
worksheet = workbook.add_worksheet()
# Add a bold format to use to highlight cells.
bold = workbook.add_format({'bold': True})
bold.set_pattern(1)
bold.set_bg_color('red')
bold.set_font_color('yellow')

#Setting row and column
row = 1
col = 0
failedSongs = 0
worksheet.write(0, 0, 'Song Name', bold)
worksheet.write(0, 1, 'Song Sub Name', bold)
worksheet.write(0, 2, 'Author', bold)
worksheet.write(0, 3, 'BPM', bold)

for root, dirs, files in os.walk(os.path.abspath(".")):
    for file in files:
        if file == "info.json":
            with open(os.path.join(root, file)) as f:
                try:
                    datastore = json.load(f)
                    songName = datastore["songName"]
                    songSubName = datastore["songSubName"]
                    authorName = datastore["authorName"]
                    bpm = datastore['beatsPerMinute']
                    worksheet.write(row, col, songName)
                    if songSubName == "":
                        worksheet.write(row, col + 1, 'None')
                    else:
                        worksheet.write(row, col + 1, songSubName)
                    worksheet.write(row, col + 2, authorName)
                    worksheet.write(row, col + 3, bpm)
                    row += 1
                    song = datastore["songName"] + ' - ' + datastore["songSubName"] + ' - ' + datastore["authorName"]
                    print(song + ' added to the sheet. #'+ str(row-1))
                    time.sleep(.02)
                except:
                    failedSongs += 1

workbook.close()
print('\n\n' + str (row - 1) + ' songs exported into the sheet.')
print(str (failedSongs) + ' songs failed to export. This could be for various reasons. Most common is the info.json isn\'t in correct format.')
print('Actual completion time: ' + str((row - 1) * .02) + ' seconds. Off by: ' + str(abs(timeToComplete - ((row - 1) * .02))) + ' seconds.')
