import pdfplumber
import PyPDF2
import os
import shutil
import time
import xlsxwriter

def get_raw_info():
    i = 0
    rawInfo = []
    outputInfo = [len(filesName)]
    print("开始提取数据")
    while i < len(filesName):
        fixedFileFullName = "input\\" + filesName[i]
        with pdfplumber.open(fixedFileFullName) as pdf:
            oneRawInfo = []
            table_settings = {
                "vertical_strategy": "text", 
                "horizontal_strategy": "lines",
                }
            for page in pdf.pages:
                tables = page.extract_tables(table_settings)
                for table in tables:
                    for row in table:
                        if row != ['', '', '', '', '', '', '', '', '', '', '', '']:
                            if row != ['', '', '', '', '', '', '', '', '', '', '']:
                                if row != ['', '', '', '', '', '', '', '', '', '']:
                                    if row != ['', '', '', '', '', '', '', '', '']:
                                        if row != ['', '', '', '', '', '', '', '']:
                                            if row != ['', '', '', '', '', '', '']:
                                                if row != ['', '', '', '', '', '']:
                                                    if row != ['', '', '', '', '']:
                                                        if row != ['', '', '', '']:
                                                            if row != ['', '', '']:
                                                                if row != ['', '']:
                                                                    if row != ['']:
                                                                        #print(row)
                                                                        infoInLine = [row]
                                                                        oneRawInfo += infoInLine
        rawInfo += [oneRawInfo]
        with pdfplumber.open(fixedFileFullName) as pdf:
            oneRawInfo = []
            table_settings = {
                "vertical_strategy": "text", 
                "horizontal_strategy": "text",
                }
            for page in pdf.pages:
                tables = page.extract_tables(table_settings)
                for table in tables:
                    for row in table:
                        if row != ['', '', '', '', '', '', '', '', '', '', '', '']:
                            if row != ['', '', '', '', '', '', '', '', '', '', '']:
                                if row != ['', '', '', '', '', '', '', '', '', '']:
                                    if row != ['', '', '', '', '', '', '', '', '']:
                                        if row != ['', '', '', '', '', '', '', '']:
                                            if row != ['', '', '', '', '', '', '']:
                                                if row != ['', '', '', '', '', '']:
                                                    if row != ['', '', '', '', '']:
                                                        if row != ['', '', '', '']:
                                                            if row != ['', '', '']:
                                                                if row != ['', '']:
                                                                    if row != ['']:
                                                                        #print(row)
                                                                        infoInLine = [row]
                                                                        oneRawInfo += infoInLine
        rawInfo += [oneRawInfo]
        i = i + 1
        print('\r' + '已提取' + str(i) + '/' + str(len(filesName)), end='', flush=True)
    print('\n提取完毕')
    return(rawInfo)

def get_info(raw):
    i = len(raw)
    j = 0
    allData = []
    while j < i:
        print(raw[j][0])
        if len(raw[j][0]) == 3:
            lineOfInfo = raw[j][0][2].split("\n")
            PONUM = lineOfInfo[0]
            DATE = lineOfInfo[1]
            if lineOfInfo[4] == "()":
                OurRef = lineOfInfo[5]
            else:
                OurRef = lineOfInfo[4]
            data = [PONUM, DATE, OurRef]
        if len(raw[j][0]) > 3:
            f = len(raw[j][0]) - 4
            lineOfInfo = raw[j][0][2].split("\n")
            PONUM = lineOfInfo[0]
            DATE = lineOfInfo[1]
            if lineOfInfo[3] == "()":
                OurRef = lineOfInfo[4]
            else:
                OurRef = lineOfInfo[3]
            while f > 0:
                c = 1
                lineOfInfo = raw[j][0][2+c].split("\n")
                PONUM = PONUM + lineOfInfo[0]
                DATE = DATE + lineOfInfo[1]
                if lineOfInfo[3] == "()":
                    OurRef = OurRef + lineOfInfo[4]
                else:
                    OurRef = OurRef + lineOfInfo[3]
                f = f - 1
                c = c + 1
            data = [PONUM, DATE, OurRef]
        allData = allData + [data]
        j = j + 2
    
    return(allData)



folder_path = "input"
filesName = os.listdir(folder_path)
output = get_raw_info()

#print(get_info(output))
'''
for i in range(len(output)):
    print(output[i])
    print("--------")
    i = i + 1'''

print(get_info(output))