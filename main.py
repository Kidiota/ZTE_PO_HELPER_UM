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
        
        #插入子项目
        #PSOS = raw[j+1].index(['Item', 'Material', '', 'Description', 'Quantity', 'UOM', 'Unit Pr'])
        #print(raw[j+1][200:])
        #print("\n\n\n")
        subItems = []
        item = 0

        itemNum = []
        material = []
        description = []
        quantity = []
        UOM = []

        print(raw[j+1])
        print("\n\n\n")
        while item < len(raw[j+1]):
            d = 0
            if raw[j+1][item][d][0:2] == '00':
                itemNum = raw[j+1][item][d]
                d = d + 1
                if raw[j+1][item][d] != '':
                    material = raw[j+1][item][d]
                    d = d + 1
                    if raw[j+1][item][d] != '':
                        r = d
                        description = raw[j+1][item][r]
                        while not raw[j+1][item][r+1][0].isdigit():
                            description = description + raw[j+1][item][d+1]
                            r = r + 1
                            if r + 1 >= len(raw[j+1][item]) - 1:
                                break
                        if raw[j+1][item+1][0:1] == ['','']:
                            l = 0
                            while raw[j+1][item+1][l] != '':
                                description = description + raw[j+1][item+1][l]
                                l = l + 1
                        d = d + 1
                        if raw[j+1][item][d] != '':
                            quantity = raw[j+1][item][d]
                            d = d + 1
                            if raw[j+1][item][d] != '':
                                UOM = raw[j+1][item][d]
                                d = d + 1
            subItems = subItems + [[itemNum, material, description, quantity, UOM]]
                                        
            item = item + 1

        data = data + [subItems]
        
        allData = allData + [data]
        #print(allData)
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