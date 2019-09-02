# /Users/azsha/Documents/HurayProject/hc_ios/HiHealthChallenge
# /Users/azsha/Documents/docs

import os
import openpyxl
import shutil

filepaths = []
filenames = []

for (path, dir, files) in os.walk("/Users/azsha/Documents/HurayProject/hc_ios/HiHealthChallenge/HiHealthChallenge"):
    for filename in files:
        ext = os.path.splitext(filename)[-1]
        if ext == '.swift':
            filepaths.append(path + '/' + filename)
            filenames.append(filename)

print(len(filenames))

filepaths.sort()

listExcelFilePath = "/Users/azsha/Documents/프로그램목록표.xlsx"

listExcel = openpyxl.load_workbook(listExcelFilePath)

listExcelSheet = listExcel.active

# 프로그램 명, 프로그램 설명 리스트 저장하기

# for i in range(0, len(filepaths)):

#     f = open(filepaths[i], 'r')

#     while True:
#         line = f.readline()
#         if not line: break

#         if "Program Name -" in line:
#             index = line.find("-")
#             iIndex = i + 2
#             listExcelSheet["C%d" % iIndex] = line[index + 2:-1]

#         if "Program Description -" in line:
#             index = line.find("-")
#             iIndex = i + 2
#             listExcelSheet["E%d" % iIndex] = line[index + 2:-1]

#     f.close()

# listExcel.save(listExcelFilePath)


# 프로젝트 명, 프로젝트 설명 뽑아내서 인수인계 문서 작성하기

for i in range(0, len(filepaths)):
    variables = []
    variableTypes = []
    variablesDetails = []
    
    functions = []
    functionReturns = []
    functionDetails = []
    
    newFilePath = "/Users/azsha/Documents/docs/" + filenames[i] + ".xlsx"
    shutil.copy('/Users/azsha/Documents/docs/base.xlsx', newFilePath)
    
    excel = openpyxl.load_workbook(newFilePath)
    excelSheet = excel.active
    
    excelSheet['G2'] = filenames[i]
    excelSheet['C5'] = filepaths[i][60:]
    
    f = open(filepaths[i] + '/' + filenames[i], 'r')
    
    isWriteClass = False
    
    while True:
        line = f.readline()
        if not line: break
        
        if "Program Name -" in line:
            index = line.find("-")
            excelSheet['C3'] = line[index + 2:-1]
        
        if "Program Description -" in line:
            index = line.find("-")
            excelSheet['C6'] = line[index + 2:-1]
    
        if "class" in line:
            if isWriteClass == False:
                words = line.split()
                if len(words) > 2:
                    if words[2].find(",") > 0:
                        excelSheet['G5'] = words[2][:-1]
                    else:
                        excelSheet['G5'] = words[2]
            
                else:
                    excelSheet['G5'] = "-"
        isWriteClass = True
        
        if "///" in line:
            index = line.find("///")
            detail = line[index + 4:]
            line = f.readline()
            
            if "func" in line:
                functionDetails.append(detail[:-1])
                
                words = line.split()
                index = words.index("func")
                funcWord = words[index + 1].split("(")
                functions.append(funcWord[0])
                
                if "->" in words:
                    index = words.index("->")
                    functionReturns.append(words[index + 1])
                
                else:
                    functionReturns.append("Void")

        if "var" in line:
            variablesDetails.append(detail[:-1])
            
            words = line.split()
                index = words.index("var")
                variables.append(words[index + 1][:-1])
                variableTypes.append(words[index + 2])

if len(variables) > 0:
    count = min(len(variables), 15)
    for i in range(0, count):
        index = i + 11
            excelSheet["A%d" % index] = variables[i]
            excelSheet["C%d" % index] = variableTypes[i]
            excelSheet["E%d" % index] = variablesDetails[i]
        
        if len(functions) > 0:
            count = min(len(functions), 10)
            for i in range(0, count):
                index = i + 28
                excelSheet["A%d" % index] = functions[i]
                excelSheet["C%d" % index] = functionReturns[i]
                excelSheet["E%d" % index] = functionDetails[i]

f.close()

excel.save(newFilePath)
