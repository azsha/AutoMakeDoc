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

for i in range(0, 2):
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
    excelSheet['C3'] = filenames[i]

    f = open(filepaths[i], 'r')
    while True:
        line = f.readline()
        if not line: break
        
        if filenames[i] in line:
            index = line.find("-")
            excelSheet['C6'] = line[index + 2:-1]

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