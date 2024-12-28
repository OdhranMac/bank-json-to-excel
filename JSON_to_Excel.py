"""
This file transforms numerous json data structure into one excel table.
The original purpose was to take JSON outputs from Azure OCR for bank statements, to be used for financial analysis.

You'll need to enter your directory in 'rootDir'.
"""

import json
import pprint as pp
import pandas as pd
import numpy as np
import os

array = np.chararray(shape = (50000, 20), itemsize = 1000, unicode = True)
docNumber = 0
index = 0
subFolderNum = 0

rootDir = r''

dirTree = os.walk(rootDir)

for dirPaths, subPaths, files in dirTree:

    for subPath in subPaths:

        for subDirPaths, subDirs, subFiles in os.walk(os.path.join(rootDir, subPath)):
            for subFile in subFiles:
                print(subFile)
                
                docNumber += 1

                filePath = os.path.join(subDirPaths, subFile)
                JSON_file = open(filePath)
                JSON_data = json.load(JSON_file)

                for i in range(len(JSON_data['tables'])):
            
                    try:
                        for j in range(len(JSON_data['tables'][str(i+1)][0]['cells'])):
                            row = JSON_data['tables'][str(i+1)][0]['cells'][j]['row']
                            column = JSON_data['tables'][str(i+1)][0]['cells'][j]['column']
                            value = JSON_data['tables'][str(i+1)][0]['cells'][j]['text']
                            array[index + row, column] = value                                    
                    except KeyError:
                        continue

                    index += JSON_data['tables'][str(i+1)][0]['rows']
                    print('index: ' + str(index))

        print('Number of documents: ' + str(docNumber))

        df = pd.DataFrame(data = array, dtype = str, index = None, columns = None)

        print('save location: ' + rootDir + '\\' + subPath)
        df.to_excel(rootDir + '\\' + subPath + '.xlsx', header = False, index = False)
        print('wrote to excel')

        subFolderNum += 1
        docNumber = 0
        index = 0

print("COMPLETE")