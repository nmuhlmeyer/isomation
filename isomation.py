'''**********************************************************************


    This script reads an ADS exported *.csv file and creates pivot tables
    based on the exported data.
    Copyright (C) 2016  Nicholas Muhlmeyer

    This program is free software: you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation, either version 3 of the License, or
    (at your option) any later version.

    This program is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU General Public License for more details.

    You should have received a copy of the GNU General Public License
    along with this program.  If not, see <http://www.gnu.org/licenses/>.
**********************************************************************'''
import re
import pandas
import numpy
import csv
import os.path

def main():
    filename = raw_input('Select a file: ')
    #filename = 'matrix'
    FullFilename = filename + '.csv'
    isoMatrix = extract_isolation(FullFilename)
    df = pandas.DataFrame(isoMatrix,columns=['Port 1','Port 2','Isolation'])
    dfFullPivot = df.pivot('Port 1', 'Port 2')
    realFile = os.path.isfile('matrixInfo_1.csv')
    subset = 1
    while realFile:
        NewFilename = 'matrixSub_' + str(subset) + '.csv'
        rowdata=[]
        with open('matrixInfo_' + str(subset) + '.csv') as csvfile:
            inF = csv.reader(csvfile,delimiter=',')
            for row in inF:
                rowdata.append(row)
        desiredRow=[int(n) for n in rowdata[0]]
        desiredColumn=[int(n) for n in rowdata[1]]
        desiredRowSorted, desiredColumnSorted = subsetOrdering(isoMatrix,desiredRow,desiredColumn)
        dfTrim = pivotSubset(dfFullPivot,desiredRowSorted,desiredColumnSorted)
        subset += 1
        InfoFile = 'matrixInfo_' + str(subset) + '.csv'        
        realFile = os.path.isfile(InfoFile)
        print(dfTrim)
        #dfTrim.to_csv(NewFilename)    
    print dfFullPivot
    #dfFullPivot.to_csv('matrix' + 'Full.csv')
    
    return

'''**************************************************'''

def extract_isolation(FullFilename):
    portNameDict, _ = portNameDictionary()
    with open(FullFilename) as inF:
        points = 0
        for index,line in enumerate(inF):
            desired_line = re.match('Num. Points : \D',line,re.M|re.I)
            if desired_line:
                desired_line2 = re.search(']',line,re.M|re.I)
                if desired_line2: points = line[desired_line.end():desired_line2.start()]
    points = int(points)
    isoMatrix = init_2dlist(points,3)
    i = 0
    with open(FullFilename) as csvFile:
        inF = csv.reader(csvFile,delimiter=',')
        for row in inF:
            if i > 0:
                if row[-1]=='0.000': row[-1] = '-1'
                isoMatrix[i-1][0] = portNameDict[row[0][1:]]                
                isoMatrix[i-1][1] = portNameDict[row[1][:-1]]
                isoMatrix[i-1][2] = round(float(row[-1]),0)
            i += 1
    return isoMatrix

def portNameDictionary():
    filename = 'portName.csv'
    with open(filename) as csvFile:
        line = csv.reader(csvFile)
        portInfo = list(line)
        numberPorts = numpy.size(portInfo)/len(portInfo)
        highestPort = int(portInfo[-1][-1])
        portNameDict = {}
        for j in range(numberPorts):
            portNameDict[portInfo[1][j]] = portInfo[0][j]
    return portNameDict, highestPort

def pivotSubset(dfFullPivot,desiredRow,desiredColumn):
    indexRow = init_count_list(len(dfFullPivot))
    indexColumn = init_count_list(len(dfFullPivot))
    [indexRow.pop(k) for k in sorted(desiredRow,reverse=True)]
    [indexColumn.pop(k) for k in sorted(desiredColumn,reverse=True)]
    dfTrim = dfFullPivot.drop(dfFullPivot.index[indexRow]) # to get desired rows
    dfTrim = dfTrim.transpose() #to be abel to edit columns
    dfTrim = dfTrim.drop(dfTrim.index[indexColumn]) # to get desired columns
    dfTrim = dfTrim.transpose()
    return dfTrim

def subsetOrdering(isoMatrix,desiredRow,desiredColumn):
    [portNameDict, highestPort] = portNameDictionary()
    desiredRowName = init_list(len(desiredRow))
    desiredColumnName = init_list(len(desiredColumn))
    portOrder=[]
    i = 0
    for j in range(highestPort):
        i += 1
        if str(j+1) in portNameDict.keys(): 
            i += -1
            portOrder.append(portNameDict[str(j+1)])
    for j in range(len(desiredRowName)): desiredRowName[j] = portNameDict[str(desiredRow[j])]
    for j in range(len(desiredColumnName)): desiredColumnName[j] = portNameDict[str(desiredColumn[j])]
    portOrderSorted = sorted(portOrder)
    desiredRowSorted = []
    for port in desiredRowName:
        desiredRowSorted.append(portOrderSorted.index(port))
    desiredColumnSorted = []
    for port in desiredColumnName:
        desiredColumnSorted.append(portOrderSorted.index(port))
    return desiredRowSorted, desiredColumnSorted

def init_2dlist(num_rows,num_cols):
    list_var = [[0 for j in range(num_cols)] for i in range(num_rows)]
    return list_var
def init_list(num_rows):
    list_var = [0 for i in range(num_rows)]
    return list_var
def init_count_list(num_elements):
    list_var = [k for k in range(num_elements)]
    return list_var

main()
'''**************************************************'''

def inverse_factorial(x):
    if x < 1: return 'Error: x less than 1'
    y = 0    
    while x > 1:
        y = y + 1        
        x = x/y
        #if x != 1: return 'Error: x is not a factorial'
    if y == 0: y = 1
        #y = y + 1
    return y
