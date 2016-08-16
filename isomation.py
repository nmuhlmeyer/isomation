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
    #filename = raw_input('Select a file: ')
    filename = 'matrix'
    FullFilename = filename + '.csv'
    NewFilename = filename + '_1.csv'
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
        #print dfTrim
        subset += 1
        InfoFile = 'matrixInfo_' + str(subset) + '.csv'        
        realFile = os.path.isfile(InfoFile)
        #print(dfTrim)
        #dfTrim.to_csv(NewFilename)
    #dfFullPivot.to_csv('matrixFull.csv')
    print dfFullPivot
    
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
    with open(FullFilename) as inF:
        for index,line in enumerate (inF):
            desired_line = re.search('(\d+,\d+)',line,re.M|re.I)
            if desired_line:
                desired_line2 = re.search('\d+,',line,re.M|re.I)
                if desired_line2:
                    snp_name1 = line[1:desired_line2.end()].strip(',')
                    desired_line3 = re.search(',\d+',line,re.M|re.I)
                    if desired_line3:
                        snp_name2 = line[desired_line2.end():desired_line3.end()]
                        desired_line4 = re.search('-\d+',line,re.M|re.I)
                        if desired_line4:
                            snp_iso = line[desired_line4.start():]
                            isoMatrix[i][0] = portNameDict[snp_name1]
                            isoMatrix[i][1] = portNameDict[snp_name2]
                            isoMatrix[i][2] = round(float(snp_iso),0)
                            i = i+1                        
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
    #numPorts = number_of_ports(len(isoMatrix))
    portOrderSorted = sorted(portOrder)
    desiredRowSorted = []
    for port in desiredRowName:
        desiredRowSorted.append(portOrderSorted.index(port))
    #print a
    desiredColumnSorted = []
    for port in desiredColumnName:
        desiredColumnSorted.append(portOrderSorted.index(port))
    #x = 3
    #print desiredColumn[x], desiredColumnName[x]
    #print desiredColumnSorted[x], portOrderSorted[desiredColumnSorted[x]]
    
    
    #print a
    #print portNames[a[2]]
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
def number_of_ports(result):
    num = 2
    while True:
        if result == num*num-num: break
        else: num += 1
    return num
def zerofix():
    with open('matrix.csv') as csvfile:
        inF = csv.reader(csvfile,delimiter=',')
        for row in inF:
            if row[-1]=='0.000': row[-1] = '-1'
    return

main()
#zerofix()
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

import random
def random_isoMatrix():
    #num = 4
    randIso = int(round(random.random()*-50))+10
    #print randIso
    randPort = int(round(random.random()*3))+2
    randPorts1 = int(round(random.random()*10))
    randPorts2 = int(round(random.random()*10))
    #print randPort, randIso
    #a = [num for num in range(num)]
    #a = [[x for x in range(randPort)] for x in range(randPort)]
    a = init_2dlist(randPort*(randPort-1),3)
    for i in range(randPort):
        for j in range(randPort):
            randIso = int(round(random.random()*-50))-10
            #a[]
            if i != j:
                a[j][0] = i+1
                a[j][1] = j+1
                a[j][2] = randIso
            #a = a.append(i)
    print a
    return
#random_isoMatrix()
