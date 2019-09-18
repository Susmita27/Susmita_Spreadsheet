# -*- coding: utf-8 -*-
"""
Created on Thu Aug 22 15:48:50 2019

@author: Susmita Dutta

Creating an excel sheet (spreadsheet) in  pythonic way and inluding spreadsheet functionalities
"""

import re
from Matrix import Matrix
import math
#from prettytable import PrettyTable



class NumberCell():
    def __init__(self, x):
        self.value = int(x)

    def __str__(self):
        return str(self.value)
    
# incase of string value this works  (additional function added ) 
class alphaCell():
    def __init__(self, x):
        self.value = x

    def __str__(self):
        return str(self.value)

class FormulaCell():
    def __init__(self, expr, sheet):  # expr is of the shape "=A1+B2" in string
        self._sheet = sheet
        self.formula = expr
        self.updateValue()
        
    def __str__(self):
        self.updateValue()
        return str(self.value)

    def updateValue(self):
        self.value = eval(self.addCalls(self.formula[1:]))

# using ASCII table, Converting  a letter into number. 
# The method is used to convert column names into int. It helps to find the associated spreadsheet cell in matrix.
    def colNameToInt(self, name):
        num = -1
        for i in range(-1, -len(name) - 1, -1):
            num += (ord(name[i]) - ord('A') + 1) * (26 ** (-i - 1))
            return num


# counter-conversion from number to letter ,example: 0=>A, 25=>Z, 26=>AB, 51=>AZ ...
    def colIntToName(self, digit):
        if digit < 26:
            return chr(digit + 65)
        else:
            part_two = chr((digit % 26) + 65)
            part_one = chr( (digit // 26)- 1 + 65)
            field_header= ("" + part_one + part_two)
            return field_header

    def sheetLookup(self, input1):
        result = []
        result.append('self._sheet.lookup(\'')
        result.append(str(input1))
        result.append('\')')
        resultString = ''.join(result)
        return resultString

    def generateLookup(self, input):
            alist = []
            cellA = input[:re.search(":", input).start()]
            cellB = input[re.search(":", input).start() + 1:]
            row_A =int( "".join(re.findall("[0-9]", cellA)))
            row_B =int( "".join(re.findall("[0-9]", cellB)))
            colA = self.colNameToInt( "".join(re.findall("[A-Z]", cellA)))
            colB = self.colNameToInt("".join(re.findall("[A-Z]", cellB)))
        
            for i in range(min(row_A, row_B), max(row_A, row_B) + 1, 1):
                for k in range(min(colA, colB), max(colA, colB) + 1, 1):
                    row_col = self.colIntToName(k) + str(i)
                    alist.append('[',self.Lookup(row_col),',',']')
            return (''.join(alist))


 # Converts the code into string, so that python can read.
 # A1 => self.lookup('A1')
 # A1 + B1 => self.lookup('A1') + self.lookup('B1')
 # sum([A1, B1, C1, D1]) => sum([ self.lookup('A1'), self.lookup('B1'), self.lookup('C1'), self.lookup('D1') ])
 # sum (A1 : D1) => sum([ self.lookup('A1'), self.lookup('B1'), self.lookup('C1'), self.lookup('D1') ])
    def addCalls(self, input):
        p = re.compile('[A-Z]+[1-9]+:[A-Z]+[1-9]+|[A-Z]+[1-9]+')
        matches = p.finditer(input)
        result = []
        prev = 0

        for match in matches:
            result.append(input[prev:match.start()])
            inputString = input[match.start():match.end()]
            pos = inputString.find(':')

            if pos < 0:
                y = self.sheetLookup(inputString)
                result.append(y)

            else:
                result.append(self.generateLookup(inputString))

            prev = match.end()
        result.append(input[prev:])
        resultString = ''.join(result)
        return resultString
     

############################################################################    
class Sheet(object):
    def __init__(self, nRows, nCols):
        self.rows = nRows
        self.cols = nCols
        self.matrix = Matrix(nRows, nCols)
        # fill the sheet with zero numbercells
        for row in range(self.rows):
            for col in range(self.cols):
                self.updateValue(row, col, "0")
                
    def __str__(self):
        return self.matrix.__str__()
    
    
    def updateValue(self, row, col, newValue):
        if newValue.isdigit():
            cellObject = NumberCell(newValue)
        elif newValue[0] != int:            # for string values to be accepted
            cellObject = alphaCell(newValue)
        else:
            cellObject = FormulaCell(newValue, self)

        self.matrix.setElementAt(row, col, cellObject)


    def cellUpdate (self, cell, newValue):
        colNumber = FormulaCell.colNameToInt(self,"".join(re.findall("[A-Z]", cell)))
        rowNumber =int( "".join(re.findall("[0-9]", cell)))-1
        Sheet.updateValue(self, rowNumber, colNumber, newValue)
    
    def lookup(self, x):
        p = re.compile('[A-Z]+')
        matches = p.match(x)
        to = matches.end()
        letters = x[:to]
        digits = x[to:]
        row = int(digits) - 1  # for 0 based matrix index
        col = FormulaCell.colNameToInt(self, letters)
        cell = self.matrix.getElementAt(row,col)
        return cell.value

   # def printTable(self):
    #    field_head = []
    #    for i in range(self.cols):
     #       field_head.append(FormulaCell.colIntToName(self, i))
    #    result = PrettyTable(field_head)
    #    
     #   for rows in range(self.rows):
       #     list = []
       #     for cols in range(self.cols):
       #         list.append(self.matrix.data[rows][cols].value)               
     #       result.add_row(list)
     #   print((str(result) + '\n'))

        

#######################  functionalities for spreadsheet  ###############################
        
#Resetting all the values of the spreadsheet :
    def resetvalue(self,value):
        for row in range(self.rows):
            for col in range(self.cols):
                self.updateValue(row, col, value)
               
#Difference between two spreadsheet:
                
    def matrixDifference(self,sheet2):
        for row in range(self.rows):
            for col in range(self.cols):
                    if self.cols == sheet2.cols and self.rows == sheet2.rows:
                        self.matrix.data[row][col].value -= sheet2.matrix.data[row][col].value
                    else:
                        print ("Matrices are incompatible, therefore cannot be subtracted" )   
                        
 
# Mathematical functions of excel executed on the values of the spreadsheet : 
#(squareroot of values which are rounded off )
                
    def matrixsqrt(self,sheet1):
        for row in range(self.rows):
            for col in range(self.cols):
               self.matrix.data[row][col].value = math.ceil(math.sqrt (sheet1.matrix.data[row][col].value ))   

                    
#Transposing one sheet (changing the values from rows to columns):
                        
    def matrixTranspose(self, sheet1):
        for rows in range(self.rows):
            for cols in range(self.cols):
                self.matrix.data[cols][rows].value = sheet1.matrix.data[rows][cols].value

              
#Looking up for a particular value by entering the row and column number            
    def getCell(self, rows, cols):
        return self.matrix.getElementAt(rows, cols)
    
    
#Saving the spreadsheet to check the output in the form of txt file            
    def saveFile(self, filename):
        with open(filename, 'w') as file:
            file.write(self.__str__())
            
            
    
            
            
            
                       
       
    

