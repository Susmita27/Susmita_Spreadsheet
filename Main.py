import spreadsheet_functionalities
import math

#New Sheet
sheet1 = spreadsheet_functionalities.Sheet(6,30)
                                                  
sheet1.updateValue(0,0,'13')                     
sheet1.updateValue(1,0,'12')                                                    
sheet1.updateValue(2,12,'=sum(A1:B3)')             
#sheet1.printTable()
sheet1.saveFile("sheet1.txt")

#New Sheet 

sheet3 = spreadsheet_functionalities.Sheet(4,4)
matrix_sheet= spreadsheet_functionalities.Sheet(4,4)

sheet3.cellUpdate('A1', '28')                      
sheet3.cellUpdate('B1', '30')
sheet3.cellUpdate('C1', '32')
sheet3.cellUpdate('A2', '34')
sheet3.cellUpdate('B2', '36')
sheet3.cellUpdate('C2', '38')
#sheet3.printTable()
sheet3.saveFile("sheet3.txt")

#implementing matrix transpose
print ("sheet 3 transpose")
matrix_sheet.matrixTranspose(sheet3)  
#matrix_sheet.printTable()                   

matrix_sheet.saveFile("matrix_transposesheet3.txt")


#New Sheet
sheet4 = spreadsheet_functionalities.Sheet(4,4)
sheet4.cellUpdate('A1', '5')
sheet4.cellUpdate('B1', '10')
sheet4.cellUpdate('A2', '15')
sheet4.cellUpdate('B2', '20')
sheet4.cellUpdate('A3', '25')
sheet4.cellUpdate('B3', '30')
sheet4.cellUpdate('C1', '24')
#sheet4.printTable()
sheet4.saveFile("sheet4.txt")

sheet4.matrixDifference(sheet3)
#sheet4.printTable()
sheet4.saveFile("updatedsheet4.txt")

#New Sheet

sheet5 = spreadsheet_functionalities.Sheet(4,4)
sheet5.cellUpdate('A1', '4')
sheet5.cellUpdate('B1', '12')
sheet5.cellUpdate('A2', '15')
sheet5.cellUpdate('B2', '22')
sheet5.cellUpdate('A3', '28')
sheet5.cellUpdate('B3', '21')
sheet5.cellUpdate('C1', '40')
sheet5.saveFile("sheet5.txt")
#sheet5.printTable()

#implementing matrix squareroot and printing round of value 
sheet5.matrixsqrt(sheet5)
#sheet5.printTable()
sheet5.saveFile("matrixsqrtsheet5.txt")


#resetting the whole sheet with a desired value
sheet5.resetvalue("1000")
#sheet5.printTable()
sheet5.saveFile("resetvaluesheet5.txt")



## New sheet
#playing around with Strings in the spreadsheet
sheet7=spreadsheet_functionalities.Sheet(4,4)

sheet7.cellUpdate('A1', 'A')
sheet7.cellUpdate('B1', 'fox')
sheet7.cellUpdate('C1', 'is')
sheet7.cellUpdate('D1', 'lying')
sheet7.cellUpdate('A2', 'across')
sheet7.cellUpdate('B2', 'the')
sheet7.cellUpdate('C2', 'shallow')
sheet7.cellUpdate('D2', 'river')

#looking for the value in a particular row and column value
value = sheet7.getCell(0,1)
print ( ("value of 1 row , 2 column = ") ,value)


#sheet7.printTable()
sheet7.saveFile("stringsheet7.txt")























