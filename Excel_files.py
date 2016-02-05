#The following code demonstrates how Python can be used to manipulate excel files
#using the openpyxl library.  Various commands to manipulate Excel files are available
#to change these files.



#import the openpyxl library which facilitates the manipulation of excel data.
import openpyxl as op



#Import numeric Python library which facilitates the manipulation of numeric data.
import numpy as np



#Create a variable to store the workbook, which is basically the excel file.
#data_only=True means that only cell values will be read, and not formulas.
#Data_only=False or left blank means that the value read in will be the formula.
workbook = op.load_workbook('Book1.xlsx', data_only=True)



#Access the desired worksheet using the name of the worksheet.
worksheet = workbook.get_sheet_by_name('Data')



#Can instead access the worksheet by which one is active.
worksheet = workbook.active



#Print the value stored in cell A5 (not the formula, since data_only=True).
print(worksheet['A5'].value)



#Calculate the number of rows in the worksheet using the max_row property
row_count = worksheet.max_row



#As a check, print the last row index.
print(row_count)



#Create an array to store cell values of Excel file.
a = np.array(range(row_count))



#For loop to print out a single column of data.
#In this case, only the first column is printed.
for cell in worksheet.columns[0]:
    print(cell.value)





#Fill an array using one column of data from the spreadsheet.
for i in range(row_count):
    a[i] = worksheet['A'+str(i+1)].value
    print(a[i])



#Set up a counter variable and fill it with zeros
counter = np.zeros(row_count)



#Set up an index increment for an array
i=0



#Check to see if column B has any numbers duplicating what is in column A.
#Use two nested for loops for the check.  Takes a while to get through it.
#for cell in worksheet.columns[0]:
    for cell_a in worksheet.columns[1]:
        if cell_a.value == cell.value:
            counter[i] = counter[i] +1
            print(cell)
    i=i+1



#How to find out the data type of a variable.
print(type(a[0]))



#Write data to an Excel file as values.
for i in range(row_count):
    worksheet['C'+str(i+1)] = i+1



#Write data to an Excel file as a formula  This formula will automatically be calculated in the spreadsheet.
for i in range(row_count):
    worksheet['D'+str(i+1)] = "=INT(RANDBETWEEN(1,100000))"



#Save the workbook.
workbook.save('Book1.xlsx')
