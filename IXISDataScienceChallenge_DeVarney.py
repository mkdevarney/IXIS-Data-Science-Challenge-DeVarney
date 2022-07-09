#IXIS Data Science Challenge
#Written by Mary DeVarney

#import libraries
import csv
import pandas as pd
import openpyxl
from openpyxl.styles import Font
from string import ascii_uppercase

#A function that extracts the month and year from a full date range
def extract_date(date):
    parts = date.split("/")
    return int(parts[0]), int(parts[2])

#reads the csv files into dataframes
data_frame1 = pd.read_csv("DataAnalyst_Ecom_data_sessionCounts.csv")
data_frame2 = pd.read_csv("DataAnalyst_Ecom_data_addsToCart.csv")

#creates an empty dictionary to hold the data for the first worksheet
W1 = {}

#also creates a list of month names for better formating
months = ['January', 'February', 'March', 'April', 'May', 'June', 'July',
          'August', 'September', 'October', 'November','December'] 

#loops through the first csv's dataframe to extract data
x = 1
while x < len(data_frame1.index):
    #finds the date and creates a key for it with the months list
    m,y = extract_date(data_frame1.iloc[x]['dim_date'])
    key = months[m-1] + " 20" + str(y)

    #checks if the key exists in the dictionary
    #if not, adds it pointing to another dictionary with the devices as the keys
    #pointing to lists keeping track of sessions, transactions, QTY and ECR in that order
    if key not in W1:
        W1[key] = {'mobile': [0,0,0,0], 'desktop': [0,0,0,0], 'tablet': [0,0,0,0]}

    #adds the number of sessions, transacations and QTY to the lists
    W1[key][data_frame1.iloc[x]['dim_deviceCategory']][0] += int(data_frame1.iloc[x]['sessions'])
    W1[key][data_frame1.iloc[x]['dim_deviceCategory']][1] += int(data_frame1.iloc[x]['transactions'])
    W1[key][data_frame1.iloc[x]['dim_deviceCategory']][2] += int(data_frame1.iloc[x]['QTY']) 

    #adds 1 to iterate through the loop
    x+=1

#after looping through all of the data, we go back and calculate the ECR, storing it at the end of the lists
for k in W1:
    W1[k]['mobile'][3] = round(W1[k]['mobile'][1]/W1[k]['mobile'][0],6)
    W1[k]['desktop'][3] = round(W1[k]['desktop'][1]/W1[k]['desktop'][0],6)
    W1[k]['tablet'][3] = round(W1[k]['tablet'][1]/W1[k]['tablet'][0],6)

#creates a new workbook and the first worksheet using openpyxl
WB = openpyxl.Workbook()
S1 = WB.active
S1.title = 'Month Device Comparison'

#sets the values of the top row of the sheet
S1['B1'].value = 'Desktop Sessions'
S1['C1'].value = 'Desktop Transactions'
S1['D1'].value = 'Desktop QTY'
S1['E1'].value = 'Desktop ECR'
S1['F1'].value = 'Mobile Sessions'
S1['G1'].value = 'Mobile Transactions'
S1['H1'].value = 'Mobile QTY'
S1['I1'].value = 'Mobile ECR'
S1['J1'].value = 'Tablet Sessions'
S1['K1'].value = 'Tablet Transactions'
S1['L1'].value = 'Tablet QTY'
S1['M1'].value = 'Tablet ECR'

#loops through the dictionary to add the values to the spreadsheet
x = 2
for k in W1:
    #adds the month, year to the first column and bolds it
    S1['A' + str(x)].value = k
    S1['A' + str(x)].font = Font(bold=True)

    #adds all of the numerical data from the dictionary
    S1['B' + str(x)].value = W1[k]['desktop'][0]
    S1['C' + str(x)].value = W1[k]['desktop'][1]
    S1['D' + str(x)].value = W1[k]['desktop'][2]
    S1['E' + str(x)].value = W1[k]['desktop'][3]
    S1['F' + str(x)].value = W1[k]['mobile'][0]
    S1['G' + str(x)].value = W1[k]['mobile'][1]
    S1['H' + str(x)].value = W1[k]['mobile'][2]
    S1['I' + str(x)].value = W1[k]['mobile'][3]
    S1['J' + str(x)].value = W1[k]['tablet'][0]
    S1['K' + str(x)].value = W1[k]['tablet'][1]
    S1['L' + str(x)].value = W1[k]['tablet'][2]
    S1['M' + str(x)].value = W1[k]['tablet'][3]

    #adds 1 to the index number    
    x+=1

#loop to bold the top row and adjust the column width
letter = 65
while letter < 78:
    S1.column_dimensions[chr(letter)].width = 20
    S1[chr(letter) + '1'].font = Font(bold=True)
    letter += 1

#creates lists to hold the information for the last two months for the second worksheet
#the order is date, sessions, transactions, QTY, ECR, Adds to Cart
d1 = 'June 2013'
d2 = 'May 2013'
M1 = [d1, W1[d1]['desktop'][0]+ W1[d1]['mobile'][0]+W1[d1]['tablet'][0], W1[d1]['desktop'][1]+ W1[d1]['mobile'][1]+W1[d1]['tablet'][1],
      W1[d1]['desktop'][2]+ W1[d1]['mobile'][2]+W1[d1]['tablet'][2], round((W1[d1]['desktop'][1]+ W1[d1]['mobile'][1]+W1[d1]['tablet'][1])/
      (W1[d1]['desktop'][0]+ W1[d1]['mobile'][0]+W1[d1]['tablet'][0]),6), data_frame2.iloc[11]['addsToCart']]
M2 = [d2, W1[d2]['desktop'][0]+ W1[d2]['mobile'][0]+W1[d2]['tablet'][0], W1[d2]['desktop'][1]+ W1[d2]['mobile'][1]+W1[d2]['tablet'][1],
      W1[d2]['desktop'][2]+ W1[d2]['mobile'][2]+W1[d2]['tablet'][2], round((W1[d2]['desktop'][1]+ W1[d2]['mobile'][1]+W1[d2]['tablet'][1])/
      (W1[d2]['desktop'][0]+ W1[d2]['mobile'][0]+W1[d2]['tablet'][0]), 6), data_frame2.iloc[10]['addsToCart']]

#create lists for the absolute and relative differences
ADiff = ['Absolute Differences', M1[1]-M2[1], M1[2]-M2[2], M1[3]-M2[3], round(M1[4]-M2[4],6), M1[5]-M2[5]]
RDiff = ['Relative Differences', round((ADiff[1]/M2[1])*100, 2), round((ADiff[2]/M2[2])*100, 2),round((ADiff[3]/M2[3])*100, 2),
         round((ADiff[4]/M2[4])*100, 2),round((ADiff[5]/M2[5])*100, 2)]

#creates new sheet for the workbook
S2 = WB.create_sheet()
S2.title = 'Month over Month Comparison'

#sets the values of the top row of the sheet
S2['B1'].value = 'Sessions'
S2['C1'].value = 'Transactions'
S2['D1'].value = 'QTY'
S2['E1'].value = 'ECR'
S2['F1'].value = 'Adds to Cart'

#loops through the lists and adds them to the worksheet
x = 0
y = 0
letter = 65
while letter < 71:
    S2[chr(letter)+str(x+2)].value = M1[y]
    S2[chr(letter)+str(x+3)].value = M2[y]
    S2[chr(letter)+str(x+4)].value = ADiff[y]
    S2[chr(letter)+str(x+5)].value = RDiff[y]
    letter += 1
    y += 1

#loop to bold the top row and adjust the column width
letter = 65
while letter < 71:
    S2.column_dimensions[chr(letter)].width = 20
    S2[chr(letter) + '1'].font = Font(bold=True)
    letter += 1

#loop to bold the first column
x = 2
while x < 6:
    S2['A' + str(x)].font = Font(bold=True)
    x += 1

#saves the workbook
WB.save('IXISDataScienceChallenge.xlsx')





