import csv
from prettytable import PrettyTable #For better visualisation

#with open('D:\Selenium Practices\Week 4\Sample.csv', mode='r') as file: #Check with this line for large set of data
with open('D:\Selenium Practices\Week 4\Sample2.csv', mode='r') as file: 
    csvFile = csv.reader(file)
    title = next(csvFile) #taking the first line to print it as header
    table = PrettyTable(title) #Table initialisation
    for value in csvFile:
        table.add_row(value)
    print(table)