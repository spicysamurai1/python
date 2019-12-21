import openpyxl
import os
import time
import csv

t = time.localtime()

curr_dir = os.path.dirname(__file__)
filename = os.path.join(curr_dir, "Log.csv")

if(os.path.isfile(filename)):
    csvfile = open(filename, 'a')
    csvwriter = csv.writer(csvfile)
else:
    fields = ['Time_of_logging', 'Message', 'Type_of_message']
    csvfile = open(filename, 'a')
    csvwriter = csv.writer(csvfile)
    csvwriter.writerow(fields)

ch = 0
while(ch != 2):
    l = []
    print("\n1. Enter Message And Type of Message\n2. Exit")
    ch = int(input("Enter your choice : "))
    if(ch == 1):
        message = input("Enter Message : ")
        l.append(message)
        type_of_message = input("Enter Type Of Message : ")
        l.append(type_of_message)
        current_time = time.strftime("%H:%M:%S", t)
        l.insert(0,current_time)
        csvwriter.writerow(l)
        print("Message logged successfully.")
    elif(ch == 2):
        csvfile.close()
    else:
        print("Enter Correct Choice.")
    


 
