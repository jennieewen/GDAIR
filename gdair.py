import datetime
from BLK_long import Report_BLK_long
from BLK_short import Report_BLK_short
from BS_long import Report_BS_long
from BS_short import Report_BS_short


def validateDate(date_text):
    try:
        datetime.datetime.strptime(date_text, '%m/%d/%y')
    except ValueError:
        print("The Date you entered is invalid, please enter as MM-DD-YY format.")
        quit()
    

# ==========  START MAIN PROGRAM  ===========
print("==============================================")
print("|       GD Air Testing Data Process          |")
print("|                Version 1.0                 |")
print("==============================================")
print("")
reporting_date = input("Please enter Report Date (MM/DD/YY): ")
validateDate(reporting_date) 

analyzed_date = input("Please enter Data Analyzed Date (MM/DD/YY): ")
validateDate(analyzed_date)

batch = input("Pleased enter GD Air QC Batch: ")
print(batch)

print("===========================================================")
print("|                         MAIN  MENU                      |")
print("|                       ==============                    |")
print("|                                                         |")
print("|  1. Report of Method Blank Results - FULL               |")
print("|  2. Report of Method Blank Results - SHORT              |")
print("|  3. Blank Spike/Blank Spike Duplicate Results - FULL    |")
print("|  4. Blank Spike/Blank Spike Duplicate Results - SHORT   |")
print("|  0. Exit                                                |")
print("|                                                         |")
print("===========================================================")
choose = 1
while choose != 0 :
    choose = input("Please Enter Your Choose: ")
    if choose == "1":
        Report_BLK_long(reporting_date, analyzed_date, batch)
    elif choose == "2":
        Report_BLK_short(reporting_date, analyzed_date, batch)
    elif choose == "3":
        Report_BS_long(reporting_date, analyzed_date, batch)
    elif choose == "4":
        Report_BS_short(reporting_date, analyzed_date, batch)
    elif choose == "0":
        print("Thank you for using this program. Good-bye!")
        break

#============= END OF MAIN ==============
