#Get current month and year
from datetime import datetime
import os
import Spreadsheet_Controller

#Get month and year.
#The month string is for 'last' month so that the bill is for July's commissions when created in August
def get_month(i):
        switcher={
                1:'Janvier',
                2:'Fevrier',
                3:'Mars',
                4:'Avril',
                5:'Mai',
                6:'Juin',
                7:'Juillet',
                8:'Aout',
                9:'Septembre',
                10:'Octobre',
                11:'Novembre',
                12:'Decembre'
             }
        return switcher.get(i,"Invalid day of week")


def get_date(optional_month=0, optional_year=0):
    day = 1
    month_num = optional_month
    month = get_month(month_num)
    year = str(optional_year)
    return day, month, month_num, year

'''
This function is made to split the service line into multiple lines while maintaining
each word whole by seperating at a space character as opposed to the middle
of a word.
If a single word is longer than 60 characters without spaces it will still function.
'''
def string_split(input, length=60):
    new_split = []
    if (len(input) <= 60):
        new_split.append(input)
        return new_split
    words = input.split()
    line = ""
    last_word = ""
    for word in words:
        last_word = word
        if (len(word) + len(line)) <= length:
            line = line + " " + word
        else:
            new_split.append(line)
            line = ""

    if last_word in new_split[-1]:
        return new_split
    else:
        new_split.append(line)
        return new_split

def create_filename(therapist, month, year, rent):
    if not os.path.exists(therapist):
        os.makedirs(therapist)
    if not os.path.exists(therapist + os.sep + year):
        os.makedirs(therapist + os.sep + year)
    excel_fn = therapist + os.sep + therapist + "_totals.xlsx"
    if not os.path.isfile(excel_fn):
        Spreadsheet_Controller.create_workbook(excel_fn, year)

    file_directory = therapist + os.sep + year + os.sep
    if (rent):
        filename =  therapist + "_facture_loyer_" + month + "_" + year + ".pdf"
    else:
        filename =  therapist + "_facture_commission_" + month + "_" + year + ".pdf"

    return file_directory, filename