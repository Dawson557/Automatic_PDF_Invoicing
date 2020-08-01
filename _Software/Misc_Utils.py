#Get current month and year
from datetime import datetime

#Get month and year.
#The month string is for 'last' month so that the bill is for July's commissions when created in August
def get_month(i):
        switcher={
                2:'Janvier',
                3:'Fevrier',
                4:'Mars',
                5:'Avril',
                6:'Mai',
                7:'Juin',
                8:'Juillet',
                9:'Aout',
                10:'Septembre',
                11:'Octobre',
                12:'Novembre',
                1:'Decembre'
             }
        return switcher.get(i,"Invalid day of week")

def get_date(optional_month=0, optional_year=0):
    if optional_month == 0:
        today = datetime.today()
        day = today.day
        month_num = today.month
        month = get_month(month_num)
        year = str(today.year)
    elif optional_year == 0:
        today = datetime.today()
        day = 1
        month_num = optional_month
        month = get_month(month_num)
        year = str(today.year)
    else: #month and year were manually set
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
