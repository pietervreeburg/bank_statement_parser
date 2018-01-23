# Bank statement parser for Nadia
# 2018_1_23

# imports
import os # from std. library, os interactions
import sys # from std. library, interpretor interactions
import csv # from std. library
from datetime import date, datetime # from std. library, date and time functionality
# import xlsxwriter # write xlsx-files

# files and settings
main_dir = sys.path[0]
trans_file = 'transactions.txt'
cats_file = 'categories.txt'

# set-up xlsx output file
# workbook = xlsxwriter.Workbook('transactions_categorised_DATE.xlsx')
# worksheet = workbook.add_worksheet()
# row = 0
# col = 0

# create cats_dict
cats_dict = {}
categories = csv.reader(open(os.path.join(main_dir, cats_file)))
for row in categories:
    cats_dict[row[0]] = row[1]

# read bank statement
statement = csv.reader(open(os.path.join(main_dir, trans_file)))
for row in statement:
    trans_counter_name = row[6]
    trans_descr = ' '.join(row[10:16]).strip()
    trans_amount = row[4].replace('.', ',')
    trans_deb_cred = row[3]
    if trans_deb_cred == 'D':
        trans_amount = '-{}'.format(trans_amount)
    trans_date = datetime.strptime(row[7], '%Y%m%d').date()
    trans = '{};{};{};{}'.format(trans_date, trans_counter_name, trans_descr, trans_amount)
    
    # reference transaction with cats_dict
    for search_item, category in cats_dict.items():
        if trans.lower().find(search_item) > -1:
            trans = '{};{}'.format(trans, category)
            break
    else:
        trans = '{};NO_CATEGORY'.format(trans)

    print trans
    
    # output transaction to xlsx output file
    # for item in trans.split(';'):
        # worksheet.write(row, col, item)
        # col += 1
    # row += 1

# workbook.close()