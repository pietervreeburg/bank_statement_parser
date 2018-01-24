# Bank statement parser for Nadia
# 2018_1_24

# imports
import os # from std. library, os interactions
import sys # from std. library, interpretor interactions
import csv # from std. library
from datetime import date, datetime # from std. library, date and time functionality
import xlsxwriter # write xlsx-files

# files and settings
main_dir = sys.path[0]
trans_file = 'transactions.csv'
cats_file = 'categories.txt'

# set-up xlsx output file
workbook = xlsxwriter.Workbook(os.path.join(main_dir,'transactions_categorised.xlsx'))
worksheet = workbook.add_worksheet()
xlsx_row = 0
xlsx_col = 0
date_format = workbook.add_format({'num_format': 'd-m-yyyy'})

# create cats_dict
cats_dict = {}
categories = csv.reader(open(os.path.join(main_dir, cats_file)))
for row in categories:
    cats_dict[row[0].strip()] = row[1].strip()

# read bank statement
statement = csv.reader(open(os.path.join(main_dir, trans_file)))
for row in statement:
    trans_date = datetime.strptime(row[7], '%Y%m%d').date()
    trans_month = trans_date.month
    trans_date = trans_date
    trans_counter_name = row[6]
    trans_descr = ' '.join(row[10:16]).strip()
    trans_amount = float(row[4])
    trans_deb_cred = row[3]
    if trans_deb_cred == 'C':
        trans_amount = -trans_amount
    trans = [trans_date, trans_month, trans_counter_name, trans_descr, trans_amount]
    trans_str = ';'.join(map(str, trans))
    
    # reference transaction with cats_dict
    for search_item, category in cats_dict.items():
        if trans_str.lower().find(search_item) > -1:
            cat = category
            break
    else:
        cat = 'NO_CATEGORY'

    # output transaction to xlsx output file
    worksheet.write_datetime(xlsx_row, xlsx_col, trans_date, date_format)
    xlsx_col += 1
    for item in trans[1:]:
        worksheet.write(xlsx_row, xlsx_col, item)
        xlsx_col += 1
    worksheet.write(xlsx_row, xlsx_col, cat)
    xlsx_col = 0
    xlsx_row += 1

workbook.close()