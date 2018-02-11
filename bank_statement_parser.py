# Bank statement parser (Rabobank, CSV)
# 2018_2_11

# imports
import os # from std. library, os interactions
import sys # from std. library, interpretor interactions
import csv # from std. library
from datetime import date, datetime # from std. library, date and time functionality
import xlsxwriter # write xlsx-files

# files and settings
main_dir = sys.path[0]
trans_file = 'transacties.csv'
cats_file = 'categorie_instellingen.txt'

# set-up xlsx output file
workbook = xlsxwriter.Workbook(os.path.join(main_dir,'transacties_gecategoriseerd.xlsx'))
worksheet = workbook.add_worksheet()
xlsx_row = 0
xlsx_col = 0
date_format = workbook.add_format({'num_format': 'd-m-yyyy'})

# create cats_dict
cats_dict = {}
categories = csv.reader(open(os.path.join(main_dir, cats_file)))
for row in categories:
    try:
        skip = row[0]
    except IndexError:
        continue
    if skip.startswith('#'):
        continue
    cats_dict[row[0].strip().lower()] = row[1].strip().lower()

# read bank statement
statement = csv.reader(open(os.path.join(main_dir, trans_file)))
next(statement) # skip header
for row in statement:
    trans_date = datetime.strptime(row[4], '%Y-%m-%d').date()
    trans_month = trans_date.month
    trans_date = trans_date
    trans_counter_name = row[9]
    trans_descr = ' '.join(row[19:21]).strip()
    trans_amount = float(row[6].replace(',', '.'))
    trans = [trans_date, trans_month, trans_counter_name, trans_descr, trans_amount]
    trans_str = ';'.join(map(str, trans))

    # reference transaction with cats_dict
    for search_item, category in cats_dict.items():
        if trans_str.lower().find(search_item) > -1:
            cat = category
            break
    else:
        cat = 'NIET_GECATEGORISEERD'

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