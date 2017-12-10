# Bank statement parser for Nadia
# 2017_12_10

import os # from std. library, os interactions
import sys # from std. library, interpretor interactions
import csv # from std. library

main_dir = sys.path[0]
trans_file = 'transactions.txt'

statement = csv.reader(open(os.path.join(main_dir, trans_file)))
for row in statement:
    trans_amount = float(row[4])
    trans_deb_cred = row[3]
    if trans_deb_cred == 'D':
        trans_amount = -trans_amount
    trans_date = row[7]
    trans_counter_name = row[6]
    trans_descr = ' '.join(row[10:16])
    
    print trans_amount, trans_date, trans_counter_name, trans_descr

# Try to regex counterpary
# use external file to save mappings, easily editable
    # gimsel/sasa = Boodschappen
    # svhw heffingen = Gemeentebelastingen
# import external file in dict to use for regexes
# Append to Excel, maybe move cursor to neewly imported files