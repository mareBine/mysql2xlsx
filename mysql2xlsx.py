# coding=utf-8
# #############################################################
#
# export sql queries directly to xlsx
#
# usage: <script name> sql_file.sql
#
# TODO:
# - catch exception, if xls file already opened (permission denied)
# - writes None into NULL fields, it would be better if they were empty -> solution to put empty in DB instead of NULL
# - split into more files if there are more than 1M rows (xlsx limitation)
#
# #############################################################

from __future__ import print_function

import xlsxwriter
import sys
import MySQLdb
import time
import ConfigParser
import os

config = ConfigParser.ConfigParser()

# check if ini file exists
def checkIniFileExists(ini_file_name):
    if not os.path.exists(ini_file_name):
        sys.exit("ini file missing: xlsx.ini !!!")
    else:
        config.read('xlsx.ini')

# check for required options in ini file
def checkIniRequiredOptions():
    if (config.has_option('default', 'host')
        and config.has_option('default', 'port')
        and config.has_option('default', 'username')
        and config.has_option('default', 'password')
        and config.has_option('default', 'database')):
        return True
    else:
        sys.exit("ini file is missing all required db params: host, port, username, password, database")

# read from sql file to string
def readSqlFromFile(sqlFile):
    if os.access(sqlFile, os.R_OK):
        file = open(sqlFile, "r+")
        # TODO: to check if output file is already opened - doesn't work
        #if not open(output_file).closed:
        #    sys.exit("xlsx file " + output_file + " already opened !!!")
        str_sql = file.read()
        file.close()
        return str_sql
    else:
        sys.exit("sql file " + inputSqlFile + " not exists !!!")

checkIniFileExists('xlsx.ini')

# read variables from ini file
if checkIniRequiredOptions():
    host = config.get('default', 'host')
    port = config.getint('default', 'port')
    username = config.get('default', 'username')
    password = config.get('default', 'password')
    database = config.get('default', 'database')


# read sql query as argument
if len(sys.argv) == 2:
    inputSqlFile = str(sys.argv[1])
else:
    sys.exit("sql file missing, usage: <script name> sql_query.sql !!!")

# read sql from file
stringSql = readSqlFromFile(inputSqlFile)
outputFile = inputSqlFile + '.xlsx'

#print(stringSql.split(';'))

# mysql connection
try:
    db = MySQLdb.connect(host=host, port=port, user=username, passwd=password, db=database)
except MySQLdb.Error as err:
    print()
    print("MySQL error: {}".format(err))

#
# loop for "loop" defined in ini file
#
if config.has_option('loop','values'):
    loopList = eval(config.get('loop','values'))
# if there is no "loop", only out
else:
    loopList = ['out']

for loopValue in loopList:

    print()
    print(inputSqlFile + " -> " + loopValue + '_' + outputFile, end='\t')
    start = time.time()

    # string_to_numbers to convert digits to nimbers
    # constant_memory writes sequentially, should not be outofmemory errors with large files (drawback 1x slower)
    # https://github.com/jmcnamara/XlsxWriter/blob/master/dev/docs/source/workbook.rst
    workbook = xlsxwriter.Workbook(loopValue + '_' + outputFile, {
        'strings_to_numbers': True,
        'constant_memory': False,
        'default_date_format': 'yyyy-mm-dd'
    })

    # Add a bold format to use to highlight cells.
    bold = workbook.add_format({'bold': True})

    cur = db.cursor()
    cur.execute("SET NAMES utf8")

    sheetNumber = 1

    #
    # loop for more worksheets (more sql queries in 1 file)
    #
    for sql in stringSql.split(';'):

        # executes only non-empty sql queries and only SELECT ones
        if sql != '' and 'select' in sql.lower():
            worksheet = workbook.add_worksheet()

            if config.has_option('sheetName', 'sheet' + str(sheetNumber)):
                worksheet.name = config.get('sheetName', 'sheet' + str(sheetNumber))

            try:
                # of "loop" is defined, then switch values
                if config.has_option('loop','values') and config.has_option('loop','variable'):
                    sql = sql.replace(config.get('loop','variable'), "'" + loopValue + "'")
                cur.execute(sql)
                numrows = cur.rowcount
                print()
                print("'" + worksheet.name + "' (" + ("{:,}".format(numrows)) + " rows)")
                # end if there are >= 1M rows, this is xlsx limitation
                if numrows > 999999:
                    print("NOTE!!! xlsx can write only 1M rows, if there are more, they will be truncated !!!")

                print("Exporting data: 0%  0 seconds", end='\r')

                # write sql query result to xlsx
                col = 0
                for column in cur.description:
                    worksheet.write(0, col, str(column[0]), bold)
                    col += 1
                row = 1

                for rowdata in cur.fetchall():
                    col = 0
                    for coldata in rowdata:
                        worksheet.write(row, col, str(coldata).decode('utf-8'))
                        col += 1
                    row += 1
                    if row % 100 == 0:
                        elapsed = round(time.time() - start, 1)
                        procent = round((float(row) / numrows) * 100, 1)
                        #estimated = round((numrows * elapsed) / row, 2)
                        #print str(procent) + "%  " + str(elapsed) + " seconds    est: " + str(estimated)
                        print("Exporting data: " + str(procent) + "%  " + str(elapsed) + " seconds", end='\r')

                # validation: read from ini file, section 'validationN', xlsx writer syntax used
                for section in config.sections():
                    if 'validation' in section:
                        if config.get(section, 'sheetName') == worksheet.name:
                            worksheet.data_validation(config.get(section, 'range'), eval(config.get(section, 'rule')))

            except MySQLdb.Error as err:
                print()
                print("MySQL error: {}".format(err))

            sheetNumber += 1


    start = time.time()
    print()
    print("Writing file ...", end=' ')
    workbook.close()

    end = time.time()
    print("100%  " + str(round(end - start, 1)) + " seconds")