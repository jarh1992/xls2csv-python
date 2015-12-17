from os import listdir
import xlrd
import csv
import os

#fn: file xls
#on:output file csv

directory = 'csv'
def csv_from_excel(fn, on):
   wb = xlrd.open_workbook(fn)
   sh = wb.sheet_by_index(0)
   if not os.path.exists(directory):
      os.makedirs(directory)
   completeName = os.path.join(directory, on+".csv")
   your_csv_file = open(completeName, 'wb')
   wr = csv.writer(your_csv_file, quoting=csv.QUOTE_ALL)
   for rownum in xrange(sh.nrows):
       wr.writerow([unicode(val).encode('utf8') for val in sh.row_values(rownum)])

   your_csv_file.close()


#leer todos los archivos (no carpetas) dentro de un directorio
i = 0
for file in listdir("."):
   print 'entro '+file
   if os.path.splitext(file)[1] != '':
      print 'a crear'
      csv_from_excel(file, os.path.splitext(file)[0])
   i = i+1
