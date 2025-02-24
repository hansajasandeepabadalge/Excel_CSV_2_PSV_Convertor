import csv
import os
import openpyxl

os.makedirs('output', exist_ok=True)

csv_filepath = 'file.csv'
xlsx_filepath = 'file.xlsx'

# Convert CSV to PSV
if os.path.exists(csv_filepath):
    with open(csv_filepath, 'r', encoding='utf-8') as infile, open('output/output.csv', 'w', newline='', encoding='utf-8') as outfile:
        reader = csv.reader(infile, delimiter=',')
        writer = csv.writer(outfile, delimiter='|')
        writer.writerows(reader)

# Convert XLSX to PSV
if os.path.exists(xlsx_filepath):
    workbook = openpyxl.load_workbook(xlsx_filepath)
    sheet = workbook.active
    with open('output/output.csv', 'w', newline='', encoding='utf-8') as outfile:
        writer = csv.writer(outfile, delimiter='|')
        for row in sheet.iter_rows(values_only=True):
            writer.writerow(row)