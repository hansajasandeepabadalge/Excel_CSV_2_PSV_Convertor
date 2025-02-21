import csv
import os

os.makedirs('output', exist_ok=True)

filepath = 'file.csv'

with open(filepath, 'r', encoding='utf-8') as infile, open('output/output.csv', 'w', newline='', encoding='utf-8') as outfile:
    reader = csv.reader(infile, delimiter=',')
    writer = csv.writer(outfile, delimiter='|')
    writer.writerows(reader)