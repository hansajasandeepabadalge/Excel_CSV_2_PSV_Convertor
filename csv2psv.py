import csv
import os

os.makedirs('output', exist_ok=True)

filepath = 'file.csv'

with open(filepath, 'r') as infile, open('output/output.csv', 'w', newline='') as outfile:
    reader = csv.reader(infile, delimiter=',')
    writer = csv.writer(outfile, delimiter='|')
    writer.writerows(reader)