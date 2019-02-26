import os
import xlrd
import xlwt
from datetime import date,datetime

output = open('data.txt', 'w')
output.write('name, surface, status, finger, part, usability, memorability, preference\n')

def read_finger(data, excel_name, status, part, finger):
    tags = excel_name.split('-')
    name = tags[1]
    surface = tags[0]
    print name, surface, status, finger, part, data[0], data[1], data[2]
    output.write(name)
    output.write(', ')
    output.write(surface)
    output.write(', ')
    output.write(status)
    output.write(', ')
    output.write(finger)
    output.write(', ')
    output.write(part)
    output.write(', ')
    output.write(str(data[0]))
    output.write(', ')
    output.write(str(data[1]))
    output.write(', ')
    output.write(str(data[2]))
    output.write('\n')

def read_part(data, excel_name, status, part):
    finger_name = ['thumb', 'index', 'middle', 'ring', 'pinkie', 'two', 'three']
    for i in range(0, 7):
        read_finger(data[i * 3 : i * 3 + 3], excel_name, status, part, finger_name[i])

def read_status(sheet, excel_name, status, rows_begin):
    fingerpulp = sheet.col_values(3)[rows_begin : rows_begin + 21]
    fingertip = sheet.col_values(5)[rows_begin : rows_begin + 21]
    fingerside = sheet.col_values(7)[rows_begin : rows_begin + 21]
    knuckle = sheet.col_values(9)[rows_begin : rows_begin + 21]
    fingernail = sheet.col_values(11)[rows_begin : rows_begin + 21]
    read_part(fingerpulp, excel_name, status, 'fingerpulp')
    read_part(fingertip, excel_name, status, 'fingertip')
    read_part(fingerside, excel_name, status, 'fingerside')
    read_part(knuckle, excel_name, status, 'knuckle')
    read_part(fingernail, excel_name, status, 'fingernail')

def read_excel(excel_name):
    tags = excel_name.split('.')
    if (len(tags) != 2 or tags[1] != 'xlsx'):
        return
    workbook = xlrd.open_workbook(excel_name)
    sheet = workbook.sheet_by_index(0)

    read_status(sheet, tags[0], 'fist', 42)
    read_status(sheet, tags[0], 'relax', 68)

if __name__ == '__main__':
    directory = os.listdir('.')

    for file in directory:
        path = os.path.join('.', file)
        if os.path.isfile(path):
            read_excel(file)
