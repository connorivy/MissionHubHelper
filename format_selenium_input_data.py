from openpyxl import Workbook
from openpyxl import load_workbook
from openpyxl.styles import Color, PatternFill, Font, Border
import time
import sys

def normalize_excel_sheet():
    # Give the location of the file 
    loc = ("./supporting_files/contacts.xlsx") 

    # To open Workbook 
    wb = load_workbook(filename = loc)
    ws = wb.active 

    # initialize variables
    white = PatternFill(start_color='FFFFFF',
                   end_color='FFFFFF',
                   fill_type='solid')

    redFill = PatternFill(start_color='FFFF0000',
                   end_color='FFFF0000',
                   fill_type='solid')

    lightredFill = PatternFill(start_color='ffcccb',
                   end_color='ffcccb',
                   fill_type='solid')

    
    # delete rows that don't have a phone number
    for row in ws:
        if not any(cell.value for cell in row):
            # print('empty row', row[0].row)
            ws.delete_rows(row[0].row, 1)

    # split full name into first and last
    row_count = ws.max_row

    for row in ws.iter_rows(min_row = 2, min_col = 1, max_col = 1, max_row = row_count):
        for cell in row:
            cell.fill = white
            name = cell.value
            try:
                name.strip()
            except:
                continue

            spaces = 0
            for char in name:
                if char == ' ':
                    spaces += 1

            # if no phone number
            if cell.offset(0,1).value == None:
                cell.fill = lightredFill
                cell.offset(0,1).fill = redFill
                cell.offset(0,2).fill = lightredFill
                cell.offset(0,3).fill = lightredFill
                cell.offset(0,4).fill = lightredFill
                cell.offset(0,5).fill = lightredFill
                cell.offset(0,6).fill = lightredFill

            if spaces == 0:
                cell.offset(0,3).value = name

            elif spaces == 1:
                first, last = name.split()
                cell.offset(0,3).value = first
                cell.offset(0,4).value = last

            elif spaces == 2:
                first, middle, last = name.split()
                cell.offset(0,3).value = first + " " + middle
                cell.offset(0,4).value = last

            else:
                cell.fill = redFill
                cell.offset(0,1).fill = lightredFill
                cell.offset(0,2).fill = lightredFill
                cell.offset(0,3).fill = lightredFill
                cell.offset(0,4).fill = lightredFill
                cell.offset(0,5).fill = lightredFill
                cell.offset(0,6).fill = lightredFill

    # format phone numbers to exclude special characters
    for row in ws.iter_rows(min_row = 2, min_col = 2, max_col = 2, max_row = row_count):
        for cell in row:
            number = cell.value
            try:
                number.strip()
            except:
                continue
            
            # remove anything that isn't a number from the string
            formatted_num = ''
            for char in number:
                if char.isdigit():
                    formatted_num += str(char)

            # if not 10 numbers, flag that cell
            if len(formatted_num) != 10:
                cell.fill = redFill
                cell.offset(0,-1).fill = lightredFill
                cell.offset(0,1).fill = lightredFill
                cell.offset(0,2).fill = lightredFill
                cell.offset(0,3).fill = lightredFill
                cell.offset(0,4).fill = lightredFill
                cell.offset(0,5).fill = lightredFill

            else:
                cell.offset(0,4).value = formatted_num

    for row in ws.iter_rows(min_row = 2, min_col = 3, max_col = 3, max_row = row_count):
        for cell in row:
            gender = str(cell.value)

        if gender.lower() == 'male' or gender.lower() == 'm' or gender.lower() == 'boy' or gender.lower() == 'guy':
            cell.offset(0,4).value = 'male'

        elif gender.lower() == 'female' or gender.lower() == 'f' or gender.lower() == 'girl' or gender.lower() == 'gal':
            cell.offset(0,4).value = 'female'

        # # flag cells that don't have a gender in them
        # elif cell.offset(0, -1).value != None and gender.lower() == 'none':
        #     cell.fill = redFill
        #     cell.offset(0,-1).fill = lightredFill
        #     cell.offset(0,-2).fill = lightredFill
        #     cell.offset(0,1).fill = lightredFill
        #     cell.offset(0,2).fill = lightredFill
        #     cell.offset(0,3).fill = lightredFill
        #     cell.offset(0,4).fill = lightredFill

        elif cell.offset(0, -1).value == None and gender.lower() == 'none':
            continue

        else:
            cell.offset(0,4).value = 'other'

    wb.save(filename = './supporting_files/contacts_formatted_do_not_edit.xlsx')

def get_contact_list():
    loc = ('./supporting_files/contacts_formatted_do_not_edit.xlsx') 

    # open Workbook 
    wb = load_workbook(filename = loc)
    ws = wb.active 

    row_count = ws.max_row

    all_contacts = []
    for row in ws.iter_rows(min_row = 2, min_col = 4, max_col = 7, max_row = row_count):
        single_contact_info = []
        for cell in row:

            # if the cell is red then there is a problem with that person's info, don't add it to the list
            if cell.fill.start_color.index == '00ffcccb':
                break
            single_contact_info.append(cell.value)
            if cell.column == 7 and single_contact_info != [None, None, None, None]:
                all_contacts.append(single_contact_info)
            
    return all_contacts

def find_labels():
    ask = False

    with open('./supporting_files/labels.txt') as f:
        labels = f.readlines()

    # remove whitespace characters like `\n` at the end of each line
    labels = [x.strip().lower() for x in labels]

    for x in labels:
        for char in range(0,4):
            if x[char].isdigit():
                continue
            else:
                if not ask:
                    ask = ask_to_follow_convention(x)

        if x[5:9].lower() == 'spri' or x[5:9].lower() == 'fall' or x[5:9].lower() == 'wint' or x[5:9].lower() == 'summ':
            continue
        else:
            if not ask:
                ask = ask_to_follow_convention(x)

    return labels

def ask_to_follow_convention(x):
    cont = 'change me'
    print('\n\n\n')
    while cont.lower() != 'y' and cont.lower() != 'n':
        cont = input('\nThe item "' + x + '" in the "labels.txt" file does NOT meet the established naming convention for labels (found in supporting_files/labels_convention). Would you like to continue anyways? (y/n)       ')
    if cont.lower() == 'n':
        print('\n\nThanks for keeping missionhub organized ;)')
        time.sleep(1)
        sys.exit()
    else:
        sure = 'change me'
        while sure.lower() != 'y' and sure.lower() != 'n':
            sure = input('\n\nThis may result in missionhub becoming unorganized. Are you sure you want to continue? (y/n)        ')
        if sure.lower() == 'n':
            print('\n\nThanks for keeping missionhub organized ;)')
            time.sleep(1)
            sys.exit()

    return True
