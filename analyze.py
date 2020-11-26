import openpyxl
from openpyxl.utils.cell import column_index_from_string

import pdb

MAX_ROWS=400

wb = openpyxl.load_workbook("./responses.xlsx")

for name in wb.sheetnames:
    sheet = wb[name]

    # For some reason max_row is coming out large value than actual entries
    # max_row = sheet.max_row

    results = dict()

    # 1) What did you like best about SKY schools?
    col = column_index_from_string('D')
    query = sheet.cell(row=1, column=col)

    # Each row can have multiple answers, such as [Breathing, Yoga, All of it]
    # Each answer must be counted.
    total_votes = 0
    data = dict()
    for i in range(2, MAX_ROWS):
        val = sheet.cell(row=i, column=col).value

        if val is None:
            # no more entries?
            break

        val = val.split(',')

        for v in val:
            data[v] = 1 + data.get(v, 0)
            total_votes += 1

    print(data, total_votes)


    # 2) Do you use what you learned in SKY schools?
    col = column_index_from_string('E')
    query = sheet.cell(row=1, column=col)

    # Each row can have 'Sometimes', 'Every day', 'Never' or some sentence which
    # is counted as 'Other'
    total_votes = 0
    data = dict()
    for i in range(2, MAX_ROWS):
        val = sheet.cell(row=i, column=col).value

        if val is None:
            # no more entries?
            break

        if val in ['Sometimes', 'Every day', 'Never']:
            data[val] = 1 + data.get(val, 0)
        else:
            data['Other'] = 1 + data.get('Other', 0)

        total_votes += 1

    print(data, total_votes)


    # 3) After SKY Schools do you feel: More focused, More calm and so on?
    for c in ['F', 'G', 'H', 'I', 'J', 'K', 'L']:
        col = column_index_from_string(c)
        query = sheet.cell(row=1, column=col)

        # Each row can only have 'Yes', 'No', 'A little bit'
        total_votes = 0
        data = dict()
        for i in range(2, MAX_ROWS):
            val = sheet.cell(row=i, column=col).value

            if val is None:
                # no more entries?
                break

            data[val] = 1 + data.get(val, 0)
            total_votes += 1

        print(data, total_votes)

    
    # For each school, create corresponding analysis sheet
    # wb.create_sheet(title="Analysis " + sheet[:10])

# print(wb.sheetnames)
