import csv

import xlsxwriter

CSV_FILE = "bank_statement.csv"
SHEET_NAME = "summary.xlsx"
A_COLUMN_WIDTH = 10
B_COLUMN_WIDTH = 20
C_COLUMN_WIDTH = 10
BORDER = 1
DARKER_BG_COLOR = '#a6c1ed'
LIGHTER_BG_COLOR = '#e6eefa'
WORKBOOK_OPTIONS = {'strings_to_numbers': True}
NUM_FORMAT = '$ #,##'

# TODO 1. create basic version
# TODO 2. add formatting with good practices


def main():
    workbook = xlsxwriter.Workbook(SHEET_NAME, WORKBOOK_OPTIONS)
    worksheet = workbook.add_worksheet()
    worksheet.set_column('A:A', A_COLUMN_WIDTH)
    worksheet.set_column('B:B', B_COLUMN_WIDTH)
    worksheet.set_column('C:C', C_COLUMN_WIDTH)

    headers_format = workbook.add_format({'bold': True, 'border': BORDER, 'bg_color': DARKER_BG_COLOR})
    total_format = workbook.add_format({'bold': True, 'border': BORDER, 'num_format': NUM_FORMAT})
    cell_format = workbook.add_format({'border': BORDER})

    headers = ["Date", "Description", "Amount"]
    worksheet.write_row('A1', headers, headers_format)

    with open(CSV_FILE, "r") as csv_file:
        csv_reader = csv.reader(csv_file)
        next(csv_reader)

        row_counter = 1
        for row_idx, row in enumerate(csv_reader, start=row_counter):
            worksheet.write(row_idx, 0, row[1], cell_format)
            worksheet.write(row_idx, 1, row[6], cell_format)
            worksheet.write(row_idx, 2, row[7], cell_format)
            row_counter += 1

    worksheet.write(row_counter, 1, "Total:", headers_format)
    worksheet.write(row_counter, 2, f"=SUM(C2:C{row_counter - 1})", total_format)

    workbook.close()


if __name__ == "__main__":
    main()
