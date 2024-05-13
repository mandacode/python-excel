import csv
import argparse

import xlsxwriter


OPTIONS = {'strings_to_numbers': True}
BORDER = 1
BG_COLOR = "#fa7346"
NUM_FORMAT = "$ #,##0.00"
COLUMN_WIDTH = 20


def main():

    parser = argparse.ArgumentParser()
    parser.add_argument("input", help="source csv file path")
    parser.add_argument("output", help="destination xlsx file path")
    args = parser.parse_args()

    workbook = xlsxwriter.Workbook(args.output, options=OPTIONS)
    cell_format = workbook.add_format({"border": BORDER})
    header_format = workbook.add_format({"border": BORDER, "bold": True, "bg_color": BG_COLOR})
    money_format = workbook.add_format({"border": BORDER, "num_format": NUM_FORMAT, "align": "left"})
    total_format = workbook.add_format({"border": BORDER, "bold": True, "align": "left", "num_format": NUM_FORMAT})

    worksheet = workbook.add_worksheet()
    worksheet.set_column("A:C", COLUMN_WIDTH)

    with open(args.input, "r") as csv_file:
        csv_reader = csv.DictReader(csv_file)

        headers = ["Date", "Description", "Amount"]
        worksheet.write_row("A1", headers, header_format)

        row = 1
        for record in csv_reader:
            worksheet.write(row, 0, record["Date"], cell_format)
            worksheet.write(row, 1, record["Description"], cell_format)
            worksheet.write(row, 2, record["Amount"], money_format)
            row += 1

        worksheet.write(row, 1, "Total:", header_format)
        worksheet.write(row, 2, f"=SUM(C2:C{row-1})", total_format)

    workbook.close()


if __name__ == "__main__":
    main()
