import csv
import argparse

import xlsxwriter


def main():

    parser = argparse.ArgumentParser()
    parser.add_argument("input", help="source csv file path")
    parser.add_argument("output", help="destination xlsx file path")
    args = parser.parse_args()

    workbook = xlsxwriter.Workbook(args.output, options={'strings_to_numbers': True})
    worksheet = workbook.add_worksheet()

    with open(args.input, "r") as csv_file:
        csv_reader = csv.DictReader(csv_file)

        headers = ["Date", "Description", "Amount"]
        worksheet.write_row("A1", headers)

        row = 1
        for record in csv_reader:
            worksheet.write(row, 0, record["Date"])
            worksheet.write(row, 1, record["Description"])
            worksheet.write(row, 2, record["Amount"])
            row += 1

        worksheet.write(row, 1, "Total:")
        worksheet.write(row, 2, f"=SUM(C2:C{row-1})")

    workbook.close()


if __name__ == "__main__":
    main()
