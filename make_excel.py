"""Take each org's results and write it into a formatted Excel file."""
import json
import os
import re

import xlsxwriter


def make_excel(orgname, results, user_info, prefs, right_now, docs_dir):
    """Wrangle the data into an Excel file and a collection of PDFs."""
    orgname = re.sub("\/+", "-", orgname)
    orgname = re.sub("\:+", "-", orgname)

    static_formats_file = open("formats.json")
    static_formats = json.load(static_formats_file)

    user_name = user_info["name"]
    more_info = user_info["more_info"]
    excel_date_format = prefs["excel_date_format"]
    folder_date_format = prefs["folder_date_format"]
    invoice_date_format = prefs["invoice_date_format"]
    headers = prefs["headers"]

    filename_string = "AWB Invoices - {} - {}.xlsx".format(
        orgname, right_now.strftime(folder_date_format))
    filename = os.path.join(docs_dir, filename_string)
    wb = xlsxwriter.Workbook(filename)
    ws = wb.add_worksheet("Overdue Invoices")

    # H/T Greg Sadetsky:
    file_formats = {format_name: wb.add_format(static_formats[format_name])
                    for format_name, description in static_formats.items()}

    maxcol = 7

    row = 0
    col = 0
    ws.merge_range(row, col, row, maxcol, orgname, file_formats["title"])

    row += 1
    col = 0
    ws.merge_range(row, col, row, maxcol,
                   "Overdue Idealist invoices as of {}"
                   .format(right_now.strftime(invoice_date_format)),
                   file_formats["subtitle"])

    row += 1
    col = 0
    col_widths = []

    for header in headers:
        width = len(str(header)) + 3
        ws.set_column(col, col, width)
        ws.write(row, col, header, file_formats["header"])
        col_widths.append(width)
        col += 1

    row += 1
    col = 0

    for item in results:
        if item["invoice_num"] is None:
            continue

        adjust_width(ws, item, prefs, col_widths, col)
        ws.write_number(row, col, item["index"], file_formats["center"])
        col += 1

        adjust_width(ws, item, prefs, col_widths, col)
        ws.write_url(row, col, item["invoice_link"],
                     file_formats["url"], str(item["invoice_num"]))
        col += 1

        adjust_width(ws, item, prefs, col_widths, col)
        ws.write(row, col, item["description"], file_formats["text"])
        col += 1

        adjust_width(ws, item, prefs, col_widths, col)
        ws.write(row, col, item["posted_by"], file_formats["text"])
        col += 1

        adjust_width(ws, item, prefs, col_widths, col)
        ws.write(row, col, item["posted_date"], file_formats["text"])
        col += 1

        adjust_width(ws, item, prefs, col_widths, col)
        ws.write(row, col, item["due_date"], file_formats["text"]),
        col += 1

        adjust_width(ws, item, prefs, col_widths, col)
        ws.write_number(row, col, item["days_overdue"], file_formats["center"])
        col += 1

        adjust_width(ws, item, prefs, col_widths, col)
        ws.write_number(row, col, item["amount_due"], file_formats["money"])

        row += 1
        col = 0

    ws.write(row, 6, "Total:", file_formats["bold"])
    sum_function = "=SUM(H4:H{})".format(row)
    ws.write_formula(row, 7, sum_function, file_formats["total"])

    row += 2
    col = 0
    excel_time = right_now.strftime(excel_date_format)
    ws.merge_range(row, col, row, maxcol,
                   "Report run by {} at {} ET.".format(
                       user_name, excel_time),
                   file_formats["footer"])
    ws.merge_range(row + 1, col, row + 1, maxcol,
                   more_info, file_formats["footer"])

    ws.hide_gridlines(2)
    wb.close()


def adjust_width(ws, item, prefs, col_widths, col):
    """Make the column wider if the content of each item is wider."""
    this_col = prefs["keys"][col]
    if len(str(item[this_col])) > col_widths[col]:
        ws.set_column(col, col, len(str(item[this_col])) + 2)
