"""Take each org's results and write it into a formatted Excel file."""
import json
import os
import re

import xlsxwriter


def make_excel(orgname, results, user_info, prefs, right_now, docs_dir):
    """Wrangle the data into an Excel file and a collection of PDFs."""
    orgname = re.sub("\/+", "-", orgname)
    orgname = re.sub("\:+", "-", orgname)

    # H/T to Greg Sadetsky:
    static_formats_file = open("formats.json")
    static_formats = json.load(static_formats_file)

    user_name = user_info["name"]
    more_info = user_info["more_info"]

    excel_date_format = prefs["excel_date_format"]
    folder_date_format = prefs["folder_date_format"]
    invoice_date_format = prefs["invoice_date_format"]

    filename_string = "AWB Invoices - {} - {}.xlsx".format(
        orgname, right_now.strftime(folder_date_format))
    filename = os.path.join(docs_dir, filename_string)

    wb = xlsxwriter.Workbook(filename)
    ws = wb.add_worksheet("Overdue Invoices")

    headers = prefs["headers"]

    file_formats = {format_name: wb.add_format(static_formats[format_name])
                    for format_name, description in static_formats.items()}

    maxcol = 7

    # Write the title (org name)
    row = 0
    col = 0
    ws.merge_range(row, col, row, maxcol, orgname, file_formats["title"])

    # Write the subtitle
    row += 1
    col = 0
    ws.merge_range(row, col, row, maxcol,
                   "Overdue Idealist invoices as of {}"
                   .format(right_now.strftime(invoice_date_format)),
                   file_formats["subtitle"])

    # Write the header row.
    row += 1
    col = 0
    col_widths = []

    for header in headers:
        width = len(str(header)) + 3
        ws.set_column(col, col, width)
        ws.write(row, col, header, file_formats["header"])
        col_widths.append(width)
        col += 1

    # Write the table.
    row += 1
    col = 0

    for item in results:
        # Skip invoices with "None" as invoice number
        if item["invoice_num"] is None:
            continue

        for key, value in item.items():
            # Don't do anything with the "invoice_link" value
            if key == "invoice_link" or key == "org_name":
                continue

            # Reset the column width if data value is wider than col header
            width = len(str(value)) + 3
            if width > col_widths[col]:
                col_widths[col] = width

            # Print invoice number as a hyperlink to the invoice
            if key == "invoice_num":
                ws.write_url(row, col, item["invoice_link"],
                             file_formats["url"], str(item["invoice_num"]))

            # Center the index, invoice, and
            elif key == "index" or key == "days_overdue":
                ws.write_number(row, col, value, file_formats["center"])

            # Print the amount due as a number
            elif key == "amount_due":
                ws.write_number(row, col, value, file_formats["money"])

            # Print all other values as normal text
            else:
                ws.write(row, col, str(value), file_formats["text"])
            col += 1
        row += 1
        col = 0

    # This is an ersatz block to automatically adjust the width
    col = 0
    for width in col_widths:
        ws.set_column(col, col, width)
        col += 1

    # write the total
    ws.write(row, 6, "Total:", file_formats["bold"])
    sum_function = "=SUM(H4:H{})".format(row)
    ws.write_formula(row, 7, sum_function, file_formats["total"])

    # Add the footer
    row += 2
    col = 0
    excel_time = right_now.strftime(excel_date_format)
    ws.merge_range(row, col, row, maxcol,
                   "Report run by {} at {} ET.".format(
                       user_name, excel_time),
                   file_formats["footer"])
    ws.merge_range(row + 1, col, row + 1, maxcol,
                   more_info, file_formats["footer"])

    # Hide the gridlines for display (argument "2" specifies hide all)
    ws.hide_gridlines(2)

    # Close the workbook
    wb.close()
