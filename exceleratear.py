"""excelerateAR for Idealist v0.1 by Karl Johnson.

Usage (Unix): python3 exceleratear.py some-list-of-orgs.txt
Usage (PC): python -m exceleratear some-list-of-orgs.txt

where some-list-of-orgs is a path to a text file that contains
a list of Idealist orgs (exact matches to names on the database),
delimited by new lines.

Note that in order to query Idealist v7's PostgreSQL database,
psycopg2 expects credentials as an environment variable I7DB_CREDS
with value as string with the following schema:

"dbname='' user='' host='' password=''"

with credential values in between single quotes.

"""

import datetime
import json
import os
import re
import sys

import psycopg2

from termcolor import colored

import xlsxwriter

# Instantiate global variables
user_name = None
results_file = None
excluded_orgs_file = None
included_orgs_file = None
data_dir = None
docs_dir = None
right_now = None
app_data = []
orgs_balances = []

keys = ["index",
        "invoice_num",
        "invoice_link",
        "description",
        "posted_by",
        "org_name",
        "posted_date",
        "due_date",
        "days_overdue",
        "amount_due"]

arial = {"font_name": "Arial Narrow"}
center = {"align": "center"}
big = {"font_size": 24}
medium = {"font_size": 18}
bold = {"bold": True}
italic = {"italic": True}
underline = {"underline": 1}
blue = {"font_color": "blue"}
gray = {"font_color": "gray"}
white = {"font_color": "white"}
black_bg = {"bg_color": "black"}
gray_bg = {"bg_color": "gray"}
money = {"num_format": 44}

headers = ["Item No.",
           "Invoice No.",
           "Posting Title",
           "Posted By",
           "Posted Date",
           "Due Date",
           "Days Overdue",
           "Amount Due"]

maxcol = 7

more_info = "For more information, call (646) 786-6875, write "
more_info += "accountsreceivable@idealist.org, or go to your "
more_info += "organization's dashboard on Idealist."


def main():
    """Run the higher-level tasks."""
    try:
        # Database credentials are stored in environment variable I7DB_CREDS
        connect_str = os.environ["I7DB_CREDS"]
    except:
        print(colored("ERROR: " +
                      "Need i7 database credentials exported as value " +
                      "for I7DB_CREDS :(",
                      "red"))
        exit(1)

    try:
        source_file = sys.argv[1]
    except:
        print(colored("ERROR: " +
                      "Need source file as command line argument :(",
                      "red"))
        exit(1)

    # Fun li'l header
    platform = sys.platform
    os.system('cls' if platform == 'nt' else 'clear')
    print(colored("{:=^79s}".format(
        "excelerateAR for Idealist, v0.1 by Karl Johnson"),
        "white"))

    # Get the user name that will be printed in the footer of the Excel file
    global user_name
    global right_now
    global data_dir
    global docs_dir
    global results_file
    global excluded_orgs_file
    global included_orgs_file

    global app_data
    global orgs_balances

    print()

    while True:
        user_name = input(colored("Your name: ", "white"))
        if user_name:
            break
    print()

    # Create a new subfolder for all the resulting files
    right_now = datetime.datetime.now()

    reports_dir = "reports"
    if not os.path.isdir(reports_dir):
        os.mkdir(reports_dir)

    data_dir_string = "{} {}".format(
        user_name, right_now.strftime("%Y %m%d %H%M%S"))
    data_dir = os.path.join(reports_dir, data_dir_string)
    os.mkdir(data_dir)

    docs_dir = os.path.join(data_dir, "docs")
    os.mkdir(docs_dir)

    logs_dir = os.path.join(data_dir, "logs")
    os.mkdir(logs_dir)

    results_path = os.path.join(logs_dir, "data.json")
    results_file = open(results_path, "a")

    balances_path = os.path.join(logs_dir, "balances.json")
    balances_file = open(balances_path, "a")

    included_orgs_path = os.path.join(logs_dir, "included.txt")
    included_orgs_file = open(included_orgs_path, "a")

    excluded_orgs_path = os.path.join(logs_dir, "excluded.txt")
    excluded_orgs_file = open(excluded_orgs_path, "a")

    print(colored("{:.<10s}".format("Running"), "yellow"))
    print()

    excluded_count = 0
    completed_count = 0

    for line in open(source_file):
        line = line.rstrip()

        status = prepare_data(line, connect_str)
        if status == 1:
            exit(1)
        elif status == 2:
            print(colored(
                "...EXCLUDED: {} - No rows returned"
                .format(line), "yellow"))
            excluded_orgs_file.write(line + "\n")
            excluded_count += 1
        elif status == 3:
            print(colored(
                "...EXCLUDED: {} - Couldn't make Excel file"
                .format(line), "yellow"))
            excluded_orgs_file.write(line + "\n")
            excluded_count += 1
        else:
            print(colored("...OK: {}".format(line), "cyan"))
            included_orgs_file.write(line + "\n")
            completed_count += 1

    results_file.write(json.dumps(app_data, indent=4))
    balances_file.write(json.dumps(orgs_balances, indent=4))

    print()
    print(colored("Process completed with {} completions and {} exclusions."
                  .format(completed_count, excluded_count), "green"))
    print()


def prepare_data(user_orgname, creds):
    """Pull the data from the database and put it in data structures."""
    try:
        conn = psycopg2.connect(creds)
        cursor = conn.cursor()

        # Get the data
        cursor.execute("""
            SELECT row_number() over(ORDER BY i.number ASC),
                i.number AS invoice_num,
                'https://www.idealist.org/invoices/' || i.id
                    AS invoice_link,
                li.description AS description,
                u.first_name || ' ' || u.last_name AS posted_by,
                o.name AS org_name,
                i.created::date AS posted_date,
                (i.created + INTERVAL '45 days')::date AS due_date,
                -- TODO: find where I got the below function and cite
                EXTRACT(EPOCH FROM(SELECT(NOW() -
                    (i.created + INTERVAL '45 days')))/86400)::int
                    AS days_overdue,
                li.unit_price AS amount_due
            FROM invoices AS i
            LEFT JOIN users AS u ON u.id = i.creator_id
            LEFT JOIN orgs AS o ON o.id = i.org_id
            LEFT JOIN line_items AS li ON li.invoice_id = i.id
            WHERE o.name = %s
            AND i.payment_settled = FALSE
            AND ((li.item_type = 'JOB') OR (li.item_type = 'INTERNSHIP'))
            AND EXTRACT(EPOCH FROM(SELECT(NOW() -
                (i.created + INTERVAL '45 days')))/86400)::int > 0
            ORDER BY i.number ASC;""", (user_orgname,))
        rows = cursor.fetchall()
    except:
        print(colored("ERROR: \
                       Something is wrong with your SQL, bro :(",
                      "red"))
        return 1

    if not rows:
        cursor.close()
        conn.close()
        return 2

    cursor.close()
    conn.close()

    i7_orgname = rows[0][5]

    # To track balance for each org
    org_balance = 0.0

    results = [dict(zip(keys, values)) for values in rows]
    for item in results:
        item["description"] = re.search(
            r"\"(.+)\"", item["description"]).group()[1:-1]
        item["amount_due"] = float(item["amount_due"])
        item["posted_date"] = item["posted_date"].strftime('%b %m, %Y')
        item["due_date"] = item["due_date"].strftime('%b %m, %Y')
        item["posted_by"] = item["posted_by"].title()
        org_balance += item["amount_due"]

    # Save the balance in the global list of org balances
    orgs_balances.append({item["org_name"]: org_balance})

    # Save the results dictionary to the global list of results
    app_data.append(results)

    if make_excel(results, i7_orgname) == 0:
        return 0
    else:
        return 3


def make_excel(data, i7_orgname):
    """Wrangle the data into an Excel file and a collection of PDFs."""
    i7_orgname = re.sub("\/+", "-", i7_orgname)
    i7_orgname = re.sub("\:+", "-", i7_orgname)

    # CREATE THE EXCEL FILE
    org_dir = os.path.join(docs_dir, i7_orgname)
    if not os.path.isdir(org_dir):
        os.mkdir(org_dir)
    filename_string = "AWB Invoices - {} - {}.xlsx".format(
        i7_orgname, right_now.strftime("%b %d %Y"))
    filename = os.path.join(org_dir, filename_string)

    wb = xlsxwriter.Workbook(filename)
    ws = wb.add_worksheet("Overdue Invoices")

    # SET UP THE FORMATS TO BE USED
    # flake8lint is throwing syntax errors on the below... hmmm...
    text_format = wb.add_format(
        dict(arial))

    title_format = wb.add_format(
        dict(arial, **center, **big, **bold, **white, **black_bg))

    subtitle_format = wb.add_format(
        dict(arial, **center, **medium, **white, **black_bg))

    header_format = wb.add_format(
        dict(arial, **italic, **white, **gray_bg))

    bold_format = wb.add_format(
        dict(arial, **bold))

    center_format = wb.add_format(
        dict(arial, **center))

    url_format = wb.add_format(
        dict(arial, **center, **underline, **blue))

    money_format = wb.add_format(
        dict(arial, **money))

    total_format = wb.add_format(
        dict(arial, **bold, **money))

    footer_format = wb.add_format(
        dict(arial, **center, **italic))

    # Write the title (org name)
    row = 0
    col = 0
    ws.merge_range(row, col, row, maxcol, i7_orgname, title_format)

    # Write the subtitle
    row += 1
    col = 0
    ws.merge_range(row, col, row, maxcol, "Overdue Idealist invoices as of {}"
                   .format(right_now.strftime("%b %d, %Y")), subtitle_format)

    # Write the header row
    row += 1
    col = 0
    col_widths = []

    for header in headers:
        width = len(str(header)) + 3
        ws.set_column(col, col, width)
        ws.write(row, col, header, header_format)
        col_widths.append(width)
        col += 1

    # WRITE THE TABLE
    row += 1
    col = 0

    for item in data:
        for key, value in item.items():
            # Don't do anything with the "invoice_link" value
            if key == "invoice_link" or key == "org_name":
                continue

            # Reset the column width if data value is wider than column header
            width = len(str(value)) + 3
            if width > col_widths[col]:
                col_widths[col] = width

            # Print invoice number as a hyperlink to the invoice on Idealist
            if key == "invoice_num":
                ws.write_url(row, col, item["invoice_link"],
                             url_format, str(item["invoice_num"]))

            # Center the index, invoice, and
            elif key == "index" or key == "days_overdue":
                ws.write_number(row, col, value, center_format)

            # Print the amount due as a number
            elif key == "amount_due":
                ws.write_number(row, col, value, money_format)
                # total_due += value

            # Print all other values as normal text
            else:
                ws.write(row, col, str(value), text_format)
            col += 1
        row += 1
        col = 0

    # This is an ersatz block to automatically adjust the width
    col = 0
    for width in col_widths:
        ws.set_column(col, col, width)
        col += 1

    # write the total
    ws.write(row, 6, "Total:", bold_format)
    ws.write_formula(row, 7, "=SUM(H4:H{})".format(row), total_format)

    # Add the footer
    row += 2
    col = 0
    ws.merge_range(row, col, row, maxcol, "Report run by {} at {} ET.".format(
        user_name, right_now.strftime("%Y.%m.%d %I:%M:%S%p")), footer_format)
    ws.merge_range(row + 1, col, row + 1, maxcol, more_info, footer_format)

    # Hide the gridlines for display
    ws.hide_gridlines(2)

    # Close the workbook
    wb.close()

    # All is good!
    return 0

if __name__ == "__main__":
    main()
