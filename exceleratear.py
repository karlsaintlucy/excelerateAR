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


def main():
    """Execute the application."""
    options = {}

    # Snag the orglist, database creds, and preferences, and save to options
    orglist_path = sys.argv[1]
    options["orglist"] = open(orglist_path)
    options["creds"] = get_db_credentials()
    options["prefs"] = get_preferences()

    # Pretty header that shows the name of the app :)
    show_interface_header()

    # Get the client info (for folder naming and writing in the Excel files)
    options["user_info"] = get_user_info()

    # Build file structure
    options["right_now"] = get_right_now()
    options["dirs"] = make_dirs(options)
    options["logs"] = make_log_files(options["dirs"]["logs_dir"])

    # Connect to database and add pointers to options dict
    options["conn"], options["cursor"] = connect_to_db(options["creds"])

    # Pretty little 'Running...' indicator :)
    show_running()

    # Snag the external SQL query and load into options dict
    query_file = open("query.sql")
    options["query"] = query_file.read()

    # Instantiate the counters and run through the orgs
    counts = {"included": 0, "excluded": 0}
    app_data = step_through_orgs(options, counts)

    # Disconnect, make logs, show the results
    disconnect_from_db(options["conn"], options["cursor"])
    make_app_data_log(app_data, options["logs"])

    # Display how many orgs were included and excluded. That's it!
    print_app_result(counts)


def get_db_credentials():
    """Get the i7 database credentials."""
    creds = os.environ["I7DB_CREDS"]
    return creds


def get_preferences():
    """Get the application preferences."""
    prefs_file = open("preferences.json")
    prefs = json.load(prefs_file)
    return prefs


def show_interface_header():
    """Clear the screen and show the interface header."""
    platform = sys.platform
    os.system('cls' if platform == 'nt' else 'clear')
    print(colored("{:=^79s}".format(
        "excelerateAR for Idealist, v0.1 by Karl Johnson"),
        "white"))


def get_user_info():
    """Get user's name, email, and callback phone number."""
    user_name = input(colored("Your name: ", "white"))
    user_email = input(colored("Write back email: ", "white"))
    user_phone = input(colored("Callback phone: ", "white"))
    more_info = "For more information, call {}, ".format(user_phone)
    more_info += "write {}, or ".format(user_email)
    more_info += "go to your organization's dashboard on Idealist."
    user_info = {
        "name": user_name,
        "email": user_email,
        "phone": user_phone,
        "more_info": more_info}
    return user_info


def get_right_now():
    """Get the current date and time."""
    right_now = datetime.datetime.now()
    return right_now


def make_dirs(options):
    """Create the folder structure to house the resulting documents."""
    right_now = options["right_now"]
    user_info = options["user_info"]
    prefs = options["prefs"]

    reports_dir = "reports"
    if not os.path.isdir(reports_dir):
        os.mkdir(reports_dir)

    user_name = user_info["name"]
    folder_date_format = prefs["folder_date_format"]

    data_dir_string = "{} {}".format(
        user_name, right_now.strftime(folder_date_format))
    data_dir = os.path.join(reports_dir, data_dir_string)
    os.mkdir(data_dir)

    docs_dir = os.path.join(data_dir, "docs")
    os.mkdir(docs_dir)

    logs_dir = os.path.join(data_dir, "logs")
    os.mkdir(logs_dir)

    dirs = {
        "data_dir": data_dir,
        "docs_dir": docs_dir,
        "logs_dir": logs_dir}
    return dirs


def make_log_files(logs_dir):
    """Create log files."""
    results_path = os.path.join(logs_dir, "data.json")
    results_file = open(results_path, "a")

    included_orgs_path = os.path.join(logs_dir, "included.txt")
    included_orgs_file = open(included_orgs_path, "a")

    excluded_orgs_path = os.path.join(logs_dir, "excluded.txt")
    excluded_orgs_file = open(excluded_orgs_path, "a")

    logs = {
        "results_file": results_file,
        "included_orgs_file": included_orgs_file,
        "excluded_orgs_file": excluded_orgs_file}
    return logs


def connect_to_db(creds):
    """Open connection with i7 database."""
    conn = psycopg2.connect(creds)
    cursor = conn.cursor()
    return conn, cursor


def show_running():
    """Show 'Running...'."""
    print()
    print(colored("{:.<10s}".format("Running"), "yellow"))
    print()


def step_through_orgs(options, counts):
    """Step through each org, making Excel file and logging results."""
    app_data = []

    orglist = options["orglist"]

    for orgname in orglist:
        orgname = orgname.rstrip()
        results, counts = get_org_invoices(options, orgname, counts)
        if not results:
            continue

        app_data = log_results(results, app_data)
        make_excel(orgname,
                   results,
                   options["user_info"],
                   options["prefs"],
                   options["right_now"],
                   options["dirs"]["docs_dir"])

    return app_data


def get_org_invoices(options, orgname, counts):
    """Read each orgname and run the queries."""
    included_orgs_file = options["logs"]["included_orgs_file"]
    excluded_orgs_file = options["logs"]["excluded_orgs_file"]
    keys = options["prefs"]["keys"]

    rows = run_query(options["cursor"], options["query"], orgname)

    if rows:
        included_orgs_file.write(orgname + "\n")
        counts["included"] += 1
        print(colored("...OK: {}".format(orgname), "cyan"))
        results = [dict(zip(keys, values)) for values in rows]
        results = sanitize_results(results, options["prefs"])
    else:
        excluded_orgs_file.write(orgname + "\n")
        counts["excluded"] += 1
        print(colored("...EXCLUDED: {}".format(orgname), "yellow"))
        results = None

    return results, counts


def run_query(cursor, query, orgname):
    """Run the query against the database with the orgname."""
    cursor.execute(query, (orgname,))
    rows = cursor.fetchall()
    return rows


def sanitize_results(results, prefs):
    """Format the query results according to preferences."""
    invoice_date_format = prefs["invoice_date_format"]
    for item in results:
        # Thanks, Antoine, for help with the below!
        description_result = re.search(r"\"(.+)\"", item["description"])
        description_group = description_result.group()
        item["description"] = description_group[1:-1]
        item["amount_due"] = float(item["amount_due"])
        item["posted_date"] = item["posted_date"].strftime(
            invoice_date_format)
        item["due_date"] = item["due_date"].strftime(
            invoice_date_format)
        item["posted_by"] = item["posted_by"].title()

    return results


def log_results(results, app_data):
    """Append each org's db results to a list of results."""
    app_data.append(results)
    return app_data


def make_excel(orgname, results, user_info, prefs, right_now, docs_dir):
    """Wrangle the data into an Excel file and a collection of PDFs."""
    orgname = re.sub("\/+", "-", orgname)
    orgname = re.sub("\:+", "-", orgname)

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

    # There's got to be a way to simplify this block:
    text_format = wb.add_format(prefs["text_format"])
    title_format = wb.add_format(prefs["title_format"])
    subtitle_format = wb.add_format(prefs["subtitle_format"])
    header_format = wb.add_format(prefs["header_format"])
    bold_format = wb.add_format(prefs["bold_format"])
    center_format = wb.add_format(prefs["center_format"])
    url_format = wb.add_format(prefs["url_format"])
    money_format = wb.add_format(prefs["money_format"])
    total_format = wb.add_format(prefs["total_format"])
    footer_format = wb.add_format(prefs["footer_format"])

    maxcol = 7

    # Write the title (org name)
    row = 0
    col = 0
    ws.merge_range(row, col, row, maxcol, orgname, title_format)

    # Write the subtitle
    row += 1
    col = 0
    ws.merge_range(row, col, row, maxcol,
                   "Overdue Idealist invoices as of {}"
                   .format(
                       right_now.strftime(invoice_date_format)),
                   subtitle_format)

    # Write the header row.
    row += 1
    col = 0
    col_widths = []

    for header in headers:
        width = len(str(header)) + 3
        ws.set_column(col, col, width)
        ws.write(row, col, header, header_format)
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
                             url_format, str(item["invoice_num"]))

            # Center the index, invoice, and
            elif key == "index" or key == "days_overdue":
                ws.write_number(row, col, value, center_format)

            # Print the amount due as a number
            elif key == "amount_due":
                ws.write_number(row, col, value, money_format)

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
    sum_function = "=SUM(H4:H{})".format(row)
    ws.write_formula(row, 7, sum_function, total_format)

    # Add the footer
    row += 2
    col = 0
    excel_time = right_now.strftime(excel_date_format)
    ws.merge_range(row, col, row, maxcol,
                   "Report run by {} at {} ET.".format(
                       user_name, excel_time),
                   footer_format)
    ws.merge_range(row + 1, col, row + 1, maxcol, more_info, footer_format)

    # Hide the gridlines for display (argument "2" specifies hide all)
    ws.hide_gridlines(2)

    # Close the workbook
    wb.close()


def disconnect_from_db(conn, cursor):
    """Close connection with i7 database."""
    cursor.close()
    conn.close()


def make_app_data_log(app_data, logs):
    """Take the list of result dicts and log them as JSON in data.json."""
    results_file = logs["results_file"]
    results_file.write(json.dumps(app_data, indent=4))


def print_app_result(counts):
    """Print how many orgs were included and excluded."""
    print()
    print(colored("Process completed with {} inclusions and {} exclusions."
                  .format(counts["included"], counts["excluded"]), "green"))
    print()


if __name__ == "__main__":
    main()
