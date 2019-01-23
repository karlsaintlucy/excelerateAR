"""Helpers for excelerateAR for Idealist."""

import json
import os
import re
import sys

import psycopg2
from termcolor import colored

from make_excel import make_excel


def get_preferences():
    """Get the application preferences."""
    prefs_file = open("preferences.json")
    prefs = json.load(prefs_file)
    return prefs


def show_interface_header():
    """Clear the screen and show the interface header."""
    platform = sys.platform
    os.system('cls' if (platform == 'nt' or platform == 'win32') else 'clear')
    print(colored("{:=^79s}".format("excelerateAR for Idealist, "
                                    "v0.1 by Karl Saint Lucy"), "white"))


def get_user_info():
    """Get user's name, email, and callback phone number."""
    user_name = input(colored("Your name: ", "white"))
    user_email = input(colored("Write back email: ", "white"))
    user_phone = input(colored("Callback phone: ", "white"))
    more_info = ("For more information, call {}, write {}, or go to your "
                 "organization's dashboard on Idealist."
                 .format(user_phone, user_email))
    return {"name": user_name,
            "email": user_email,
            "phone": user_phone,
            "more_info": more_info}


def make_dirs(options):
    """Create the folder structure to house the resulting documents."""
    right_now = options["right_now"]
    user_info = options["user_info"]
    prefs = options["prefs"]

    reports_dir = "reports"
    if not os.path.isdir(reports_dir):
        os.mkdir(reports_dir)

    user_name = user_info["name"].replace(" ", "_")
    folder_date_format = prefs["folder_date_format"]

    data_dir_string = "{}_{}".format(
        right_now.strftime(folder_date_format), user_name)
    data_dir = os.path.join(reports_dir, data_dir_string)
    os.mkdir(data_dir)

    docs_dir = os.path.join(data_dir, "docs")
    os.mkdir(docs_dir)

    logs_dir = os.path.join(data_dir, "logs")
    os.mkdir(logs_dir)

    return {"data_dir": data_dir,
            "docs_dir": docs_dir,
            "logs_dir": logs_dir}


def make_log_files(logs_dir):
    """Create log files."""
    results_path = os.path.join(logs_dir, "data.json")
    results_file = open(results_path, "a")

    included_orgs_path = os.path.join(logs_dir, "included.txt")
    included_orgs_file = open(included_orgs_path, "a")

    excluded_orgs_path = os.path.join(logs_dir, "excluded.txt")
    excluded_orgs_file = open(excluded_orgs_path, "a")

    return {"results_file": results_file,
            "included_orgs_file": included_orgs_file,
            "excluded_orgs_file": excluded_orgs_file}


def connect_to_db(creds):
    """Open connection with i7 database."""
    conn = psycopg2.connect(creds)
    cursor = conn.cursor()
    return conn, cursor


def show_interface_running():
    """Show 'Running...'."""
    print()
    print(colored("{:.<10s}".format("Running"), "yellow"))
    print()


def excelerate_orgs(options, counts):
    """Step through each org, making Excel file and logging results."""
    app_data = []
    orglist = options["orglist"]

    for orgname in orglist:
        orgname = orgname.rstrip()
        results, counts = get_org_invoices(options, orgname, counts)
        if not results:
            continue

        make_excel(orgname,
                   results,
                   options["user_info"],
                   options["prefs"],
                   options["right_now"],
                   options["dirs"]["docs_dir"])

        app_data.append(results)

    return app_data, counts


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
    return cursor.fetchall()


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
