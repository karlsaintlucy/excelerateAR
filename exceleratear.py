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
import os
import sys

from helpers import (
    get_preferences, show_interface_header, get_user_info,
    make_dirs, make_log_files, connect_to_db,
    show_interface_running, excelerate_orgs, disconnect_from_db,
    make_app_data_log, print_app_result
)


def main():
    """Execute the core functionality."""
    counts = {"included": 0, "excluded": 0}

    options = {"orglist": open(sys.argv[1]),
               "creds": os.environ.get("I7DB_CREDS"),
               "prefs": get_preferences()}

    show_interface_header()

    options["user_info"] = get_user_info()
    options["right_now"] = datetime.datetime.now()
    options["dirs"] = make_dirs(options)
    options["logs"] = make_log_files(options["dirs"]["logs_dir"])
    options["conn"], options["cursor"] = connect_to_db(options["creds"])
    query_file = open("query.sql")
    options["query"] = query_file.read()

    show_interface_running()

    app_data, counts = excelerate_orgs(options, counts)
    disconnect_from_db(options["conn"], options["cursor"])
    make_app_data_log(app_data, options["logs"])

    print_app_result(counts)


if __name__ == "__main__":
    main()
