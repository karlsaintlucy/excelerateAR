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
import sys

from helpers import *

# from sqlalchemy import create_engine


def main():
    """Execute the core functionality."""
    counts = {"included": 0, "excluded": 0}

    # Get app-specific data
    options = get_options()
    show_interface_header()

    # Get user-specific data and build file structures.
    options = prepare_files(options)
    show_running()

    # Run through the orgs, log, and make Excel sheets.
    app_data, counts = step_through_orgs(options, counts)
    end_app(options, app_data)

    # Display how many orgs were included and excluded. That's it!
    print_app_result(counts)


def get_options():
    """Execute the application."""
    return {
        "orglist": open(sys.argv[1]),
        "creds": get_db_credentials(),
        "prefs": get_preferences()
    }


def prepare_files(options):
    """Get user data and prepare file structure."""
    options["user_info"] = get_user_info()

    # Build file structure
    options["right_now"] = datetime.datetime.now()
    options["dirs"] = make_dirs(options)
    options["logs"] = make_log_files(options["dirs"]["logs_dir"])

    # Connect to database and add pointers to options dict
    options["conn"], options["cursor"] = connect_to_db(options["creds"])

    # Snag the external SQL query and load into options dict
    query_file = open("query.sql")
    options["query"] = query_file.read()

    return options


def end_app(options, app_data):
    """Disconnect from database and make logs."""
    disconnect_from_db(options["conn"], options["cursor"])
    make_app_data_log(app_data, options["logs"])


if __name__ == "__main__":
    main()
