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
    options["right_now"] = datetime.datetime.now()
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


if __name__ == "__main__":
    main()
