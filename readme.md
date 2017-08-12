# excelerateAR for Idealist
August 12, 2017 by Karl Johnson

## Usage

Mac: `python3 exceleratear.py path-to-orglist`
PC: `python -m exceleratear path-to-orglist`

“path-to-orglist” corresponds to a text file with org names (exact matches with name string on Idealist) delimited by newlines.

NB: In order to query the i7 database, excelerateAR expects an environment variable with the following schema:

`I7DB_CRED="dbname=**dbname** user=**username** host=**hostname** password=**password**”`
(where values in [] are surrounded by single quotes)

The username is collected via terminal input; this helps name the directory where files are stored and which name is printed at the bottom of each Excel sheet.


## Dependencies
**Built-in:**
- re
- os
- sys
- json
- datetime

**Third Party:**
- psycopg2
- termcolor
- xlsxwriter

All third-party modules can be installed with `pip`.


## Behavior
excelerateAR collects username, then creates a `reports` directory in the same folder as the application, with `docs` and `logs` subfolders. It then reads org names from the source file one line at a time, querying the i7 PostgreSQL database and creating an Excel (2007+) file with the returned data for each organization in the list in its own folder.


The terminal displays color-coded status for each line item (cyan for “included”, yellow for “excluded”), logs each line-item as a JSON object in `data.json`, and writes the org name in either the `included.txt` or `excluded.txt` log file. Problematic characters (i.e., “/“ and “:”) in org names are replaced with “-“ when passed to `xlsxwriter`. Additionally, each organization's name and balance is logged in JSON in `balances.json`.


xlsxwriter lists each outstanding line item, its invoice number (formatted as an hyperlink to its public invoice on Idealist), its corresponding job title, the name of the person who posted the job, the posting date, due date, quantity of days overdue, and amount owed. At the bottom, a sum function calculates the grand total owed by the organization. The name entered as username at application execution is printed in the bottom section of the Excel file, along with the date, time, and “more info.”