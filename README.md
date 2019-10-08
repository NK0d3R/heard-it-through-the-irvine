# Heard It Through the Irvine v001
### Basic data analytics for Irvine Company's apartment rental public data

A Python 2 script which pulls data from Irvine Company's public website,
analyses it and outputs an XLS file summarizing the data on available
apartments for rent for their apartment complexes.

#### Requirements:
* Python 2
* xlsxwriter for Python 2 (`pip install xlsxwriter`)
* A valid url.txt containing the IDs for the apartment building(s) you're
interested in - a sample one for RiverView in North San Jose is included,
you can get the IDs from their websites

#### Usage
The script needs to be run continuosly and pulls data daily at 18:00.
You can start it by typing `python datagetter.py`. If you want to force
a retrieval, you can do `python datagetter.py 1`.

The script will output `results.xlsx` which contains 4 sheets, one
for each unit types: studio, 1bd1ba, 2bd2ba, and one summarizing the
availability over time of these unit types. The first 3 sheets show
data on the available units, sorted by price. If an available unit
has just showed up it will be marked with "New". A sheet containing
an 'new' unit will have '(!)' in its name (for example _2Bd2Ba(!)_).
Units which are no longer available will still show up, but they
will be marked with 'Expired'. For each unit it will show historic data,
and for the first 5 on each sheet it will also output a bar chart
containing that data.

#### Disclaimer
This script analyzes data that is already publicly available. It is meant
for educational purposes only.
