# description
Request SEC (Security and Exchange Commission) Data via their API-URL
This script should request data from https://data.sec.gov/api/xbrl (in different endpoints) and export them to excel.
Beforehand it will filter "double data" in the form of:
- reports will be reduced to one per filing date. if the same date is found in multiple filings, only the newest filing will be exported.
> default setting: the "most recent" filing number (date) should be the most correct one. older values will be replaced by newer values (filing dates).

# install

download the script
edit personal settings (can be found on top of script)
run the script with powershell

# use

has only been used in Powershell ISE until now.
Doesn't run perfectly yet

## info
when running the script it creates an Excel file (currently in C:\Temp\Reports\).

this script has not been optimized on every stock. 
errors may arrive if certain data is not found that has not yet been tested.
