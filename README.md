# Running Log

A command-line tool for logging runs and tracking training over time.
Built with Python and openpyxl, no database or web framework, just a
simple script that writes straight to an Excel file.

## What it does

Run it and it walks you through entering details about a run: date,
distance, total time, run type, location, shoes, how you felt, and any
notes. From distance and time it calculates your pace automatically.
Everything gets appended to `running_log.xlsx`, and after each entry it
prints your total mileage for the last 7 days.

## Features

- Validates input as you go (distance has to be a positive number, time
  has to be in HH:MM:SS or MM:SS format, feel rating has to be 1-10) so
  you can't accidentally save garbage data
- Calculates pace per mile automatically from distance and time
- Builds the Excel file with headers and bold formatting on first run,
  then just appends new rows after that
- Auto-adjusts column widths so the spreadsheet stays readable
- Tracks rolling 7-day mileage and prints it after every entry
- Lets you log multiple runs in one session before quitting

## Running it

```bash
pip install openpyxl
python running_log.py
```

It'll prompt you for each field. Press Enter on the date field to use
today's date automatically. After saving a run it asks if you want to
log another one, type "yes" to keep going or anything else to exit.

## Example

```
Welcome to your Running Log!
Date of run (YYYY-MM-DD, or press Enter for today):
Distance in miles: 5.2
Total time (HH:MM:SS or MM:SS): 42:30
Run type (easy, workout, long run, race, recovery): easy
Location: Home loop
Shoes worn: Pegasus
How did you feel? 1-10: 7
Notes: felt good

Run saved successfully!
Saved to running_log.xlsx
Your mileage for the last 7 days is: 5.20 miles

Would you like to enter another run? yes/no:
```

## Background

I ran NCAA Division I track and cross country in college and have kept
logging my training ever since. This started as a way to replace a
manual spreadsheet with something that handles the math (pace, weekly
mileage) and validation automatically, while still landing in a normal
Excel file I can open and look at anytime.

## Author

Timothy Champlin
