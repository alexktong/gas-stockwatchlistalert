# Objective
Emails addressee with daily statistics of stocks on a watchlist on weekdays only.

Current defined metrics are are "Top 5 Daily Losers", "Top 5 Daily Risers" and "% From 52-Week Low".

## Step 1

Open `Stock Watchlist.xlsx` with Google Sheets.

## Step 2

In the script `Code.gs`:

* Assign variable "emailAddress" to the email address of the recipient.
* Assign variable "timeZoneOffset" to the timezone of the recipient's location.

## Step 3

For the script to run *once daily*, set a Trigger with following parameters:

**Which runs at deployment**: `Head`

**Select event source**: `Time-driven`

**Select type of time based trigger**: `Day timer` Enables choosing time of the day the script will run.

**Select time of day**: `5pm to 6pm` The script will run between 5pm and 6pm daily.
