# gsheets-email
Google Apps Script that interacts with Google Sheets that contains a stock watchlist and emails the addressee with daily statistics.

## Step 1

Open `Watchlist.xlsx` with Google Sheets. It contains named ranges necessary to interact with the Google Apps Script.

## Step 2

In the script `Code.gs`, assign variable "emailAddress" to the email address of recipient.

## Step 3

For the script to run once daily, set a Trigger with following parameters:

**Which runs at deployment**: `Head`

**Select event source**: `Time-driven`

**Select type of time based trigger**: `Day timer` Enables choosing time of the day the script will run.

**Select time of day**: `5pm to 6pm` The script will run between 5pm and 6pm daily.

Even though the script is automatically run daily, you will only receilve daily email alert on **weekdays** because of the condition in the script `if (day >= 1 && day <= 5)`.
