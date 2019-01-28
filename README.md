# Google Apps Script - From Gmail to Google Sheets

Here is a simple [Google Apps Script](https://developers.google.com/apps-script/) that processes credit card transaction emails and store the info (merchant, datetime, amount etc) on a spreadsheet.

Currently the script is written to process particular emails (those from PC Financial) with regular expressions.  You probably want to change that to suit your needs.

## Set up Gmail
To better organize the credit card emails and to analyze my expenses, my Gmail inbox is set up with:
- [a filter](https://support.google.com/mail/answer/6579?hl=en) to label these credit card transaction emails.  This makes it easier for the script to find them
- the same filter also "star" the emails.  This serves as an indicator on new emails that haven't been processed. The script will "unstar" processed emails

## The script
To deploy the script for your own, login [G Suite Developer Hub](https://script.google.com/home) with your Google account.  Create a new project and paste the code.  Remember to set `main` as the startup function.

For configuration, refer to the top of the code.  The script can be configured to find emails by labels.  You can also set the Google Sheets file name for storing the data. If no such Sheet exists, a new one will be created.

The first time you run the script, you will be prompted to grant OAuth access rights as it needs to access your Gmail, Google Drive, and Google Sheets.
