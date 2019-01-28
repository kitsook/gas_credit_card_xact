// Google Apps Script to process Gmail and log data into Google Sheets

// MIT License
//
// Copyright (c) 2019 Clarence Ho (clarenceho at gmail dot com)
//
// Permission is hereby granted, free of charge, to any person obtaining a copy
// of this software and associated documentation files (the "Software"), to deal
// in the Software without restriction, including without limitation the rights
// to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
// copies of the Software, and to permit persons to whom the Software is
// furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in
// all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
// SOFTWARE.

/**** Configurations ****/

// Gmail lables that contain threads of credit card purchases.
// Unprocessed threads are marked with stars
var GMAIL_LABELS = ['Credit Card Purchases'];

// The Google Sheet that transactions will be appended to.
// Records are added to the first sheet.
// You can add further processing / charts etc in other sheets
var SPREADSHEET_NAME = 'Credit Card Transactions';

/**** End of Configurations ****/


// Regular expressions to process PC Financial MasterCard
var REGEXP_PC_CARD_NUMBER = /Card number: \*+([0-9]+)/;
var REGEXP_PC_MERCHANT = /Merchant: (.+)/;
var REGEXP_PC_AMOUNT = /Purchase amount: (\$[0-9\.,]+)/;
var REGEXP_PC_TRANSACTION_DATE = /Transaction date: ([a-zA-Z0-9 ,]+)/;
var REGEXP_PC_AVAILABLE_CREDIT = /Available credit: (\$[0-9\.,]+)/;


/**
 * The main function. Select this as the function to run in Google Apps Script
 */
function main() {
  var labelObjs = getGmailLabels(GMAIL_LABELS);
  var ss = getSpreadsheet();
  var sheet = ss.getSheets()[0];
  
  for(var i=0; i < labelObjs.length; i++) {
    var threads = getUnprocessedThreads(labelObjs[i]);
    for(var j=threads.length-1; j>=0; j--) {
      processThread(threads[j], sheet);
    }
  }
}

/**
 * Get Gmail label objects to be processed
 *
 * @param {string[]} labels
 * @return {GmailLabel[]}
 */
function getGmailLabels(labels) {
  var result = [];
  for(var i=0; i < labels.length; i++){
    result.push(GmailApp.getUserLabelByName(labels[i]));
  }
  
  return result;
}

/**
 * Get unprocessed email threads with given label. Unprocessed emails are those starred.
 *
 * @param {GmailLabel} label
 * @param {Sheet} sheet
 * @return {GmailThread[]}
 */
function getUnprocessedThreads(label, sheet) {
  var from = 0;
  var perrun = 50; //maximum is 500
  var threads;
  var result = [];
  
  do {
    threads = label.getThreads(from, perrun);
    from += perrun;

    for (var i=0; i<threads.length; i++) {
      if (threads[i].hasStarredMessages()) {
        result.push(threads[i]);
      }
    }
  } while (threads.length > 0);
  
  Logger.log(result.length + ' threads to process in ' + label.getName());
  return result;
}

/**
 * Process messages in the thread
 *
 * @param {GmailThread} thread
 * @param {Sheet} sheet where to append the transaction
 */
function processThread(thread, sheet) {
  var messages = thread.getMessages();
  for(var i=0; i < messages.length; i++) {
    var message = messages[i];
    if(!message.isStarred()) {
      continue;
    }

    if (message.getSubject() === 'PC Financial Mastercard purchase notice') {
      var xact = processPCMessage(message.getPlainBody())
      Logger.log(xact);
      addTransaction(sheet, xact);
    }
    
    // mark the message as processed by removing the star
    message.unstar();
  }
}

/**
 * Process the message from PC Financial. Return an object of transaction details.
 *
 * @param {GMailMessage} message
 * @return {Object}
 */
function processPCMessage(message) {
  var cardNumber = REGEXP_PC_CARD_NUMBER.exec(message);
  var merchant = REGEXP_PC_MERCHANT.exec(message);
  var amount = REGEXP_PC_AMOUNT.exec(message);
  var transactionDate = REGEXP_PC_TRANSACTION_DATE.exec(message);
  var availableCredit = REGEXP_PC_AVAILABLE_CREDIT.exec(message);
  
  return {
    cardNumber: cardNumber? cardNumber[1] : null,
    merchant: merchant? merchant[1].trim() : null,
    amount: amount? amount[1] : null,
    transactionDate: transactionDate? transactionDate[1] : null,
    availableCredit: availableCredit? availableCredit[1] : null
  };
}

/**
 * Create a new Spreadsheet and add headers to the first sheet
 *
 * @return {Spreadsheet}
 */
function createSpreadsheet() {
  var ssNew = SpreadsheetApp.create(SPREADSHEET_NAME);
  var sheet = ssNew.getSheets()[0];
  addSheetHeader(sheet);

  return ssNew;
}

/**
 * Add header to the worksheet
 *
 * @param {Sheet} sheet
 */
function addSheetHeader(sheet) {
  var headers = ['Card Number', 'Merchant', 'Amount', 'Date', 'Available Credit'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
}

/**
 * Get the Spreadsheet to update. Create one if no existing Spreadsheet found
 *
 * @return {Spreadsheet}
 */
function getSpreadsheet() {
  var files = DriveApp.searchFiles(
    'title="' + SPREADSHEET_NAME + '" and mimeType="' + MimeType.GOOGLE_SHEETS + '"');
  if (files.hasNext()) {
    // take the first one
    return SpreadsheetApp.open(files.next());
  }

  return createSpreadsheet();
}

/**
 * Add a transaction record to the sheet
 *
 * @param {Sheet} sheet
 * @param {Object} xact
 */
function addTransaction(sheet, xact) {
  var lastRow = sheet.getLastRow();
  var values = [ xact['cardNumber'], xact['merchant'], xact['amount'], xact['transactionDate'], xact['availableCredit'] ];
  sheet.getRange(lastRow+1, 1, 1, values.length).setValues([ values ]);
}