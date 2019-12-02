// --------------------------------------------------------------------------------------------------
//    SquareSpace Orders API to Google Sheets translator
//    Copyright (C) 2019  Will Bradley. Licensed under GNU GPLv3:
//
//    This program is free software: you can redistribute it and/or modify
//    it under the terms of the GNU General Public License as published by
//    the Free Software Foundation, either version 3 of the License, or
//    (at your option) any later version.
//
//    This program is distributed in the hope that it will be useful,
//    but WITHOUT ANY WARRANTY; without even the implied warranty of
//    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
//    GNU General Public License for more details.
//
//    You should have received a copy of the GNU General Public License
//    along with this program.  If not, see <https://www.gnu.org/licenses/>.
// --------------------------------------------------------------------------------------------------

var API_KEY = "abcd1234-1234-12ab-123-abcd1234abcd"; // set to your squarespace api key
var DEBUG = 1; // 0 is no logging, 1 is some logging, 2 is lots of logging
var TIMEZONE = "America/Los_Angeles"; // assume California
var START_ROW = 1; // fill the whole sheet (no other header data)
var SORT_COLUMN = 5; // col 5 should be the order date
var SHEET_NAME = "All Orders"; // set to the spreadsheet tab name you want data to appear in

/*
 * Structure of SquareSpace Orders API response:
 *
 
{
  "result" : [ {
    "id" : "1a2c3d1a2c3d1a2c3d1a2c3d",
    "orderNumber" : "108",
    "createdOn" : DateTime,
    "modifiedOn" : DateTime,
    "testmode" : false,
    "customerEmail" : "cust@example.com",
    "billingAddress" : CustAddress,
    "shippingAddress" : CustAddress,
    "fulfillmentStatus" : "PENDING",
    "lineItems" : [ {
      "id" : "1a2c3d1a2c3d1a2c3d1a2c3d",
      "sku" : "SQ0123456",
      "productId" : "1a2c3d1a2c3d1a2c3d1a2c3d",
      "productName" : "Some Product Name",
      "quantity" : 1,
      "unitPricePaid" : Currency,
      "variantOptions" : [ {
        "optionName" : "Date/Time",
        "value" : "January 1st, 2019 10:00 AM-2:30 PM"
      } ],
      "customizations": [ {
        "label":"Name",
        "value":"Jane Doe"
      }, {
        "label":"E-mail",
        "value":"jdoe@example.com"
      }, {
        "label":"Are you currently a Chimera member?",
        "value":"No"
      }, {
        "label":"Phone Number",
        "value":"707-555-1212"
      }, {
        "label":"Refund Policy",
        "value":"I Agree"
      } ],
      "imageUrl" : "https://static1.squarespace.com/static/1234/1234/1234/IMG_1234.JPG?format=300w"
    } ],
    "internalNotes" : [ ],
    "shippingLines" : [ ],
    "discountLines" : [ ],
    "formSubmission" : [ {
      "label" : "Note / Additional Info",
      "value" : ""
    }, {
      "label" : "If purchasing more than one class or item, please tell us who it is for:",
      "value" : " "
    } ],
    "fulfillments" : [ ],
    "subtotal" : Currency,
    "shippingTotal" : Currency,
    "discountTotal" : Currency,
    "taxTotal" : Currency,
    "refundedTotal" : Currency,
    "grandTotal" : Currency
  },
  {...},
  {...},
  {...},
  etc
}

*
* Objects referenced above:
*

DateTime: "2019-04-25T13:39:38.187Z"
Currency: {
    "currency" : "USD",
    "value" : "5.00"
}
CustAddress: {
    "firstName" : "John",
    "lastName" : "Smith",
    "address1" : "123 main st",
    "address2" : null,
    "city" : "Sometown",
    "state" : "CA",
    "countryCode" : "US",
    "postalCode" : "98765",
    "phone" : "123-123-1234"
}

*
* Example usage:
*

response['result'][0]['id']
response['result'][0]['customerEmail']
response['result'][0]['lineItems'][0]['productName']
response['result'][0]['lineItems'][0]['quantity']
response['result'][0]['lineItems'][0]['unitPricePaid']['value']
response['result'][0]['lineItems'][0]['unitPricePaid']['value']
response['result'][0]['grandTotal']['value']

*
* End API example
*/

var activeSheet = null;

// custom menu
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('SquareSpace Order Import')
      .addItem('Import Order Data','manuallyImportOrderData')
      .addToUi();
}

// function to call SquareSpace Orders API
function callSquareSpaceOrdersAPI() {
  var headers = {'Authorization': 'Bearer '+API_KEY};
  var options = {'headers': headers};
  var url = "https://api.squarespace.com/1.0/commerce/orders/";
  var raw = UrlFetchApp.fetch(url, options);
  
  // Parse the JSON reply
  var response = JSON.parse(raw.getContentText());
  
  if(DEBUG>=2){Logger.log(response);}
  
  return response;
}
 
function manuallyImportOrderData() {
  // when doing it manually, set the sheet to the currently active one
  activeSheet = ss.getActiveSheet();
  importOrderData();
}
 
function importOrderData() {
  
  // pick up the search term from the Google Sheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  // when doing it automatically, set the sheet based on the variable.
  if (activeSheet == null) {
    sheet = ss.getSheetByName(SHEET_NAME);
  } else {
    sheet = activeSheet;
  }
  
  //var artist = sheet.getRange(11,2).getValue();
  
  var response = callSquareSpaceOrdersAPI();
  var results = response["result"];
  
  var output = [];
  
  // loop through each order
  if (results) {
    results.forEach(function(order,orderIndex) {
      // orders can have multiple unrelated line items: loop through them too.
      if (order['lineItems']) {
        order['lineItems'].forEach(function(lineItem,lineItemIndex) {
          if(DEBUG>=2){Logger.log(lineItemIndex+": "+JSON.stringify(lineItem));}
          if(DEBUG>=2){Logger.log(lineItemIndex+": "+lineItem['unitPricePaid']);}
          unitPrice = lineItem['unitPricePaid']['value'];
          if(DEBUG>=2){Logger.log(orderIndex+": "+order['grandTotal']);}
          grandTotal = order['grandTotal']['value'];
          orderDate = new Date(order["createdOn"]);
          orderDateStr = Utilities.formatDate(orderDate, TIMEZONE, "yyyy-MM-dd HH:mm:ss z");
          
          // extract class date, name, phone, and notes if they exist
          classDate = null;
          customerName = null;
          customerPhone = null;
          additionalInfoArray = [];
          isMember = null;
          if (lineItem["variantOptions"]) {
            lineItem["variantOptions"].forEach(function(variant){
              if (variant["optionName"] == "Date/Time" || variant["optionName"] == "Date") {
                classDate = variant["value"];
              }
            });
          }
          // start by assuming name/phone from the address, overwrite it below if provided
          if (order["billingAddress"]){
            customerName = [order["billingAddress"]["firstName"],order["billingAddress"]["lastName"]].join(" "); // full name is first/last joined by a space
            customerPhone = order["billingAddress"]["phone"];
          }
          if (lineItem["customizations"]) {
            lineItem["customizations"].forEach(function(customization){
              if (customization["label"] == "Name" && customization["value"] != "" && customization["value"] != " ") {
                customerName = customization["value"];
              }
              if (customization["label"] == "Phone Number" && customization["value"] != "" && customization["value"] != " ") {
                customerPhone = customization["value"];
              }
              if (customization["label"] == "Are you currently a Chimera member?" && customization["value"] != "" && customization["value"] != " ") {
                isMember = customization["value"];
              }
            });
          }
          // append additional info together in the notes field separated by a semicolon
          if (order["formSubmission"]) {
            order["formSubmission"].forEach(function(submissionInfo){
              if (submissionInfo["label"] == "Note / Additional Info" && submissionInfo["value"] != "" && submissionInfo["value"] != " ") {
                additionalInfoArray.push(submissionInfo["value"]);
              }
              if (submissionInfo["label"] == "If purchasing more than one class or item, please tell us who it is for:" && submissionInfo["value"] != "" && submissionInfo["value"] != " ") {
                additionalInfoArray.push("Purchased for: "+submissionInfo["value"]);
              }
            });
          }
          additionalInfo = additionalInfoArray.join("; ");
          
          // Make this row match the header titles below exactly.
          row = [
            JSON.stringify(order),
            customerName,
            lineItem["productName"],
            classDate,
            orderDateStr,
            lineItem["quantity"],
            "$"+unitPrice,
            customerPhone,
            order["customerEmail"],
              additionalInfo,
                isMember
                ];
          if(DEBUG>=1){Logger.log(lineItemIndex+": "+orderDateStr);}
          output.push(row);
        });
      }
    });
  }

  // custom comparator function to sort by date
  var sortedOutput = output.sort( function(a,b) {
    var orderADate = a[SORT_COLUMN-1]; // column 1 is at the 0 index number
    var orderBDate = b[SORT_COLUMN-1]; // column 1 is at the 0 index number
    if (orderADate < orderBDate) {
      return -1; // A is less than B, move B up
    }
    else if (orderADate > orderBDate) {
      return 1; // A is greater than B, move A down
    }
    return 0; // names are equal, return zero
  });

  // set names for each column of output
  var header = [
    "Imported E-Mail Data",
    "Attendee Name",
    "Purchased",
    "Class Date",
    "Order placed on",
    "Qty",
    'Unit Price',
    "Phone",
    "E-mail",
    "Additional Info",
    "Member?"
  ];
  // this variable needs to be the column number (starting at 1) of the Order Date (Order Placed On) column.
  var ORDER_DATE_COLUMN_NUMBER = 5;
  // calculate width and length
  var columnCount = header.length;
  var rowCount = sortedOutput.length;
  
  // for now, don't clear and paste in headers. we don't want to overwrite.
  // sheet.getRange(START_ROW,1,1,columnCount).clearContent();
  // sheet.getRange(START_ROW,1,1,columnCount).setValues([header]); // single row needs to be passed inside of a parent array

  lastInsertedRowNumber = sheet.getLastRow()-1; // gotta pick somewhere to start, let's start at the last row number
  lastInsertedOrderDate = sheet.getRange(lastInsertedRowNumber, ORDER_DATE_COLUMN_NUMBER).getValue(); // get the last order date inserted into the DB.

  if(DEBUG>=1){Logger.log(JSON.stringify( sheet.getRange(lastInsertedRowNumber, ORDER_DATE_COLUMN_NUMBER) ));}
  sortedOutput.forEach(function(row, index){
    // if this row's order number is greater than the last inserted order number, then insert.
    // (we basically only want to add new rows, not overwrite existing rows.)
    // (this relies on the sheet and the sortedOutput being sorted by order date ascending, by the way)
    if (row[ORDER_DATE_COLUMN_NUMBER-1] > lastInsertedOrderDate) {
      thisRowNumber = lastInsertedRowNumber+1;
      // clear any previous content (there should be none)
      sheet.getRange(thisRowNumber,1,1,columnCount).clearContent();      
      // paste in the values
      sheet.getRange(thisRowNumber,1,1,columnCount).setValues([row]); // gotta put the row inside an array cuz setValues accepts two dimensions
      // we just inserted a row, so now this row is the last inserted row (for the next round)
      lastInsertedRowNumber = thisRowNumber;
      lastInsertedOrderDate = row[ORDER_DATE_COLUMN_NUMBER-1];
      if(DEBUG>=1){Logger.log(thisRowNumber+": inserted "+lastInsertedOrderDate);}
    } else {
      if(DEBUG>=1){Logger.log(lastInsertedRowNumber+": did not insert "+row[ORDER_DATE_COLUMN_NUMBER-1]+" because prior date was "+lastInsertedOrderDate);}
    }
  });

  // formatting
  //sheet.setRowHeight(START_ROW,65);
  //sheet.getRange(START_ROW,1,500,columnCount).setVerticalAlignment("middle");
  //sheet.getRange(START_ROW,5,500,1).setHorizontalAlignment("center");
  //sheet.getRange(START_ROW,2,len,3).setWrap(true);
}


