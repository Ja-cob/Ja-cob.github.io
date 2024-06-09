---
title: "That One Time I Over-engineered A Spreadsheet"
---

## The Problem

This was my first major project using Google Sheets; updating a paper only workflow for tracking corporate orders. There was a short 6 week timeline to implement a new system. The major complication with corporate orders is the use of multiple shipping addresses per order while also tracking the batch of shipments under one order number. The secondary problem was having a way to export order data in a format that could be used within the limitations of ShipStation (our fulfillment software). The original data has been replaced with sample data for this demonstration. Looking back, the end result was slightly overengineered and messy; but it truly shows the many different capabilities of Google Sheets and Apps Script. Functionality was tailored to fit the format of most orders we received. Exceptions will be noted at the end.

## Page 1

![Image]({{site.baseurl}}/assets/images/post2a.webp){: width="650" }

This was the Order Info page. A very generic looking spreadsheet. Green cells indicate a place where information needs to be manually entered (across all pages). The user was to copy and paste (from our order form) information into the gray columns and fill in any pertinent green cells. The rest of the information was auto-filled. This page formats all the information into what is needed to bulk import shipments in ShipStation and can be saved as a .csv file. 

## Page 2

![Image]({{site.baseurl}}/assets/images/post2b.webp){: width="650" }

Page 2 is for quotes and scheduling. Info from page 1 is pulled, product cost is totaled, discount tier calculated and applied, available shipping days can be selected based on the holiday or a custom date can be entered, and there are buttons for clearing or submitting the information. There are scripts attached to the buttons that either erase the first two sheets or copy the order information into the "backend" of the sheet for the actual tracking of orders:

~~~ javascript
function sendOrder() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var result = SpreadsheetApp.getUi().alert("Are you sure you want to submit this order and clear the sheet?", SpreadsheetApp.getUi().ButtonSet.OK_CANCEL);
  if(result === SpreadsheetApp.getUi().Button.OK) {
    var orderTracker = ss.getSheetByName("Order Status Tracker");
    var shipSched = ss.getSheetByName("Ship Date Scheduler");
    var copyPage = ss.getSheetByName("Copy Logic")
    var rangers = ss.getRangeList(['D2:D32','K2','N6:N15']);

  // get source range
    var source = ss.getRangeByName("quoteinfo");
    var csv = ss.getRangeByName("copydata");
    var namers = ss.getRangeByName("cusname");
    var source2 = copyPage.getRange(2,1,copyPage.getLastRow(),4);

  // get destination range
    var destination = orderTracker.getRange(orderTracker.getLastRow()+1,1,1,7);
    var destiny2 = shipSched.getRange(shipSched.getLastRow()+1,1,1,4);

  // copy values to destination range
    source.copyTo(destination,{contentsOnly:true});
    source2.copyTo(destiny2,{contentsOnly:true});

  // clear source values
    rangers.clearContent();
    csv.clearContent();
    namers.clearContent();

function clearAll() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var csv = ss.getRangeByName("copydata");
  var namers = ss.getRangeByName("cusname");
  var rangers = ss.getRangeList(['D2:D32','K2','N6:N15']);
  rangers.clearContent();
  csv.clearContent();
  namers.clearContent();
}
~~~

## Pages 3-8 

![Image]({{site.baseurl}}/assets/images/post2c.webp){: width="650" }
![Image]({{site.baseurl}}/assets/images/post2d.webp){: width="650" }
![Image]({{site.baseurl}}/assets/images/post2e.webp){: width="650" }

The next series of pages hold different sets of order information that were relevant to different departments. So, there is a general order tracking page for sales and fulfilment tracking, a shipment total calendar for quick reference, a calendar for the production team broken down by product and the day they will ship, and then the back end (product database, a page the script uses to format and copy order info, etc...). 

There were additional pages that including one made to format and print order info (similar to my recipe sheet) and a page to track changes that needed to be made to orders. These were removed when I replaced the data due to some redundancies and changes in the way we ended up handling orders.

## Issues

Despite all the features I threw into it, like all spreadsheets, the formulas and functionality were highly susceptible to user error, even with most regions being protected. 

Any special cases with multiple products per shipment or anything that required new order numbers required manual manipulation (I personally handled all special cases).

Only one person could be entering an order at a time

Any changes made after the order was entered were usually done via ShipStation and weren't reflected in the calculations on the tracking areas.

Overall, the majority of orders went fairly smooth. I'll be covering the system that replaced this one in a separate post