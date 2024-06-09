---
title: "Google AppSheet: From Sheet to Database to Sheet"
---
## Perpetually Pinching Pennies
Working within a small company that prefers the $0 solution, I often found myself trying to make everything fit into Free tiers or Open Source software. I had a very short time to convert the workflow from our [corporate ordering spreadsheet]({% post_url 2024-03-29-that-one-time-I-overengineered-a-ss %}) to something more user friendly. Originally I was recreating it in Oracle Apex. I had half an app before deciding to switch for the sake of time and easier deployment. 

## An okay outcome
I was originally using the new DB features in AppSheet. It was easy to use and was solid for most functions. I soon found that there was a limit to the number of entries available in these databases and had to switch to using Sheets. There is a function that allows users to upload a CSV file. However, this function is not able to add any information. This created a problem where an order number needed to be generated and filled along with the line items that were being uploaded via CSV. The workaround was to have a placeholder table that was used for copying over the line items. I added a script to run immediately after the upload that filled the needed information and copied the line items into the main sheet/database. 

## The script to make it work

~~~ javascript
function addNewFileData(rowID) {
  const ss = SpreadsheetApp;
  const order = ss.openById('sheet_id').getActiveSheet();
  let ordervals = order.getSheetValues(2,1,order.getLastRow() -1, order.getLastColumn());
  let cleararea = order.getRange(2,1,order.getLastRow() -1, order.getLastColumn());
  cleararea.clear()
  const sheet = ss.openById('sheet_id').getActiveSheet();
  let numrange = sheet.getRange(1,1,sheet.getLastRow()).getValues();
  var numlast = numrange.length;
  sheet.getRange(sheet.getLastRow() +1, 2, ordervals.length, ordervals[0].length).setValues(ordervals);
  sheet.getRange(numlast +1,1, ordervals.length).setValue(rowID);
}
~~~