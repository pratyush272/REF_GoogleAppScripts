function myFunction() {
   

 // Log the subject lines of the threads labeled with MyLabel
 var label = GmailApp.getUserLabelByName("Contact");
 var threads = label.getThreads(0,50);
  var file = DriveApp.getFilesByName("contactemails")
  var spreadsheet = SpreadsheetApp.open(file.next());
  var sheet = spreadsheet.getSheets()[0];
  
 for (var i = 0; i < threads.length; i++) {
   //Logger.log(threads[i].getFirstMessageSubject());
   var label2 = GmailApp.getUserLabelByName("Contact/Scanned");
   threads[i].addLabel(label2);
   var extractedemail = (threads[i].getMessages()[0]).getFrom();
   
   if (extractedemail.indexOf("<") > 0)
   {
     var n = extractedemail.indexOf("<");
     extractedemail = extractedemail.substring (n+1,extractedemail.length-1);
   }
   var duplicatechecker =0;
   var data = sheet.getDataRange().getValues();
   for (var j = 0; j < data.length;j++ ) {
    if(extractedemail === data[j][0])
    {duplicatechecker++;
     Logger.log("found a duplicate..skipping");
     Logger.log("j = " + j + "checking " + extractedemail + " against " + data[j][0])
     break;
    }
  }
   if (duplicatechecker ==0){ Logger.log("no duplicate for"+extractedemail);
     sheet.appendRow([extractedemail]);}
   //sheet.appendRow(['Cotton Sweatshirt XL', 'css004']);
   //Logger.log(extractedemail.substring(1,extractedemail.length-1));
 }
   
 
  //
  
  

 }

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index');
}
 

