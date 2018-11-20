//Citing my sources
//https://ctrlq.org/code/20279-import-csv-into-google-spreadsheet
//https://developers.google.com/apps-script/reference/gmail/gmail-attachment
//https://productforums.google.com/forum/#!msg/docs/-JmsVUBGcRY/D-m0O-wNDgQJ
//https://stackoverflow.com/questions/22898501/create-a-new-sheet-in-a-google-spreadsheet-with-google-apps-script
//https://stackoverflow.com/questions/41106334/set-date-as-sheet-name-in-spreadsheet
//https://ctrlq.org/code/19973-duplicate-sheet-google-spreadsheets
//https://stackoverflow.com/questions/41106334/set-date-as-sheet-name-in-spreadsheet
//https://stackoverflow.com/questions/28295056/google-apps-script-appendrow-to-the-top
//http://www.ryanpraski.com/google-sheets-remove-empty-columns-rows-automatically/

//*******************************************************************************************************************
//*******************************************************************************************************************
//*******************************************************************************************************************

//create a menu option for script functions
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Parse for FS CSV Files in Gmail')
  .addItem('Return to Overview', 'overview')
  .addItem('Import Latest', 'importLatestCSVFromGmail')
  .addItem('Update Changelog', 'update')
  .addItem('Update Changelog with New Employees Only!', 'updateNewEmployees')
  .addItem('Status', 'status')
  .addToUi();
}
//*******************************************************************************************************************
//Show the current and backup dates for the current user
function status(){
  
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName('Current_Employees');
  var ssBackup = spreadsheet.getSheetByName('Backup');
  var ssNew = spreadsheet.getSheetByName('CLAS New Employee Status Change');  
  var ssDept = spreadsheet.getSheetByName('Department Change');
  var ssFormer = spreadsheet.getSheetByName('CLAS Former Employee Status Change'); 
  var ssHistory = spreadsheet.getSheetByName('History'); 
  
  var import = sheet.getRange(1, 1).getValue();
  var backup = ssBackup.getRange(1, 1).getValue();
  var sheetRange = sheet.getLastRow()-2;
  var ssBackupRange = ssBackup.getLastRow()-2;
  var ssNewRange = ssNew.getLastRow()-2;
  var ssDeptRange = ssDept.getLastRow()-2;
  var ssFormerRange = ssFormer.getLastRow()-2;
  var ssHistoryRange = ssHistory.getLastRow()-2;
  Logger.log(ssDeptRange);
  if (ssDept.getRange("A3").isBlank()) {
    
    ssDeptRange--;
    
  }
  Logger.log(ssDeptRange);
  Browser.msgBox("Status", "Last import date: " + import + "\\nDate of last backup: " + backup +
                 "\\nEmployee count in last import: " + sheetRange + "\\nEmployee count in last backup: " + ssBackupRange +
                 "\\nNew CLAS employees: " + ssNewRange + "\\nEmployees who changed departments: " + ssDeptRange + "\\nEmployees who left the college: " + ssFormerRange +
                 "\\nEmployees in History Change Log: " + ssHistoryRange, Browser.Buttons.OK);
}
//*******************************************************************************************************************
//Go back to the starting page for easy navigation
function overview(){
  
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();//.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName('Overview');
  sheet.activate();
  spreadsheet.moveActiveSheet(1);
}

//*******************************************************************************************************************
//*******************************************************************************************************************
//*******************************************************************************************************************
/**
* Searches all banner dump csvs, make the newest one the primary, back up the old one
*
* 
*
* @customfunction
*/
function importLatestCSVFromGmail() {
  
  Logger.log("Link to spreadsheet: https://docs.google.com/spreadsheets/d/1OW_u0j97OLu_kmcyysdohlgsM22ymSNZk29k5BkNBxA/edit#gid=413263414");
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ssNew = ss.getSheetByName('CLAS New Employee Status Change');  
  var ssDept = ss.getSheetByName('Department Change');
  var ssFormer = ss.getSheetByName('CLAS Former Employee Status Change'); 
  var ssHistory = ss.getSheetByName('History'); 
  var sheetCE = ss.getSheetByName('Current_Employees');
  var ssOverview = ss.getSheetByName('Overview');
  //If you get a timezone error "Invalid argument: timeZone. Should be of type: String (line 57, file "Code")", go to the Spreadsheet, +
  // "File" -> "Spreadsheet Settings", change timezone, save, and then change it back. Appears to happen with spreadshet revision 
  var timeZone = Session.getScriptTimeZone();
  var hold_date = Utilities.formatDate(new Date(), timeZone, 'MM-dd-yyyy');
  var todaysDate = Utilities.formatDate(new Date(), timeZone, 'MM-dd-yyyy');
  var archiveFolder = DriveApp.getFolderById('1Fqb901vr5CVSzyhBHHeabdMHVKSl890S');
  var updater = 0;
  
  var ssNewTime = ssNew.getRange("A1");
  var ssNewFilter = ssNew.getRange("A2");
  var ssDeptTime = ssDept.getRange("A1");
  var ssDeptFilter = ssDept.getRange("A2");
  var ssDeptFD = ssDept.getRange("H2");
  var ssDeptFDHD = ssDept.getRange("I2");
  var ssFormerTime = ssFormer.getRange("A1");
  var ssFormerFilter = ssFormer.getRange("A2");
  
  //This is the date of the last backup, on the spreadsheet in A1
  if (sheetCE.getRange(1,1).getValue() == null) {
    var sheet_date = 1;
  }else{
    var sheet_date = sheetCE.getRange(1,1).getValue();  
  }
  
  Logger.log('Sheet date is '+sheet_date);
  
  //Searches Gmail for latest message with attachment that fits the criteria "filename: *current_employee_directory_listing*.csv has:attachment"
  var threads = GmailApp.search("filename:*current_employee_directory_listing.csv* has:attachment ");
  
  if (threads.length==0){
    
    //No results from search
    ss.toast('No results from search', 'Complete', 10);
    Logger.log('Quitting, no backup performed, no results.');
    return;
  }
  
  var msgs = GmailApp.getMessagesForThreads(threads);
  var email_date = threads[0].getMessages()[0].getDate();
  var temp_date = threads[0].getMessages()[0].getDate();
  var email_date = temp_date;
  var attachment = threads[0].getMessages()[0].getAttachments()[0];
  
  for (var i = 0 ; i < msgs.length; i++) {        
    for (var j = 0; j < msgs[i].length; j++) {
      //var attachments = msgs[i][j].getAttachments();
      temp_date = threads[i].getMessages()[j].getDate();
      Logger.log('attachment date: '+email_date);
      
      if (temp_date.valueOf() > email_date.valueOf()){
        Logger.log('temp_date.valueOf() > email_date.valueOf(): '+ temp_date.valueOf() + '>' + email_date.valueOf());
        email_date = temp_date;
        attachment = threads[i].getMessages()[j].getAttachments()[0];
      } else {
        Logger.log('temp_date.valueOf() != email_date.valueOf(): '+ temp_date.valueOf() + '!=' + email_date.valueOf());
      }
      
      /*for (var k = 0; k < attachments.length; k++) {
      Logger.log('Message "%s" contains the attachment "%s" (%s bytes)',
      msgs[i][j].getSubject(), attachments[k].getName(), attachments[k].getSize());
      }*/
    }
  }
  //Log information for easy visual & troubleshooting
  Logger.log('attachment date FINAL: '+email_date);  
  //UPDATE: Fixed! Old: 'Need to get email date of last message in latest thread! Right now we're getting the date of the first email in the latest thread'
  Logger.log('The email date is: '+email_date);
  Logger.log('attachment: '+attachment);
  Logger.log('EMAIL DATE IS '+email_date);  
  Logger.log('TODAY\'S DATE IS '+todaysDate);
  //Logger.log('ATTACHMENT TYPE IS '+attachment.getContentType());
  
  //Make sure sheets exist
  if (ss.getSheetByName('Current_Employees') == null) {    
    var sheetCE = ss.insertSheet();
    sheetCE.setName('Current_Employees');
    sheetCE = ss.getSheetByName('Current_Employees');
    
  } else {        
    var sheetCE = ss.getSheetByName('Current_Employees');   
  }
  Logger.log('sheetCE: '+sheetCE);
  
  
  //Parse the CSV attachment
  
  
  var csvData = Utilities.parseCsv(attachment.getDataAsString(), ",");
  //Backup csv file to folder
  backupToCSV(archiveFolder,attachment,sheetCE);
  
  //check to see if the csv from the email is newer than the current sheet  
  if (email_date.valueOf() > sheet_date.valueOf()){
    Logger.log(email_date.valueOf()+' is newer than '+ sheet_date.valueOf());
    
    //******************************************************************************************************************************************************
    //First make sure the spreadsheet isn't damaged by seeing if it has less than 4000 employees in the sheet
    if (csvData.length < 4000){
      
      //var ui = SpreadsheetApp.getUi();
      var result = Browser.msgBox(
        'The CSV found only has '+ (csvData.length-1) +' employees in it, indicating there may be damage or it is incomplete. '+
        'Thus, we will update the History Log with the new employees, but not with the employees who left (since it may be inaccurate). \\n'+
          '\\nDo you want to proceed and update the History Log? \\n'+
            '\\nYes - Yes, update the History Log with new employees (be sure to verify for accuracy) '+
              '\\nNo - No, run the script and update the employee lists but do not update the History Log '+
                '\\nCancel - Exit this script, change nothing ',
                  Browser.Buttons.YES_NO_CANCEL);
      
      // Process the user's response.
      if (result == 'yes') {
        // User clicked "Yes".
        updater = -1;
        Browser.msgBox('Confirmation received. The CSV will be parsed, but the History change log will only update with the new employees. Make sure you double check for accuracy.');
        
      } else if (result == 'no') {
        // User clicked "No".
        updater = 1;
        Browser.msgBox('Confirmation received. The CSV will be parsed, but the History change log will not be updated. If you would like it to be updated, run the Update Changelog function.');
        
      } else {
        // User clicked "Cancel" or X in the title bar.
        //Sheet from email is broken or incomplete, so let's not do anything
        ss.toast('The newest CSV you have is broken or incomplete.', 'Unsuccessful', 10);
        Logger.log('Quitting, no backup performed.');
        
        //Browser.msgBox('Denied.');
        
        Logger.log('Result is: ' + result);
        return;
      }
    }
    //remove formulas so they won't calculate until the end
    ssNewTime.clear();
    ssNewFilter.clear();
    ssDeptTime.clear();
    ssDeptFilter.clear();
    ssDeptFD.clear();
    ssDeptFDHD.clear();
    ssFormerTime.clear();
    ssFormerFilter.clear();
    
    //******************************************************************************************************************************************************
    
    //Back up older sheet before deleting it and replacing it with newest
    ss.toast('Parsing CSV...', 'Newer CSV Found', 10);
    
    var yourNewSheet = ss.getSheetByName('Backup');
    
    if (yourNewSheet != null) {
      ss.deleteSheet(yourNewSheet);
    }
    
    yourNewSheet = ss.insertSheet();
    yourNewSheet.setName('Backup');
    
    //copy sheet to new sheet with name + date
    //https://ctrlq.org/code/19973-duplicate-sheet-google-spreadsheets
    
    var name = yourNewSheet.getName();
    var backup_sheet = ss.getSheetByName('Current_Employees').copyTo(ss);
    //SpreadsheetApp.flush();
    
    //Delete the Backup sheet so we can back up the next-to-newest sheet
    var old = ss.getSheetByName(name);
    if (old) ss.deleteSheet(old); // or old.setName(new Name);  
    //SpreadsheetApp.flush(); // Utilities.sleep(2000);
    backup_sheet.setName(name);
    //backup_sheet.protect();
    //ssOverview.activate();
    //ss.moveActiveSheet(1);
    
    // Clear the Current Employees sheet, then can import new data
    sheetCE.clearContents().clearFormats();
    sheetCE.getRange(1, 1).setValue(email_date);
    sheetCE.getRange(2, 1, csvData.length, csvData[0].length).setValues(csvData);
    
    //Enter date at top of data, freeze rows + columns, protect sheet
    //sheetCE.insertRowBefore(1).getRange(1, 1).setValue(hold_date);
    sheetCE.setFrozenRows(2);
    sheetCE.setFrozenColumns(3);
    //sheetCE.protect();
    
    //******************************************************************************************************************************************************
    
    //Set formulas in case they get broken
    
    ssNewTime.setFormula('=TEXT(INDIRECT("Current_Employees!A1"),"mm-dd-yyyy") &": New Employees (In Current_Employees, not in Backup)"');
    
    ssNewFilter.setFormula('=IFERROR({{INDIRECT("Current_Employees!$A$2:$C$2"),INDIRECT("Current_Employees!$J$2"), INDIRECT("Current_Employees!$K$2"),INDIRECT("Current_Employees!$O$2"),INDIRECT("Current_Employees!$P$2")};{FILTER({INDIRECT("Current_Employees!$A:$C"),INDIRECT("Current_Employees!$J:$J"),INDIRECT("Current_Employees!$K:$K"),INDIRECT("Current_Employees!$O:$O"),INDIRECT("Current_Employees!$P:$P")},ISERROR(MATCH(INDIRECT("Current_Employees!$AB:$AB"),INDIRECT("Backup!$AB:$AB"),0)),len(INDIRECT("Current_Employees!$A:$A")),INDIRECT("Current_Employees!$O:$O")="Col Liberal Arts & Science (Col)")}},{INDIRECT("Current_Employees!$A$2:$C$2"),INDIRECT("Current_Employees!$J$2"), INDIRECT("Current_Employees!$K$2"),INDIRECT("Current_Employees!$O$2"),INDIRECT("Current_Employees!$P$2")})');
    
    
    ssDeptTime.setFormula('=TEXT(INDIRECT("Current_Employees!A1"),"mm-dd-yyyy") & ": Department changed (Current_Employees value is different from Backup)"');
    
    ssDeptFilter.setFormula('=IFERROR({{INDIRECT("Current_Employees!$A$2:$C$2"),INDIRECT("Current_Employees!$J$2"),INDIRECT("Current_Employees!$K$2"),INDIRECT("Current_Employees!$O$2"),INDIRECT("Current_Employees!$P$2")};{FILTER(({INDIRECT("Current_Employees!$A:$C"),'+
                            'INDIRECT("Current_Employees!$J:$J"),INDIRECT("Current_Employees!$K:$K"),INDIRECT("Current_Employees!$O:$O"),INDIRECT("Current_Employees!$P:$P")}),ISERROR(MATCH(INDIRECT("Current_Employees!$P:$P"),INDIRECT("Backup!$P:$P"),0)),'+
                            'len(INDIRECT("Current_Employees!$P:$P")),INDIRECT("Current_Employees!$O:$O")="Col Liberal Arts & Science (Col)")}},{INDIRECT("Current_Employees!$A$2:$C$2"),INDIRECT("Current_Employees!$J$2"),INDIRECT("Current_Employees!$K$2"),'+
                            'INDIRECT("Current_Employees!$O$2"),INDIRECT("Current_Employees!$P$2")})');
    
    //ssDeptFD.setFormula('=ARRAY_CONSTRAIN(ARRAYFORMULA({"Former Department";if(indirect("$A$3:$A")<>"",VLOOKUP(indirect("$A$3:$A"),indirect("Backup!$A$2:$P"),COUNTA(indirect("Backup!$A$2:$P$2")),0),"")}), 2, 1)');
    ssDeptFD.setFormula('=ARRAYFORMULA({"Former Department";if(indirect("$A$3:$A")<>"",VLOOKUP(indirect("$B$3:$B"),indirect("Backup!$B$2:$P"),COUNTA(indirect("Backup!$B$2:$P$2")),0),"")})');
    
    
    //ssDeptFDHD.setFormula('=ARRAY_CONSTRAIN(ARRAYFORMULA({"Former Dept Hiring Date";if(indirect("$A$3:$A")<>"",VLOOKUP(indirect("$A$3:$A"),indirect("Backup!$A$2:$J"),COUNTA(indirect("Backup!$A$2:$J$2")),0),"")}), 2, 1)');
    ssDeptFDHD.setFormula('=ARRAYFORMULA({"Former Dept Hiring Date";if(indirect("$A$3:$A")<>"",VLOOKUP(indirect("$B$3:$B"),indirect("Backup!$B$2:$J"),COUNTA(indirect("Backup!$B$2:$J$2")),0),"")})');
    
    
    ssFormerTime.setFormula('="Employees no longer here (In "& TEXT(INDIRECT("Backup!A1"),"mm-dd-yyyy") & " Backup, not in Current_Employees)"');
    
    ssFormerFilter.setFormula('=iferror({{INDIRECT("Backup!$A$2:$C$2"),INDIRECT("Backup!$J$2"),INDIRECT("Backup!$K$2"),INDIRECT("Backup!$O$2"),INDIRECT("Backup!$P$2")};{FILTER({INDIRECT("Backup!$A:$C"),INDIRECT("Backup!$J:$J"),INDIRECT("Backup!$K:$K"),INDIRECT("Backup!$O:$O"),INDIRECT("Backup!$P:$P")},ISERROR(MATCH(INDIRECT("Backup!$AB:$AB"),INDIRECT("Current_Employees!$AB:$AB"),0)),len(INDIRECT("Backup!$A:$A")),INDIRECT("Backup!$O:$O")="Col Liberal Arts & Science (Col)")}},{INDIRECT("Backup!$A$2:$C$2"),INDIRECT("Backup!$J$2"),INDIRECT("Backup!$K$2"),INDIRECT("Backup!$O$2"),INDIRECT("Backup!$P$2")})');
    
    var ssHistoryTime = ssHistory.getRange("A1");
    ssHistoryTime.setFormula('="Last updated: " & TEXT(INDIRECT("Current_Employees!A1"),"mm-dd-yyyy")');
    var ssHistoryFilter = ssHistory.getRange("A2");
    ssHistoryFilter.setFormula("=ARRAYFORMULA('Department Change'!A2:I2)");
    var ssHistoryChange = ssHistory.getRange("J2");
    ssHistoryChange.setFormula('="Change"');
    var ssHistoryDate = ssHistory.getRange("K2");
    ssHistoryDate.setFormula('="Date"');
    var ssPosition = ssHistory.getRange("L2").setFormula('="Position #"');
    var ssPositionFormula = ssHistory.getRange("L1").setFormula('=iferror(IFERROR(VLOOKUP({A1:C1},Current_Employees!$A:$T,COUNTA(Current_Employees!$A$2:$T$2),0),VLOOKUP({A1:C1},Backup!$A:$T,COUNTA(Backup!$A$2:$T$2),0)),"")');
    
    //Run Update functions to append data to the History changelog
    if (updater == 0) {
      ss.toast('Updating History Change Log...', 'Running update() function', 10);
      Utilities.sleep(10000);// pause for 10 seconds, troubleshooting why update() is inaccurate within function but works by itself
      update();
    } else if (updater == -1) {
      ss.toast('Updating History Change Log with new employees only...', 'Running updateNewEmployees() function', 10);
      Utilities.sleep(10000);// pause for 10 seconds, troubleshooting why update() is inaccurate within function but works by itself
      updateNewEmployees();
    }
    
    //Backup csv file to folder
    backupToCSV(archiveFolder,attachment,sheetCE);/*
    Logger.log(archiveFolder.getName());
    var attBlob = attachment.copyBlob();
    var add = archiveFolder.createFile(attBlob);
    var attName = 'FS_'+sheetCE.getRange(1, 1).getDisplayValue()+".csv";
    add.setName(attName);
    Logger.log(archiveFolder.getFiles());*/
    
  } else if (email_date.valueOf() == sheet_date.valueOf()){
    
    //We are finished
    ss.toast('The latest CSV has already been imported.', 'Complete', 10);
    Logger.log('Quitting, no backup performed, script was successful but the latest CSV has already been imported.');
    
    
  } else {
    
    //Sheet from email is not newer, so let's not do anything
    ss.toast('The latest CSV has already been imported.', 'Unsuccessful', 10);
    Logger.log('Quitting, no backup performed.');
  }
  MailApp.sendEmail({to:'rmccal14+logger@uncc.edu',subject: "FS Banner Log!",body: Logger.getLog()});
}

function backupToCSV(archiveFolder,attachment,sheetCE){
  //Backup csv file to folder
  attachment = attachment || GmailApp.search("filename:*current_employee_directory_listing.csv* has:attachment ")[0].getMessages()[0].getAttachments()[0];
  archiveFolder = archiveFolder || DriveApp.getFolderById('1Fqb901vr5CVSzyhBHHeabdMHVKSl890S');
  sheetCE = sheetCE || SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Current_Employees');
  Logger.log(archiveFolder.getName());
  var attBlob = attachment.copyBlob();
  var add = archiveFolder.createFile(attBlob);
  var attName = 'FS_'+sheetCE.getRange(1, 1).getDisplayValue()+".csv";
  add.setName(attName);
  Logger.log(archiveFolder.getFiles());    
}

//***************************************************************************************************************
//https://productforums.google.com/forum/#!msg/docs/I4sMYP3MjWU/EfgSq44c91gJ
//Append all changes to History sheet

function update(){
  
  // source doc
  var ssFS_Banner = SpreadsheetApp.getActiveSpreadsheet();
  
  // source sheets
  var ssNew = ssFS_Banner.getSheetByName('CLAS New Employee Status Change');  
  var ssDept = ssFS_Banner.getSheetByName('Department Change');
  var ssFormer = ssFS_Banner.getSheetByName('CLAS Former Employee Status Change'); 
  var ssHistory = ssFS_Banner.getSheetByName('History');  
  var sheetCE = ssFS_Banner.getSheetByName('Current_Employees');
  
  //Delete empty rows and columns so the iterators below won't get hung (plus for better visibility)
  //ssFS_Banner.toast('Deleting empty cells...', 'First the columns, then the rows', 5);
  ssFS_Banner.toast('Deleting empty cells...', 'Only deleting the rows, 10 second countdown', 5);
  Utilities.sleep(10000);// pause for 10 seconds, troubleshooting why update() is inaccurate within function but works by itself
  //removeEmptyColumns();
  //removeEmptyRows();
  
  // Get full range of data
  var lastRow = ssNew.getLastRow();
  var lastColumn = ssNew.getLastColumn();
  var lastRow2 = ssDept.getLastRow();
  var lastColumn2 = ssDept.getLastColumn();
  var lastRow3 = ssFormer.getLastRow();
  var lastColumn3 = ssFormer.getLastColumn();
  Logger.log("lastRow is " + lastRow);
  Logger.log("lastCol is " + lastColumn);
  
  var SRangeNew = ssNew.getRange(3,1,lastRow,lastColumn);
  var SRangeDept = ssDept.getRange(3,1,lastRow2,lastColumn2);
  var SRangeHistory = ssFormer.getRange(3,1,lastRow3,lastColumn3);
  
  // get the data values in range
  var sDataNew = SRangeNew.getValues();
  var sDataDept = SRangeDept.getValues();
  var sDataFormer = SRangeHistory.getValues();
  var email_date = sheetCE.getRange(1, 1).getDisplayValue();  
  var plainText;
  var updatedFormula;
  
  //!!!!--------------------------------------------------!!!!!
  //Adds the changes per category, filters by rows with '@' in them which should be all employees (email addresses), skips blank rows  
  ssFS_Banner.toast('Updating Change Log...', 'Appending employees with modified statuses to History sheet', 10);
  for (var i = 0; i < sDataNew.length; i++){
    if (sDataNew.join('').indexOf('@') > 0){
      if (sDataNew[i].join('').indexOf('@') > 0){
        ssHistory.appendRow(sDataNew[i]);
        ssHistory.getRange(ssHistory.getLastRow(), 10).setValue('New Employee');
        ssHistory.getRange(ssHistory.getLastRow(), 11).setValue(email_date); 
        ssHistory.getRange(1, 12).copyTo(ssHistory.getRange(ssHistory.getLastRow(), 12));
        //Get employee's position number, paste as plain text
        plainText = ssHistory.getRange(ssHistory.getLastRow(), 12).getDisplayValue();
        Logger.log("plainText is " + plainText);
        ssHistory.getRange(ssHistory.getLastRow(), 12).clear();
        ssHistory.getRange(ssHistory.getLastRow(), 12).setValue(plainText);
      }
    }
  }
  SpreadsheetApp.flush();
  
  for (var i = 0; i < sDataDept.length; i++){
    if (sDataDept.join('').indexOf('@') > 0){
      if (sDataDept[i].join('').indexOf('@') > 0){
        ssHistory.appendRow(sDataDept[i]);
        ssHistory.getRange(ssHistory.getLastRow(), 10).setValue('Employee Changed Departments');
        ssHistory.getRange(ssHistory.getLastRow(), 11).setValue(email_date);
        ssHistory.getRange(1, 12).copyTo(ssHistory.getRange(ssHistory.getLastRow(), 12));
        //Get employee's position number, paste as plain text
        plainText = ssHistory.getRange(ssHistory.getLastRow(), 12).getDisplayValue();
        Logger.log("plainText is " + plainText);
        ssHistory.getRange(ssHistory.getLastRow(), 12).clear();
        ssHistory.getRange(ssHistory.getLastRow(), 12).setValue(plainText);
      }
    }
  }
  SpreadsheetApp.flush();
  
  for (var i = 0; i < sDataFormer.length; i++){
    if (sDataFormer.join('').indexOf('@') > 0){
      if (sDataFormer[i].join('').indexOf('@') > 0){
        ssHistory.appendRow(sDataFormer[i]);
        ssHistory.getRange(ssHistory.getLastRow(), 10).setValue('Employee Left');
        ssHistory.getRange(ssHistory.getLastRow(), 11).setValue(email_date);
        ssHistory.getRange(1, 12).copyTo(ssHistory.getRange(ssHistory.getLastRow(), 12));
        //Get employee's position number, paste as plain text
        plainText = ssHistory.getRange(ssHistory.getLastRow(), 12).getDisplayValue();
        Logger.log("plainText is " + plainText);
        ssHistory.getRange(ssHistory.getLastRow(), 12).clear();
        ssHistory.getRange(ssHistory.getLastRow(), 12).setValue(plainText);
      }
    }
  }
  SpreadsheetApp.flush();
  MailApp.sendEmail({to:'rmccal14+logger@uncc.edu',subject: "FS Banner Log!",body: Logger.getLog()});
  
  //!!!!--------------------------------------------------!!!!!
  
}
//*******************************************************************************************************************
//*******************************************************************************************************************
//Append all changes to History sheet

function updateNewEmployees(){
  
  // source doc
  var ssFS_Banner = SpreadsheetApp.getActiveSpreadsheet();
  
  // source sheets
  var ssNew = ssFS_Banner.getSheetByName('CLAS New Employee Status Change');  
  var ssDept = ssFS_Banner.getSheetByName('Department Change');
  var ssFormer = ssFS_Banner.getSheetByName('CLAS Former Employee Status Change'); 
  var ssHistory = ssFS_Banner.getSheetByName('History');  
  var sheetCE = ssFS_Banner.getSheetByName('Current_Employees');
  
  //Delete empty rows and columns so the iterators below won't get hung (plus for better visibility)
  //ssFS_Banner.toast('Deleting empty cells...', 'First the columns, then the rows', 5);
  ssFS_Banner.toast('Deleting empty cells...', 'Only deleting the rows, 10 second countdown', 5);
  Utilities.sleep(10000);// pause for 10 seconds, troubleshooting why update() is inaccurate within function but works by itself
  //removeEmptyColumns();
  removeEmptyRows();
  
  // Get full range of data
  var lastRow = ssNew.getLastRow();
  var lastColumn = ssNew.getLastColumn();
  var lastRow2 = ssDept.getLastRow();
  var lastColumn2 = ssDept.getLastColumn();
  var lastRow3 = ssFormer.getLastRow();
  var lastColumn3 = ssFormer.getLastColumn();
  Logger.log("lastRow is " + lastRow);
  Logger.log("lastCol is " + lastColumn);
  
  var SRangeNew = ssNew.getRange(3,1,lastRow,lastColumn);
  var SRangeDept = ssDept.getRange(3,1,lastRow2,lastColumn2);
  var SRangeHistory = ssFormer.getRange(3,1,lastRow3,lastColumn3);
  
  // get the data values in range
  var sDataNew = SRangeNew.getValues();
  var sDataDept = SRangeDept.getValues();
  var sDataFormer = SRangeHistory.getValues();
  var email_date = sheetCE.getRange(1, 1).getDisplayValue();  
  var plainText;
  
  //!!!!--------------------------------------------------!!!!!
  //Adds the changes per category, filters by rows with '@' in them which should be all employees (email addresses), skips blank rows  
  ssFS_Banner.toast('Updating Change Log...', 'Appending employees with modified statuses to History sheet', 10);
  for (var i = 0; i < sDataNew.length; i++){
    if (sDataNew.join('').indexOf('@') > 0){
      if (sDataNew[i].join('').indexOf('@') > 0){
        ssHistory.appendRow(sDataNew[i]);
        ssHistory.getRange(ssHistory.getLastRow(), 10).setValue('New Employee');
        ssHistory.getRange(ssHistory.getLastRow(), 11).setValue(email_date); 
        ssHistory.getRange(1, 12).copyTo(ssHistory.getRange(ssHistory.getLastRow(), 12));
        //Get employee's position number, paste as plain text
        plainText = ssHistory.getRange(ssHistory.getLastRow(), 12).getDisplayValue();
        Logger.log("plainText is " + plainText);
        ssHistory.getRange(ssHistory.getLastRow(), 12).clear();
        ssHistory.getRange(ssHistory.getLastRow(), 12).setValue(plainText);
      }
    }
  }

SpreadsheetApp.flush();

for (var i = 0; i < sDataDept.length; i++){
  if (sDataDept.join('').indexOf('@') > 0){
    if (sDataDept[i].join('').indexOf('@') > 0){
      ssHistory.appendRow(sDataDept[i]);
      ssHistory.getRange(ssHistory.getLastRow(), 10).setValue('Employee Changed Departments');
      ssHistory.getRange(ssHistory.getLastRow(), 11).setValue(email_date);
      ssHistory.getRange(1, 12).copyTo(ssHistory.getRange(ssHistory.getLastRow(), 12));
      //Get employee's position number, paste as plain text
      plainText = ssHistory.getRange(ssHistory.getLastRow(), 12).getDisplayValue();
      Logger.log("plainText is " + plainText);
      ssHistory.getRange(ssHistory.getLastRow(), 12).clear();
      ssHistory.getRange(ssHistory.getLastRow(), 12).setValue(plainText);
    }
  }
}

SpreadsheetApp.flush();

}
//*******************************************************************************************************************
//*******************************************************************************************************************

//Delete empty columns

function removeEmptyColumns() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var allsheets = ss.getSheets();
  for (var s in allsheets){
    var sheet=allsheets[s];
    var maxColumns = sheet.getMaxColumns(); 
    var lastColumn = sheet.getLastColumn();
    if (maxColumns-lastColumn != 0){
      sheet.deleteColumns(lastColumn+1, maxColumns-lastColumn);
    }
  }
}

//Delete empty rows
function removeEmptyRows() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var allsheets = ss.getSheets();
  for (var s in allsheets){
    var sheet=allsheets[s];
    var maxRows = sheet.getMaxRows(); 
    var lastRow = sheet.getLastRow();
    Logger.log("sheet is "+sheet.getName());
    if (maxRows-lastRow > 1){
      //try{
      sheet.deleteRows(lastRow+1, maxRows-lastRow);
      /*} catch (e) {
      Logger.log(e);
      }*/
    }
  }
}

//*******************************************************************************************************************
//*******************************************************************************************************************
//*******************************************************************************************************************
