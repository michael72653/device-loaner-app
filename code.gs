function doGet(e) {

  // DEFAULT TO HOME PAGE IF NO PARAMETERS ARE FOUND...
  if(!e.parameters.page) {

      return  HtmlService.createTemplateFromFile("deployed").evaluate()
                          .setTitle("BHN/BHS Loaner Deployment")
                          .addMetaTag("viewport", "width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no")
                          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } else {
  
      // IF PARAMETERS ARE FOUND, GO HERE...
      return  HtmlService.createTemplateFromFile(e.parameters['page']).evaluate()
                          .setTitle("BHN/BHS Loaner Deployment")
                          .addMetaTag("viewport", "width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no")
                          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }; // end check 

}; //end doGet function


function GetScriptURL() {
  return ScriptApp.getService().getUrl();
};


function IncludeFile(fn) {
  return HtmlService.createHtmlOutputFromFile(fn).getContent();
};

function AppTitle() {
  return 'BHS Loaner Deployment';
};


function URLFooter() {
  return 'Best viewed in full wide screen desktop mode [1920 x 1080px @ 100% zoom]';
};


function SubTitle() {
  return 'Entries will be deleted once loaner is returned.';
};


/*
==============================================================================================================================
*/
const CUSTID = "C04f1hg9m";  //"customerId"

const SHEETNAME     = "BHS_LOANER_CHECKOUT";
const STAFFSHEET    = "BHS_STAFF_LOANER_CKOUT";
const FUNCSHEET     = "CHOICES";
const VARSSHEET     = "VARIABLES";
const LOGSSHEET     = "LOANER_HISTORY_LOG";
const EXCLUDESHEET  = "NO_LOANER_DEPLOYS";
const REPAIRSHEET   = "LOANER_FOR_REPAIRS";

const date = new Date();
const currentDay = date.toLocaleString('en-US', { month:'numeric', day:'numeric', year:'numeric' });

///////////////////////////////////////////////// copied code /////////////////////////////////////////////////////////////////

let spsh = SpreadsheetApp.getActiveSpreadsheet();
let actsh = spsh.getSheetByName(VARSSHEET);

/////////////////////////// variables ////////////////////////////////
///////// to pull variable values from internal spreadsheet //////////
/////// THESE CONSTANTS MUST MATCH ON THE DATA SHEET AS BELOW! ///////

const FOLDERID      = actsh.getRange(2,2).getValue();
const SHEETID       = actsh.getRange(3,2).getValue();
const UNIQUEID      = actsh.getRange(4,2).getValue();
const DATE          = actsh.getRange(5,2).getValue();
const STUID         = actsh.getRange(6,2).getValue();
const STULASTNAME   = actsh.getRange(7,2).getValue();
const STUFIRSTNAME  = actsh.getRange(8,2).getValue();
const LOANERNO      = actsh.getRange(9,2).getValue();
const SERIALNO      = actsh.getRange(10,2).getValue();
const REASON        = actsh.getRange(11,2).getValue();
const RETURNED      = actsh.getRange(12,2).getValue(); 
const EMAILED       = actsh.getRange(13,2).getValue();
const REPAIRCOL     = actsh.getRange(14,2).getValue();
const STUEMAIL      = actsh.getRange(15,2).getValue();
const STAFFEMAIL    = actsh.getRange(16,2).getValue(); 
const NUMOFCOLS     = actsh.getRange(17,2).getValue();
const STAFFNUMOFCOLS= actsh.getRange(18,2).getValue();
const NUMOFDAYSOLD  = actsh.getRange(19,2).getValue();
const EMAILON       = actsh.getRange(20,2).getValue();
const EMAILRECEPTS  = actsh.getRange(21,2).getValue();
const EMAILADMIN    = actsh.getRange(22,2).getValue();    
const HISTNUMOFCOLS = actsh.getRange(23,2).getValue();    
const EXCNUMOFCOLS  = actsh.getRange(24,2).getValue();    
const DUPLSHEETID   = actsh.getRange(25,2).getValue();    
const TECHTICKETNO  = actsh.getRange(26,2).getValue();    
const FEESPAID      = actsh.getRange(27,2).getValue();    
const TIMEEXPIRED   = actsh.getRange(28,2).getValue();    
const REPAIRNUMOFCOLS = actsh.getRange(29,2).getValue();    
const REPAIRNUMOFDAYS = actsh.getRange(30,2).getValue();
const EXPIRATIONDATE  = actsh.getRange(31,2).getValue();
const REPAIRSURL      = actsh.getRange(32,2).getValue();     

/////////////////////////////////////////////////////   FUNCTIONS   ///////////////////////////////////////////////////////////////


/* ===============================================   CHOICES FUNCTIONS   ========================================================*/

function LoanerChoices() {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(FUNCSHEET);

  SpreadsheetApp.setActiveSheet(sh);

  var choices = sh.getRange(2,1,sh.getLastRow(),1).getValues().filter(r => r.every(Boolean));  // [filter] is to ignore blank cells
  var options = '';

  for(var i=0; i < choices.length; i++) {
    options += '<option value="' + choices[i] + '">' + choices[i] + '</option>\n';
  }

  return options;

}; // end


function SerialNoChoices() {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(FUNCSHEET);

  SpreadsheetApp.setActiveSheet(sh);

  var choices = sh.getRange(2,2,sh.getLastRow(),1).getValues().filter(r => r.every(Boolean));  // [filter] is to ignore blank cells
  var options = '';

  for(var i=0; i < choices.length; i++) {
    options += '<option value="' + choices[i] + '">' + choices[i] + '</option>\n';
  }

  return options;

}; // end


function ReasonChoices() {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(FUNCSHEET);

  SpreadsheetApp.setActiveSheet(sh);

  var choices = sh.getRange(2,3,sh.getLastRow(),1).getValues().filter(r => r.every(Boolean));  // [filter] is to ignore blank cells
  var options = '';

  for(var i=0; i < choices.length; i++) {
    options += '<option value="' + choices[i] + '">' + choices[i] + '</option>\n';
  }

  return options;

}; // end


function StaffLoanerChoices() {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(FUNCSHEET);

  SpreadsheetApp.setActiveSheet(sh);

  var choices = sh.getRange(2,4,sh.getLastRow(),1).getValues().filter(r => r.every(Boolean));  // [filter] is to ignore blank cells
  var options = '';

  for(var i=0; i < choices.length; i++) {
    options += '<option value="' + choices[i] + '">' + choices[i] + '</option>\n';
  }

  return options;

}; // end


function StaffSerialNoChoices() {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(FUNCSHEET);

  SpreadsheetApp.setActiveSheet(sh);

  var choices = sh.getRange(2,5,sh.getLastRow(),1).getValues().filter(r => r.every(Boolean));  // [filter] is to ignore blank cells
  var options = '';

  for(var i=0; i < choices.length; i++) {
    options += '<option value="' + choices[i] + '">' + choices[i] + '</option>\n';
  }

  return options;

}; // end


function StaffReasonChoices() {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(FUNCSHEET);

  SpreadsheetApp.setActiveSheet(sh);

  var choices = sh.getRange(2,6,sh.getLastRow(),1).getValues().filter(r => r.every(Boolean));  // [filter] is to ignore blank cells
  var options = '';

  for(var i=0; i < choices.length; i++) {
    options += '<option value="' + choices[i] + '">' + choices[i] + '</option>\n';
  }

  return options;

}; // end


function StudentRepairPayChoices() {

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(FUNCSHEET);

    SpreadsheetApp.setActiveSheet(sh);

    var choices = sh.getRange(2,7,sh.getLastRow(),1).getValues().filter(r => r.every(Boolean));  // [filter] is to ignore blank cells
    var options = '';

    for(var i=0; i < choices.length; i++) {
      options += '<option value="' + choices[i] + '"';
      options += (i==0) ? ' selected>':'>';  // create the first value to be the default selection when form is loaded. 
      options += choices[i] + '</option>\n';
    }

    return options;
     
}; //end


/* =========================================================================================================================================== */


/* ================================================    CUSTOM FUNCTIONS   =====================================================================*/


function SubmitEntry(student_id, last_name, first_name, loaner_no, serial_no, reason, email, paid) {
  
  //DEFINE ALL ACTIVE SHEETS
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  //DEFINE MAIN SHEET          
  var ms = ss.getSheetByName(SHEETNAME);
  var es = ss.getSheetByName(EXCLUDESHEET); // exclude list

  //DEFINE LAST ROW
  var lastRow = ms.getLastRow();
  var elastRow = es.getLastRow(); // exclude list
  
  //Define Return Variables
  var return_date = '';
  var error = 'SUCCESS';
  var return_array = [];

  ////////////////////////////////////////////  CHECK FOR EXCLUSIONS! ///////////////////////////////////////////////////

  for(var k=2; k <= elastRow; k++) {

    if( student_id == es.getRange(k, STUID).getValue() ) {

          // only check by ID instead of name & ID 
          first_name = es.getRange(k, STUFIRSTNAME).getValue(); 
          last_name = es.getRange(k, STULASTNAME).getValue();
          
          error = 'Student NOT allowed to be assigned a loaner!';
          return_array.push([error, return_date, first_name, last_name, student_id]);
          
          return return_array;

    }; // end if

  }; // end for loop
  
  //////////////////////////////////////////  CHECK FOR ANY DUPLICATES! /////////////////////////////////////////////////
  
  for(var j=2; j <= lastRow; j++) {
    
    if( first_name == ms.getRange(j, STUFIRSTNAME).getValue() 
          && last_name == ms.getRange(j, STULASTNAME).getValue()
              && student_id == ms.getRange(j, STUID).getValue() ) {
      
          error = 'Duplicate Record Found:<br>Student has already checked out a loaner!';
          return_array.push([error, return_date, first_name, last_name, student_id]);
          
          return return_array;
    
    }; // end if 

    if( loaner_no == ms.getRange(j, LOANERNO).getValue() 
          && ms.getRange(j, RETURNED).getValue() == 'no' ) {

          error = 'Loaner already checked out.<br>Please select another loaner.';
          return_array.push([error, return_date, first_name, last_name, student_id]);

          return return_array;

    }; // end if
    
  }; // end for loop

  ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
  
  if(error == 'SUCCESS') {

        var unique_id = (Math.floor(Math.random() * 2) + Date.now()).toString(36).slice(-5);
        return_date = GetDate(date);
        var datenotime = GetDateNoTime(date);

        // pull value from functions sheet
        var rf = ss.getSheetByName(FUNCSHEET);
        
        // for repairs list - pulling dropdown value - 2nd value [3] in the invoice [G] list
        const CHECKREPAIRVALUE = rf.getRange(3,7).getValue();

        // set icon in database
        var setrepairvalue = ( CHECKREPAIRVALUE == paid ) ? '<span class="mif-checkmark fg-green"></span>' 
                                                          : '<span class="mif-cross fg-red"></span>'; 
        
        // ADD RECORD
        ms.getRange(lastRow + 1, UNIQUEID).setValue(unique_id);
        ms.getRange(lastRow + 1, DATE).setValue(return_date).setNumberFormat("MM/dd/yy hh:mm am/pm");
        ms.getRange(lastRow + 1, STUID).setValue(student_id);
        ms.getRange(lastRow + 1, STULASTNAME).setValue(last_name);
        ms.getRange(lastRow + 1, STUFIRSTNAME).setValue(first_name);
        ms.getRange(lastRow + 1, LOANERNO).setValue(loaner_no);
        ms.getRange(lastRow + 1, SERIALNO).setValue(serial_no);
        ms.getRange(lastRow + 1, STUEMAIL).setValue(email);
        ms.getRange(lastRow + 1, REASON).setValue(reason);
        ms.getRange(lastRow + 1, RETURNED).setValue('no');
        ms.getRange(lastRow + 1, REPAIRCOL).setValue(setrepairvalue);
        ms.getRange(lastRow + 1, EMAILED).setValue('no');

        // Send Email out
        if(EMAILON) {
            // sendEmail(recipient, subject, body, options)   
            MailApp.sendEmail(EMAILRECEPTS, `NEW BHS Student Loaner Issued - ${datenotime}`, 'Hello!', 
                              { cc: EMAILADMIN,
                                name: 'Michael Stapleton',
                                noReply: true,
                                htmlBody: `[This is an auto-generated message]<br>
                                          ====================================<br>
                                          <font size="+1">Loaner was issued out to:<br>
                                          Student Name: <strong>${first_name} ${last_name} - ${student_id}</strong><br>
                                          CB Serial No.: <strong>${serial_no}</strong><br>
                                          Date/Time Issued: <strong>${return_date}</strong></font>` 
                              });
            var emailQuotaRemaining = MailApp.getRemainingDailyQuota();
            Logger.log("Remaining email quota: " + emailQuotaRemaining);
        }; // end email on

          ////////////////////////////    add to History logs sheet    ////////////////////////////////////

                    var hs = ss.getSheetByName(LOGSSHEET);
                    var lr = hs.getLastRow(); 

                    hs.getRange(lr + 1, UNIQUEID).setValue(unique_id);
                    hs.getRange(lr + 1, DATE).setValue(return_date).setNumberFormat("MM/dd/yy hh:mm am/pm");
                    hs.getRange(lr + 1, STUID).setValue(student_id);
                    hs.getRange(lr + 1, STULASTNAME).setValue(last_name);
                    hs.getRange(lr + 1, STUFIRSTNAME).setValue(first_name);
                    hs.getRange(lr + 1, LOANERNO).setValue(loaner_no);
                    hs.getRange(lr + 1, SERIALNO).setValue(serial_no);
                    hs.getRange(lr + 1, REASON).setValue(reason);
                    hs.getRange(lr + 1, RETURNED).setValue('-no-');
                    hs.getRange(lr + 1, STUEMAIL-1).setValue(email);  //adjust to fit with repairs module
                    
          /////////////////////////////////////////////////////////////////////////////////////////////////

          ///////////////////////////////////////////   TO ADD TO THE REPAIRS LIST   /////////////////////////////////////////////////
            
                    var rs = ss.getSheetByName(REPAIRSHEET);
                    var llr = rs.getLastRow(); 

                    // create 2-week future date
                    var expire_date = new Date(date.getTime() + 14 * 24 * 3600 * 1000);  // 2 wks out * 24 * 60 * 60 * 1000;
                    
                    // add to repairs list if true
                    if(CHECKREPAIRVALUE == paid) {

                          rs.getRange(llr + 1, UNIQUEID).setValue(unique_id);
                          rs.getRange(llr + 1, DATE).setValue(return_date).setNumberFormat("MM/dd/yy hh:mm am/pm");
                          rs.getRange(llr + 1, STUID).setValue(student_id);
                          rs.getRange(llr + 1, STULASTNAME).setValue(last_name);
                          rs.getRange(llr + 1, STUFIRSTNAME).setValue(first_name);
                          rs.getRange(llr + 1, TECHTICKETNO).setValue("pending");
                          rs.getRange(llr + 1, FEESPAID).setValue("no");
                          rs.getRange(llr + 1, TIMEEXPIRED).setValue("no");
                          rs.getRange(llr + 1, EXPIRATIONDATE).setValue(expire_date).setNumberFormat("MM/dd/yy hh:mm am/pm");

                    }; // end 

          /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////  

  }; // end

  return_array.push([error, return_date, first_name, last_name, student_id, loaner_no, serial_no, email]);
  
  return return_array;
  
}; // end function


/* ========================================================================================================================= */


function StaffSubmitEntry(room_no, email, last_name, first_name, loaner_no, serial_no, reason) {
  
  //DEFINE ALL ACTIVE SHEETS
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  //DEFINE MAIN SHEET          
  var ms = ss.getSheetByName(STAFFSHEET);
  var es = ss.getSheetByName(EXCLUDESHEET); // exclude list

  //DEFINE LAST ROW
  var lastRow = ms.getLastRow();
  var elastRow = es.getLastRow(); // exclude list
  
  //Define Return Variables
  var return_date = '';
  var error = 'SUCCESS';
  var return_array = [];

  ////////////////////////////////////////////  CHECK FOR EXCLUSIONS! ///////////////////////////////////////////////////

  for(var k=2; k <= elastRow; k++) {

    if( email == es.getRange(k, STAFFEMAIL).getValue() ) {

          first_name = es.getRange(k, STUFIRSTNAME).getValue(); 
          last_name = es.getRange(k, STULASTNAME).getValue();
          room_no = es.getRange(k, STUID).getValue();
          
          error = 'Staff member NOT allowed to be assigned a loaner!';
          return_array.push([error, return_date, first_name, last_name, email]);
          
          return return_array;

    }; // end if

  }; // end for loop
  
  //////////////////////////////////////////  CHECK FOR ANY DUPLICATES! /////////////////////////////////////////////////
  
  for(var j=2; j <= lastRow; j++) {
    
    if( first_name == ms.getRange(j, STUFIRSTNAME).getValue() 
          && last_name == ms.getRange(j, STULASTNAME).getValue()
              && email == ms.getRange(j, STAFFEMAIL).getValue() ) {
      
          error = 'Duplicate Record Found:<br>Staff member has already checked out a loaner!';
          return_array.push([error, return_date, first_name, last_name, email]);
          
          return return_array;
    
    }; // end if 

    if( loaner_no == ms.getRange(j, LOANERNO).getValue() 
          && ms.getRange(j, RETURNED).getValue() == 'no' ) {

          error = 'Loaner already checked out.<br>Please select another loaner.';
          return_array.push([error, return_date, first_name, last_name, email]);

          return return_array;

    }; // end if
    
  }; // end for loop

  ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
  
  if(error == 'SUCCESS') {

        var unique_id = (Math.floor(Math.random() * 2) + Date.now()).toString(36).slice(-5);
        return_date = GetDate(date);
        var datenotime = GetDateNoTime(date);
        
        // ADD RECORD
        ms.getRange(lastRow + 1, UNIQUEID).setValue(unique_id);
        ms.getRange(lastRow + 1, DATE).setValue(return_date).setNumberFormat("MM/dd/yy hh:mm am/pm");
        ms.getRange(lastRow + 1, STUID).setValue(room_no);
        ms.getRange(lastRow + 1, STULASTNAME).setValue(last_name);
        ms.getRange(lastRow + 1, STUFIRSTNAME).setValue(first_name);
        ms.getRange(lastRow + 1, LOANERNO).setValue(loaner_no);
        ms.getRange(lastRow + 1, SERIALNO).setValue(serial_no);
        ms.getRange(lastRow + 1, REASON).setValue(reason);
        ms.getRange(lastRow + 1, RETURNED).setValue('no');
        ms.getRange(lastRow + 1, EMAILED).setValue('no');
        ms.getRange(lastRow + 1, STAFFEMAIL).setValue(email);

        // Send Email out
        if(EMAILON) {
            // sendEmail(recipient, subject, body, options)   
            MailApp.sendEmail(EMAILRECEPTS, `NEW BHS Staff Loaner Issued - ${datenotime}`, 'Hello!', 
                              { cc: EMAILADMIN,
                                name: 'Michael Stapleton',
                                noReply: true,
                                htmlBody: `[This is an auto-generated message]<br>
                                          ====================================<br>
                                          <font size="+1">Loaner was issued out to:<br>
                                          Staff Name: <strong>${first_name} ${last_name} - ${email}</strong><br>
                                          CB Serial No.: <strong>${serial_no}</strong><br>
                                          Date/Time Issued: <strong>${return_date}</strong></font>` 
                              });
            var emailQuotaRemaining = MailApp.getRemainingDailyQuota();
            Logger.log("Remaining email quota: " + emailQuotaRemaining);
        }; // end email on

          ////////////////////////////    add to History logs sheet    ////////////////////////////////////

                    var hs = ss.getSheetByName(LOGSSHEET);
                    var lr = hs.getLastRow(); 

                    hs.getRange(lr + 1, UNIQUEID).setValue(unique_id);
                    hs.getRange(lr + 1, DATE).setValue(return_date).setNumberFormat("MM/dd/yy hh:mm am/pm");
                    hs.getRange(lr + 1, STUID).setValue(room_no);
                    hs.getRange(lr + 1, STULASTNAME).setValue(last_name);
                    hs.getRange(lr + 1, STUFIRSTNAME).setValue(first_name);
                    hs.getRange(lr + 1, LOANERNO).setValue(loaner_no);
                    hs.getRange(lr + 1, SERIALNO).setValue(serial_no);
                    hs.getRange(lr + 1, REASON).setValue(reason);
                    hs.getRange(lr + 1, RETURNED).setValue('-no-');
                    hs.getRange(lr + 1, STAFFEMAIL).setValue(email);

          /////////////////////////////////////////////////////////////////////////////////////////////////

  }; // end

  return_array.push([error, return_date, first_name, last_name, room_no, email, loaner_no, serial_no]);
  
  return return_array;
  
}; // end function


/* =========================================================================================================================================== */


function ListHeaders() {

    // DEFINE ALL ACTIVE SHEETS
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // DEFINE MAIN SHEET          
    var ms = ss.getSheetByName(SHEETNAME);

    // GET DATA
    var data = ms.getRange(1, 3, 1, ms.getMaxColumns()-2).getValues();  // to hide ID & DATE fields, add [3] & [-2]
        data = data.toString().toUpperCase().split(",");

    // EXPORT DATA
    var titleHTML = '<thead><tr>';
    
    for(var k=0, y=1; k < data.length; k++, y++) {
      
      titleHTML += ((y % NUMOFCOLS) == 0) ? '<th class="text-center">' + data[k] + '</th>'
                                              + '<th class="text-center">LATE ACTION</th></tr><tr>' 
                                          : '<th class="text-center">' + data[k] + '</th>'; 

    };

    return titleHTML + '</tr></thead>';

}; // end function


function ListOfStudents() {
    
    // DEFINE ALL ACTIVE SHEETS
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // DEFINE MAIN SHEET          
    var ms = ss.getSheetByName(SHEETNAME);

    // CHECK IF DATASHEET IS BLANK 
    if(ms.getLastRow()-1 < 1) { 
                              var colspan = NUMOFCOLS; colspan++;
                              return '<tbody><tr><td colspan="'+ colspan +'">'
                                      + '<span class="mif-checkmark fg-green"></span>'
                                      + '&nbsp;All loaners are checked in!&nbsp;'
                                      + '<span class="mif-checkmark fg-green"></span>'
                                      + '</td></tr></tbody>' };

    // GET DATA
    var data = ms.getRange(2, 3, ms.getLastRow()-1, ms.getMaxColumns()-2).getValues();  // to hide ID & DATE fields, add [3] & [-2]
    
    // CONVERT TIMES & SPLIT DATA   --- useGrouping removes commas from numeric values [such as ID]----
    data = data.toLocaleString("en-US", { month:"numeric", day:"numeric", year:"numeric", useGrouping:false }).split(","); 

    // EXPORT DATA
    var listHTML = '<tbody><tr>';
    
    for(var j=0, x=1; j < data.length; j++, x++) {

        var email_field   = data[j-0],
            emailed_field = data[j-2],
            serial_field  = data[j-5],
            loaner_field  = data[j-6],
            stuid_field   = data[j-9];
      
        listHTML += '<td>' + data[j] + '</td>';
        
        if( (x % NUMOFCOLS) == 0 ) {

              listHTML += '<td>';  
              
              listHTML += ( emailed_field == "yes" ) ? '<span class="mif-checkmark fg-green"></span>'
                                                      : '<button type="submit" class="button info small" '
                                                          + 'value="' + email_field + '|' + stuid_field + '|' + loaner_field + '|' + serial_field + '" ' 
                                                          + 'onclick="return confirm(\'Are you sure you want to send email to client?\')'
                                                          + '?ProcessEmail(this.value):alert(\'Email cancelled\');">EMAIL</button>';

              listHTML += '</td></tr><tr>';
          
        }; // end if remainder

    };  // end loop

    return listHTML + '</tr></tbody>';

}; // end function


/* =========================================================================================================================================== */


function ListStaffHeaders() {

    // DEFINE ALL ACTIVE SHEETS
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // DEFINE MAIN SHEET          
    var ms = ss.getSheetByName(STAFFSHEET);

    // GET DATA
    var data = ms.getRange(1, 3, 1, ms.getMaxColumns()-2).getValues();  // to hide ID & DATE fields, add [3] & [-2]
        data = data.toString().toUpperCase().split(",");

    // EXPORT DATA
    var titleHTML = '<thead><tr>';
    
    for(var k=0, y=1; k < data.length; k++, y++) {
      
      titleHTML += ((y % STAFFNUMOFCOLS) == 0) ? '<th class="text-center">' + data[k] + '</th>'
                                              + '<th class="text-center">LATE ACTION</th></tr><tr>' 
                                          : '<th class="text-center">' + data[k] + '</th>'; 

    };

    return titleHTML + '</tr></thead>';

}; // end function


function ListOfStaff() {
    
    // DEFINE ALL ACTIVE SHEETS
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // DEFINE MAIN SHEET          
    var ms = ss.getSheetByName(STAFFSHEET);

    // CHECK IF DATASHEET IS BLANK 
    if(ms.getLastRow()-1 < 1) { 
                              var colspan = STAFFNUMOFCOLS; colspan++;
                              return '<tbody><tr><td colspan="'+ colspan +'">'
                                      + '<span class="mif-checkmark fg-green"></span>'
                                      + '&nbsp;All loaners are checked in!&nbsp;'
                                      + '<span class="mif-checkmark fg-green"></span>'
                                      + '</td></tr></tbody>' };

    // GET DATA
    var data = ms.getRange(2, 3, ms.getLastRow()-1, ms.getMaxColumns()-2).getValues();  // to hide ID & DATE fields, add [3] & [-2]
    
    // CONVERT TIMES & SPLIT DATA   --- useGrouping removes commas from numeric values [such as ID]----
    data = data.toLocaleString("en-US", { month:"numeric", day:"numeric", year:"numeric", useGrouping:false }).split(","); 

    // EXPORT DATA
    var listHTML = '<tbody><tr>';
    
    for(var j=0, x=1; j < data.length; j++, x++) {

        var email_field   = data[j-0],
            emailed_field = data[j-1],
            serial_field  = data[j-4],
            loaner_field  = data[j-5];
      
        listHTML += '<td>' + data[j] + '</td>';
        
        if( (x % STAFFNUMOFCOLS) == 0 ) {

              listHTML += '<td>';
              
              listHTML += ( emailed_field == "yes" ) ? '<span class="mif-checkmark fg-green"></span>'
                                                      : '<button type="button" class="button info small" '
                                                          + 'value="' + email_field + '|' + loaner_field + '|' + serial_field + '" ' 
                                                          + 'onclick="return confirm(\'Are you sure you want to send email to client?\')'
                                                          + '?ProcessEmail(this.value):alert(\'Email cancelled\');">EMAIL</button>';

              listHTML += '</td></tr><tr>';
          
        }; // end if remainder

    };  // end loop

    return listHTML + '</tr></tbody>';

}; // end function


/* =========================================================================================================================================== */


function UpdateSerialNo(e) {

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var ms = ss.getSheetByName(FUNCSHEET);

    // GET DATA 
    var e = parseInt(e).toFixed();  // convert string to integer/numbers
        e = (e < 10) ? e.replace(/^0+/, '') : e;    // remove leading zeros
    var f = ms.getRange(++e, 2).getValues();  // add 1 to the [e] variable to bypass header on database

    return f;

}; // end function


function UpdateStaffSerialNo(e) {

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var ms = ss.getSheetByName(FUNCSHEET);

    // GET DATA 
    var e = parseInt(e).toFixed();  // convert string to integer/numbers
        e = (e < 10) ? e.replace(/^0+/, '') : e;    // remove leading zeros
    var f = ms.getRange(++e, 5).getValues();  // add 1 to the [e] variable to bypass header on database

    return f;

}; // end function


/* ========================================================================================================================================= */ 


function SendLateEmail(e) {

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var ms = ss.getSheetByName(SHEETNAME);

    /* PULL DATE FRON const STATED ABOVE */

    // GET DATA
    var name = e.toString().split("|");
    var error = 'ERROR';
    var return_array = [];
    
    // DEFINE LAST ROW
    var lastRow = ms.getLastRow();

    // DEFINE NAME AFTER SPLIT IN EXACT ORDER
    for(var x=0; x < name.length; x++) {
        var stuemail  = name[x-3];
        var stuid     = name[x-2];
        var loanerno  = name[x-1];
        var serial    = name[x];
    };
    
    // CHECK FOR RECORD & UPDATE
    for(var j=2; j <= lastRow; j++) {
      
      if( stuid == ms.getRange(j, STUID).getValue() 
              && stuemail == ms.getRange(j, STUEMAIL).getValue() 
                    && serial == ms.getRange(j, SERIALNO).getValue() ) {
      
          error = 'SUCCESS';
          
          ms.getRange(j, EMAILED).setValue('yes');

          // SEND EMAIL TO STUDENT //
          // sendEmail(recipient, subject, body, options)   
          MailApp.sendEmail(stuemail, `Chromebook Loaner was NOT returned!`, 'Hello!', 
                            { bcc: EMAILADMIN,
                              name: 'Michael Stapleton',
                              noReply: true,
                              replyTo: 'mstapleton@bentonvillek12.org',
                              htmlBody: `[This is an auto-generated message]<br>
                                        ====================================<br>
                                        <font size="+1"><strong>CB Info:</strong><br>
                                        Loaner #: ${loanerno}<br>
                                        Serial No: ${serial} <br>
                                        ====================================<br>
                                        Hello!<br><br>
                                        According to our records shown above, you had checked out a loaner from the BHS tech office, 
                                        but did not return it back to us by the end of the school day as instructed.  
                                        Therefore, the loaner and your assigned Chromebook both have been disabled until the loaner has been returned. 
                                        <br><br>
                                        Please do not hesitate in returning the loaner by the next school day. <strong><u>If failed to return the loaner 
                                        within 2 school days from the date of checkout, a referral will be sent to the Dean's Office for disciplinary action 
                                        &amp; you will lose all privileges from checking out another loaner.</u></strong>
                                        <br><br>
                                        Thanks!<br>
                                        Michael Stapleton<br>
                                        BHS Tech Support Specialist</font>` 
                            });
          var emailQuotaRemaining = MailApp.getRemainingDailyQuota();
          Logger.log("Remaining email quota: " + emailQuotaRemaining);

              /////////////////////////// add to History logs sheet ///////////////////////////////////////
                
              /////////////////////////////////////////////////////////////////////////////////////////////
    
      }; // end if      
      
    }; // end for loop

    return_array.push([error, stuid, stuemail, loanerno, serial]);

    return return_array;

}; // end function


function SendLateEmailStaff(e) {

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var ms = ss.getSheetByName(STAFFSHEET);

    /* PULL DATE FRON const STATED ABOVE */

    // GET DATA
    var name = e.toString().split("|");
    var error = 'ERROR';
    var return_array = [];
    
    // DEFINE LAST ROW
    var lastRow = ms.getLastRow();

    // DEFINE NAME AFTER SPLIT IN EXACT ORDER
    for(var x=0; x < name.length; x++) {
        var stuemail  = name[x-2];
        var loanerno  = name[x-1];
        var serial    = name[x];
    };
    
    // CHECK FOR RECORD & UPDATE
    for(var j=2; j <= lastRow; j++) {
      
      if( stuemail == ms.getRange(j, STAFFEMAIL).getValue() 
              && serial == ms.getRange(j, SERIALNO).getValue() ) {
      
          error = 'SUCCESS';
          
          ms.getRange(j, EMAILED).setValue('yes');

          // SEND EMAIL TO STUDENT //
          // sendEmail(recipient, subject, body, options)   
          MailApp.sendEmail(stuemail, `Staff Loaner was NOT returned!`, 'Hello!', 
                            { bcc: EMAILADMIN,
                              name: 'Michael Stapleton',
                              noReply: true,
                              replyTo: 'mstapleton@bentonvillek12.org',
                              htmlBody: `[This is an auto-generated message]<br>
                                        ====================================<br>
                                        <font size="+1"><strong>LT Info:</strong><br>
                                        Loaner #: ${loanerno}<br>
                                        Serial No: ${serial} <br>
                                        ====================================<br>
                                        Hello!<br><br>
                                        According to our records shown above, you had checked out a loaner from the BHS tech office, 
                                        but did not return it back to us by the end of the school day as instructed.  
                                        <br>
                                        Please do not hesitate in returning the loaner by the next school day.
                                        <br><br>
                                        Thanks!<br>
                                        Michael Stapleton<br>
                                        BHS Tech Support Specialist</font>` 
                            });
          var emailQuotaRemaining = MailApp.getRemainingDailyQuota();
          Logger.log("Remaining email quota: " + emailQuotaRemaining);

              /////////////////////////// add to History logs sheet ///////////////////////////////////////
                
              /////////////////////////////////////////////////////////////////////////////////////////////
    
      }; // end if      
      
    }; // end for loop

    return_array.push([error, stuemail, loanerno, serial]);

    return return_array;

}; // end function


/* =========================================================================================================================================== */


function AddReturnEntry(serialno) {

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var ms = ss.getSheetByName(SHEETNAME);

    /* PULL DATE FROM const STATED ABOVE */

    // GET DATA
    var error = 'ERROR';
    var return_array = [];
    
    // DEFINE LAST ROW
    var lastRow = ms.getLastRow();
    
    ///////////////////////////////////// CHECK FOR RECORD & UPDATE  /////////////////////////////////////////
    
    for(var j=2; j <= lastRow; j++) {
      
      if( serialno == ms.getRange(j, SERIALNO).getValue() && ms.getRange(j, RETURNED).getValue() == 'no') {
      
          error = 'SUCCESS';
          
          ms.getRange(j, RETURNED).setValue('yes');
          
          var fname = ms.getRange(j, STUFIRSTNAME).getValue();
          var lname = ms.getRange(j, STULASTNAME).getValue();
          var stuid = ms.getRange(j, STUID).getValue();

          // Send Email out
          if(EMAILON) {
              // sendEmail(recipient, subject, body, options)   
              MailApp.sendEmail(EMAILRECEPTS, `BHS Student Loaner Returned - ${currentDay}`, 'Hello!', 
                                { cc: EMAILADMIN,
                                  name: 'Michael Stapleton',
                                  noReply: true,
                                  htmlBody: `[This is an auto-generated message]<br>
                                            ====================================<br>
                                            <font size="+1">Chromebook Loaner was returned:<br>
                                            Student Name: <strong>${fname} ${lname} - ${stuid}</strong><br>
                                            CB Serial No.: <strong>${serialno}</strong><br>
                                            Date Returned: <strong>${currentDay}</strong></font>` 
                                });
              var emailQuotaRemaining = MailApp.getRemainingDailyQuota();
              Logger.log("Remaining email quota: " + emailQuotaRemaining);
          }; // end email on

              /////////////////////////////// add to History logs sheet ///////////////////////////////////
                var hs = ss.getSheetByName(LOGSSHEET);
                var lr = hs.getLastRow(); 

                var unique_id = (Math.floor(Math.random() * 2) + Date.now()).toString(36).slice(-5);
                return_date = GetDate(date);

                hs.getRange(lr + 1, UNIQUEID).setValue(unique_id);
                hs.getRange(lr + 1, DATE).setValue(return_date).setNumberFormat("MM/dd/yy hh:mm am/pm");
                hs.getRange(lr + 1, STUID).setValue(stuid);
                hs.getRange(lr + 1, STULASTNAME).setValue(lname);
                hs.getRange(lr + 1, STUFIRSTNAME).setValue(fname);
                /*hs.getRange(lr + 1, LOANERNO).setValue(loaner_no);*/
                hs.getRange(lr + 1, SERIALNO).setValue(serialno);
                /*hs.getRange(lr + 1, REASON).setValue(reason);*/
                hs.getRange(lr + 1, RETURNED).setValue('-yes-');
              //////////////////////////////////////////////////////////////////////////////////////////////

          // Insert blank row at end of last row BEFORE deletion loop, otherwise deletion will crash! 
          ms.insertRowsAfter(lastRow, 1);

          // delete row
          ms.deleteRow(j);  // delete row j
      
      }; // end if      
      
    }; // end for loop

    return_array.push([error, serialno]);

    return return_array;

}; // end function


function AddStaffReturnEntry(serialno) {

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var ms = ss.getSheetByName(STAFFSHEET);

    /* PULL DATE FROM const STATED ABOVE */

    // GET DATA
    var error = 'ERROR';
    var return_array = [];
    
    // DEFINE LAST ROW
    var lastRow = ms.getLastRow();
    
    ///////////////////////////////////// CHECK FOR RECORD & UPDATE  /////////////////////////////////////////
    
    for(var j=2; j <= lastRow; j++) {
      
      if( serialno == ms.getRange(j, SERIALNO).getValue() 
            && ms.getRange(j, RETURNED).getValue() == 'no') {
      
          error = 'SUCCESS';
          
          ms.getRange(j, RETURNED).setValue('yes');
          
          var fname = ms.getRange(j, STUFIRSTNAME).getValue();
          var lname = ms.getRange(j, STULASTNAME).getValue();
          var stuid = ms.getRange(j, STUID).getValue();

          // Send Email out
          if(EMAILON) {
              // sendEmail(recipient, subject, body, options)   
              MailApp.sendEmail(EMAILRECEPTS, `BHS Staff Loaner Returned - ${currentDay}`, 'Hello!', 
                                { cc: EMAILADMIN,
                                  name: 'Michael Stapleton',
                                  noReply: true,
                                  htmlBody: `[This is an auto-generated message]<br>
                                            ====================================<br>
                                            <font size="+1">Chromebook Loaner was returned:<br>
                                            Staff Name: <strong>${fname} ${lname} - ${stuid}</strong><br>
                                            CB Serial No.: <strong>${serialno}</strong><br>
                                            Date Returned: <strong>${currentDay}</strong></font>` 
                                });
              var emailQuotaRemaining = MailApp.getRemainingDailyQuota();
              Logger.log("Remaining email quota: " + emailQuotaRemaining);
          }; // end email on

              /////////////////////////////// add to History logs sheet ///////////////////////////////////
                var hs = ss.getSheetByName(LOGSSHEET);
                var lr = hs.getLastRow(); 

                var unique_id = (Math.floor(Math.random() * 2) + Date.now()).toString(36).slice(-5);
                return_date = GetDate(date);

                hs.getRange(lr + 1, UNIQUEID).setValue(unique_id);
                hs.getRange(lr + 1, DATE).setValue(return_date).setNumberFormat("MM/dd/yy hh:mm am/pm");
                hs.getRange(lr + 1, STUID).setValue(stuid);
                hs.getRange(lr + 1, STULASTNAME).setValue(lname);
                hs.getRange(lr + 1, STUFIRSTNAME).setValue(fname);
                /*hs.getRange(lr + 1, LOANERNO).setValue(loaner_no);*/
                hs.getRange(lr + 1, SERIALNO).setValue(serialno);
                /*hs.getRange(lr + 1, REASON).setValue(reason);*/
                hs.getRange(lr + 1, RETURNED).setValue('-yes-');
              //////////////////////////////////////////////////////////////////////////////////////////////

          // Insert blank row at end of last row BEFORE deletion loop, otherwise deletion will crash! 
          ms.insertRowsAfter(lastRow, 1);

          // delete row
          ms.deleteRow(j);  // delete row j
      
      }; // end if      
      
    }; // end for loop

    return_array.push([error, serialno]);

    return return_array;

}; // end function


/* =========================================================================================================================================== */


function SubmitExcludeEntry(student_id, last_name, first_name, reason) {
  
  //DEFINE ALL ACTIVE SHEETS
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  //DEFINE MAIN SHEET
  var es = ss.getSheetByName(EXCLUDESHEET); // exclude list

  //DEFINE LAST ROW
  var lastRow = es.getLastRow(); // exclude list
  
  //Define Return Variables
  var return_date = '';
  var error = 'SUCCESS';
  var return_array = [];

  //////////////////////////////////////////  CHECK FOR ANY DUPLICATES! /////////////////////////////////////////////////
  
  for(var j=2; j <= lastRow; j++) {
    
    if( first_name == es.getRange(j, STUFIRSTNAME).getValue() 
          && last_name == es.getRange(j, STULASTNAME).getValue()
              && student_id == es.getRange(j, STUID).getValue() ) {
      
          error = 'Error: Duplicate Record Found!';
          return_array.push([error, return_date, first_name, last_name, student_id]);
          
          return return_array;
    
    }; // end if 
    
  }; // end for loop

  ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
  
  if(error == 'SUCCESS') {

        var unique_id = (Math.floor(Math.random() * 2) + Date.now()).toString(36).slice(-5);
        return_date = GetDate(date);
        var datenotime = GetDateNoTime(date);
        
        // ADD RECORD
        es.getRange(lastRow + 1, UNIQUEID).setValue(unique_id);
        es.getRange(lastRow + 1, DATE).setValue(return_date).setNumberFormat("MM/dd/yy hh:mm am/pm");
        es.getRange(lastRow + 1, STUID).setValue(student_id);
        es.getRange(lastRow + 1, STULASTNAME).setValue(last_name);
        es.getRange(lastRow + 1, STUFIRSTNAME).setValue(first_name);
        es.getRange(lastRow + 1, REASON - 2).setValue(reason); // adjust the REASON value down to 2 [LOANERNO]

        // Send Email out
        if(EMAILON) {
            // sendEmail(recipient, subject, body, options)   
            MailApp.sendEmail(EMAILRECEPTS, `NEW BHS Student Exclusion Added - ${datenotime}`, 'Hello!', 
                              { cc: EMAILADMIN,
                                name: 'Michael Stapleton',
                                noReply: true,
                                htmlBody: `[This is an auto-generated message]<br>
                                          ====================================<br>
                                          <font size="+1">New student added to exclusion list:<br>
                                          Student Name: <strong>${first_name} ${last_name} - ${student_id}</strong><br>
                                          Date/Time Added: <strong>${return_date}</strong></font>` 
                              });
            var emailQuotaRemaining = MailApp.getRemainingDailyQuota();
            Logger.log("Remaining email quota: " + emailQuotaRemaining);
        }; // end email on

  }; // end

  return_array.push([error, return_date, first_name, last_name, student_id, reason]);
  
  return return_array;

}; // end function


function DeleteExcludeEntry(e) {

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var ms = ss.getSheetByName(EXCLUDESHEET);

    /* PULL DATE FRON const STATED ABOVE */

    // GET DATA
    var name = e.toString().split("|");
    var error = 'ERROR';
    var return_array = [];
    
    // DEFINE LAST ROW
    var lastRow = ms.getLastRow();

    // DEFINE NAME AFTER SPLIT IN EXACT ORDER
    for(var x=0; x < name.length; x++) {
        var stuid = name[x-2];
        var lname = name[x-1];
        var fname = name[x];
    };
    
    // CHECK FOR RECORD & UPDATE
    for(var j=2; j <= lastRow; j++) {
      
      if( fname == ms.getRange(j, STUFIRSTNAME).getValue() 
              && lname == ms.getRange(j, STULASTNAME).getValue() 
                    && stuid == ms.getRange(j, STUID).getValue() ) {
      
          // Send Email out
          if(EMAILON) {
              // sendEmail(recipient, subject, body, options)   
              MailApp.sendEmail(EMAILRECEPTS, `BHS Student Exclusion Removed - ${currentDay}`, 'Hello!', 
                                { cc: EMAILADMIN,
                                  name: 'Michael Stapleton',
                                  noReply: true,
                                  htmlBody: `[This is an auto-generated message]<br>
                                            ====================================<br>
                                            <font size="+1">Student was removed from exclusion list:<br>
                                            Student Name: <strong>${fname} ${lname} - ${stuid}</strong><br>
                                            Date Removed: <strong>${currentDay}</strong></font>` 
                                });
              var emailQuotaRemaining = MailApp.getRemainingDailyQuota();
              Logger.log("Remaining email quota: " + emailQuotaRemaining);
          }; // end email on

          // Insert blank row at end of last row BEFORE deletion loop, otherwise deletion will crash! 
          ms.insertRowsAfter(lastRow, 1);

          // delete row
          ms.deleteRow(j);  // delete row j
          
          error = 'SUCCESS';
    
      }; // end if      
      
    }; // end for loop

    return_array.push([error, fname, lname, stuid]);

    return return_array;

}; // end function


/* =========================================================================================================================================== */


function ListExclusionHeaders() {

    // DEFINE ALL ACTIVE SHEETS
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // DEFINE MAIN SHEET          
    var ms = ss.getSheetByName(EXCLUDESHEET);

    // GET DATA
    var data = ms.getRange(1, 3, 1, ms.getMaxColumns()-2).getValues();  // to hide ID & DATE fields, add [3] & [-2] to cols
        data = data.toString().toUpperCase().split(",");

    // EXPORT DATA
    var titleHTML = '<thead><tr>';
    
    // styling for last column [reason] to stretch 40% than all other cells
    for(var k=0, y=1; k < data.length; k++, y++) {

        titleHTML += ( (y % EXCNUMOFCOLS) == 0 ) ? '<th class="text-center w-40">' + data[k] + '</th>'
                                                    + '<th class="text-center">DELETE STUDENT?</th></tr><tr>' 
                                                 : '<th class="text-center">' + data[k] + '</th>'; 

    }; // end loop

    return titleHTML + '</tr></thead>';

}; // end function


function ListExclusionList() {
    
    // DEFINE ALL ACTIVE SHEETS
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // DEFINE MAIN SHEET          
    var ms = ss.getSheetByName(EXCLUDESHEET);

    // CHECK IF DATASHEET IS BLANK 
    if(ms.getLastRow()-1 < 1) { return };

    // GET DATA
    var data = ms.getRange(2, 3, ms.getLastRow()-1, ms.getMaxColumns()-2).getValues();  // to hide ID & DATE fields, add [3] & [-2] to cols
    
    // CONVERT TIMES & SPLIT DATA   --- useGrouping removes commas from numeric values [such as ID]----
    data = data.toLocaleString("en-US", { month:"numeric", day:"numeric", year:"numeric", useGrouping:false }).split(",");

    // EXPORT DATA
    var listHTML = '<tbody><tr>';
    
    for(var j=0, x=1; j < data.length; j++, x++) {
      
        listHTML += ( (x % EXCNUMOFCOLS) == 0 ) ? '<td>' + data[j] + '</td>'
                                                      + '<td><button type="button" class="button info" '
                                                      + 'value="' + data[j-3] + '|' + data[j-2] + '|' + data[j-1] + '" ' 
                                                      + 'onclick="return confirm(\'Are you sure you want to delete client?\')'
                                                      + '?Delete(this.value):alert(\'Removal cancelled\');">REMOVE</button></td></tr><tr>'
                                                : '<td>' + data[j] + '</td>';
      
    };  // end loop

    return listHTML + '</tr></tbody>';

}; // end function


/* =========================================================================================================================================== */


function ListHistoryHeaders() {

    // DEFINE ALL ACTIVE SHEETS
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // DEFINE MAIN SHEET          
    var ms = ss.getSheetByName(LOGSSHEET);

    // GET DATA
    var data = ms.getRange(1, 2, 1, ms.getMaxColumns()-1).getValues();  // to hide ID & DATE fields, add [3] & [-2]
        data = data.toString().toUpperCase().split(",");

    // EXPORT DATA
    var titleHTML = '<thead><tr>';
    
    for(var k=0, y=1; k < data.length; k++, y++) {
      
      titleHTML += '<th class="text-center">' + data[k] + '</th>'; 

        if( (y % HISTNUMOFCOLS) == 0 ) 
            
            titleHTML += '</tr><tr>'; // end if remainder

    }; // end loop

    return titleHTML + '</tr></thead>';

}; // end function


function ListHistoryLogs() {
    
    // DEFINE ALL ACTIVE SHEETS
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // DEFINE MAIN SHEET          
    var ms = ss.getSheetByName(LOGSSHEET);

    // CHECK IF DATASHEET IS BLANK 
    if(ms.getLastRow()-1 < 1) { return };

    // GET DATA
    var data = ms.getRange(2, 2, ms.getLastRow()-1, ms.getMaxColumns()-1).getValues();  // to hide ID & DATE fields, add [3] & [-2]
    
    // CONVERT TIMES & SPLIT DATA   --- useGrouping removes commas from numeric values [such as ID]----
    /*data = data.toLocaleString("en-US", { month:"numeric", day:"numeric", year:"numeric", 
                                            hour:"numeric", minute:"numeric", useGrouping:false }).split(","); */
    data = data.toString().split(",");

    // EXPORT DATA
    var listHTML = '<tbody><tr>';
    
    for(var j=0, x=1; j < data.length; j++, x++) {
      
        listHTML += '<td>' + data[j].toString().replace("GMT-0600 (Central Standard Time)","(CST)")
                                                .replace("GMT-0500 (Central Daylight Time)","(CDT)")
                                                .replace("-yes-","RETURNED")
                                                .replace("-no-","CHECKED OUT") + '</td>';
        
        if( (x % HISTNUMOFCOLS) == 0 ) 
              
            listHTML += '</tr><tr>'; // end if remainder

    };  // end loop

    return listHTML + '</tr></tbody>';

}; // end function


/* =========================================================================================================================================== */


function AddExpressEntry(serial_no, loaner_no, student_id, reason) {

    //DEFINE ALL ACTIVE SHEETS
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    //DEFINE MAIN SHEET          
    var ms = ss.getSheetByName(SHEETNAME);
    var es = ss.getSheetByName(EXCLUDESHEET); // exclude list

    //DEFINE LAST ROW
    var lastRow = ms.getLastRow();
    var elastRow = es.getLastRow(); // exclude list
    
    //Define Return Variables
    var return_date = '';
    var error = 'SUCCESS';
    var return_array = [];

    ////////////////////////////////////////////  CHECK FOR EXCLUSIONS! ///////////////////////////////////////////////////

    for(var k=2; k <= elastRow; k++) {

      if( student_id == es.getRange(k, STUID).getValue() ) {

            error = 'Student NOT allowed to be assigned a loaner!';
            return_array.push([error, return_date, student_id]);
            
            return return_array;

      }; // end if

    }; // end for loop
    
    //////////////////////////////////////////  CHECK FOR ANY DUPLICATES! /////////////////////////////////////////////////
    
    for(var j=2; j <= lastRow; j++) {
      
      if( student_id == ms.getRange(j, STUID).getValue() ) {
        
            error = 'Error: Duplicate Record Found:<br>Student has already checked out a loaner!';
            return_array.push([error, return_date, student_id]);
            
            return return_array;
      
      }; // end if 

      if( serial_no == ms.getRange(j, SERIALNO).getValue() 
            && ms.getRange(j, RETURNED).getValue() == 'no' ) {

            error = 'Error: Loaner already checked out.<br>Please select another loaner.';
            return_array.push([error, return_date, student_id]);

            return return_array;

      }; // end if
      
    }; // end for loop

    ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    
    if(error == 'SUCCESS') {

          var unique_id = (Math.floor(Math.random() * 2) + Date.now()).toString(36).slice(-5);
          return_date = GetDate(date);
          var datenotime = GetDateNoTime(date);
          
          // ADD RECORD
          ms.getRange(lastRow + 1, UNIQUEID).setValue(unique_id);
          ms.getRange(lastRow + 1, DATE).setValue(return_date).setNumberFormat("MM/dd/yy hh:mm am/pm");
          ms.getRange(lastRow + 1, STUID).setValue(student_id);
          ms.getRange(lastRow + 1, STULASTNAME).setValue("--Express--");
          ms.getRange(lastRow + 1, STUFIRSTNAME).setValue("--Express--");
          ms.getRange(lastRow + 1, LOANERNO).setValue(loaner_no);
          ms.getRange(lastRow + 1, SERIALNO).setValue(serial_no);
          ms.getRange(lastRow + 1, REASON).setValue(reason);
          ms.getRange(lastRow + 1, RETURNED).setValue('no');

          // Send Email out
          if(EMAILON) {
              // sendEmail(recipient, subject, body, options)   
              MailApp.sendEmail(EMAILRECEPTS, `NEW BHS Student Loaner Issued by Express Ckout - ${datenotime}`, 'Hello!', 
                                { cc: EMAILADMIN,
                                  name: 'Michael Stapleton',
                                  noReply: true,
                                  htmlBody: `[This is an auto-generated message]<br>
                                            ====================================<br>
                                            <font size="+1">Loaner was issued out by Express Ckout to:<br>
                                            Student ID: <strong>${student_id}</strong><br>
                                            CB Serial No.: <strong>${serial_no}</strong><br>
                                            Date/Time Issued: <strong>${return_date}</strong></font>` 
                                });
              var emailQuotaRemaining = MailApp.getRemainingDailyQuota();
              Logger.log("Remaining email quota: " + emailQuotaRemaining);
          }; // end email on

            //////////////////////////// add to History logs sheet /////////////////////////////////////////
              var hs = ss.getSheetByName(LOGSSHEET);
              var lr = hs.getLastRow(); 

              hs.getRange(lr + 1, UNIQUEID).setValue(unique_id);
              hs.getRange(lr + 1, DATE).setValue(return_date).setNumberFormat("MM/dd/yy hh:mm am/pm");
              hs.getRange(lr + 1, STUID).setValue(student_id);
              hs.getRange(lr + 1, STULASTNAME).setValue("--Express--");
              hs.getRange(lr + 1, STUFIRSTNAME).setValue("--Express--");
              hs.getRange(lr + 1, LOANERNO).setValue(loaner_no);
              hs.getRange(lr + 1, SERIALNO).setValue(serial_no);
              hs.getRange(lr + 1, REASON).setValue(reason);
              hs.getRange(lr + 1, RETURNED).setValue('-no-');
            /////////////////////////////////////////////////////////////////////////////////////////////////

    }; // end

    return_array.push([error, return_date, student_id, serial_no, loaner_no]);
    
    return return_array;
  
}; // end function


/* ============================================================================================================================================== */


function SubmitRepairsEntry(student_id, last_name, first_name, ticket_no, paid) {
  
      //DEFINE ALL ACTIVE SHEETS
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      
      //DEFINE MAIN SHEET
      var es = ss.getSheetByName(REPAIRSHEET); // repair loaners list

      //DEFINE LAST ROW
      var lastRow = es.getLastRow(); // exclude list
      
      //Define Return Variables
      var return_date = '';
      var error = 'success';
      var return_array = [];

      //////////////////////////////////////////  CHECK FOR ANY DUPLICATES! /////////////////////////////////////////////////
      
      for(var j=2; j <= lastRow; j++) {
        
            if( first_name == es.getRange(j, STUFIRSTNAME).getValue() 
                  && last_name == es.getRange(j, STULASTNAME).getValue()
                      && student_id == es.getRange(j, STUID).getValue() ) {
              
                  error = 'Error: Duplicate Record Found!';
                  return_array.push([error, return_date, first_name, last_name, student_id]);
                  
                  return return_array;
            
            }; // end if 

            if( student_id == es.getRange(j, STUID).getValue() ) {
              
                  error = 'Error: Duplicate Record Found!';
                  return_array.push([error, return_date, student_id, '', '']);
                  
                  return return_array;
            
            }; // end if 
        
      }; // end for loop

      ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
      
      if(error == 'success') {

            var unique_id = (Math.floor(Math.random() * 2) + Date.now()).toString(36).slice(-5);
            return_date = GetDate(date);
            var datenotime = GetDateNoTime(date);
            var expire_date = new Date(date.getTime() + 14 * 24 * 3600 * 1000);  // 2 wks out * 24 * 60 * 60 * 1000;
            
            // ADD RECORD
            es.getRange(lastRow + 1, UNIQUEID).setValue(unique_id);
            es.getRange(lastRow + 1, DATE).setValue(return_date).setNumberFormat("MM/dd/yy hh:mm am/pm");
            es.getRange(lastRow + 1, STUID).setValue(student_id);
            es.getRange(lastRow + 1, STULASTNAME).setValue(last_name);
            es.getRange(lastRow + 1, STUFIRSTNAME).setValue(first_name);
            es.getRange(lastRow + 1, TECHTICKETNO).setValue(ticket_no); 
            es.getRange(lastRow + 1, FEESPAID).setValue(paid); 
            es.getRange(lastRow + 1, TIMEEXPIRED).setValue('no');
            es.getRange(lastRow + 1, EXPIRATIONDATE).setValue(expire_date).setNumberFormat("MM/dd/yy hh:mm am/pm"); 

            // Send Email out
            if(EMAILON) {
                // sendEmail(recipient, subject, body, options)   
                MailApp.sendEmail(EMAILRECEPTS, `NEW BHS Student Repairs Entry Added - ${datenotime}`, 'Hello!', 
                                  { cc: EMAILADMIN,
                                    name: 'Michael Stapleton',
                                    noReply: true,
                                    htmlBody: `[This is an auto-generated message]<br>
                                              ====================================<br>
                                              <font size="+1">New student added to the repairs list:<br>
                                              Student Name: <strong>${first_name} ${last_name} - ${student_id}</strong><br>
                                              Date/Time Added: <strong>${return_date}</strong><br>
                                              Tech Ticket No.: <strong>${ticket_no}</strong><br>
                                              Paid: <strong>${paid}</strong><br>
                                              Time Expired: <strong>NO</strong></font>` 
                                  });
                var emailQuotaRemaining = MailApp.getRemainingDailyQuota();
                Logger.log("Remaining email quota: " + emailQuotaRemaining);
            }; // end email on

      }; // end

      return_array.push([error, return_date, first_name, last_name, student_id, ticket_no, paid]);
      
      return return_array;

}; // end function


function DeleteRepairsEntry(e) {

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var ms = ss.getSheetByName(REPAIRSHEET);
    var es = ss.getSheetByName(EXCLUDESHEET);  // exclude list

    /* PULL DATE FRON const STATED ABOVE */

    // GET DATA
    var name = e.toString().split("|");
    var error = 'Error: ';
    var return_array = [];
    
    // DEFINE LAST ROW
    var lastRow = ms.getLastRow();
    var elastRow = es.getLastRow();   // exclude sheet

    // DEFINE NAME AFTER SPLIT IN EXACT ORDER - "x" VARIABLE IS IN LAST PLACE
    for(var x=0; x < name.length; x++) {
        var stuid = name[x-2];
        var lname = name[x-1];
        var fname = name[x];
    };
    
    /* ===========================   CHECK FOR BLACKLIST ENTRY   ============================== */
        
        for(var k=2; k <= elastRow; k++) {

            if( stuid == es.getRange(k, STUID).getValue() 
                  && fname == es.getRange(k, STUFIRSTNAME).getValue()
                      && lname == es.getRange(k, STULASTNAME).getValue() ) {

                        // Insert blank row at end of last row BEFORE deletion loop, otherwise deletion will crash! 
                        if( elastRow - 1 < 1) es.insertRowsAfter(elastRow, 1);
                        
                        // delete row
                        es.deleteRow(k);  // delete row k
                        
                        Logger.log('delete blacklist: ' + k);
             }; 

        }; // end for loop

    /* ======================================================================================== */
    
    // CHECK FOR RECORD & UPDATE
    for(var j=2; j <= lastRow; j++) {
      
          if( fname == ms.getRange(j, STUFIRSTNAME).getValue() 
                  && lname == ms.getRange(j, STULASTNAME).getValue() 
                        && stuid == ms.getRange(j, STUID).getValue() ) {
          
                  // Send Email out
                  if(EMAILON) {
                      // sendEmail(recipient, subject, body, options)   
                      MailApp.sendEmail(EMAILRECEPTS, `BHS Student Repair List Removed - ${currentDay}`, 'Hello!', 
                                        { cc: EMAILADMIN,
                                          name: 'Michael Stapleton',
                                          noReply: true,
                                          htmlBody: `[This is an auto-generated message]<br>
                                                    ====================================<br>
                                                    <font size="+1">Student was removed from the repairs list:<br>
                                                    Student Name: <strong>${fname} ${lname} - ${stuid}</strong><br>
                                                    Date Removed: <strong>${currentDay}</strong></font>` 
                                        });
                      var emailQuotaRemaining = MailApp.getRemainingDailyQuota();
                      Logger.log("Remaining email quota: " + emailQuotaRemaining);
                  }; // end email on

                  // Insert blank row at end of last row BEFORE deletion loop, otherwise deletion will crash! 
                  if( lastRow - 1 < 1 ) ms.insertRowsAfter(lastRow, 1);
                  
                  // delete row
                  ms.deleteRow(j);  // delete row j

                  Logger.log('deleted paid: ' + j);
                  
                  error = 'success';
        
          }; // end if      
      
    }; // end for loop

    return_array.push([error, fname, lname, stuid]);

    return return_array;

}; // end function


function AddTicketNoToEntry(student_id, ticket_no) {

        //DEFINE ALL ACTIVE SHEETS
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        
        //DEFINE MAIN SHEET
        var rs = ss.getSheetByName(REPAIRSHEET); // repair loaners list

        //DEFINE LAST ROW
        var lastRow = rs.getLastRow(); // exclude list
        
        /* PULL DATE FROM const STATED ABOVE */

        //Define Return Variables
        var error = 'Error: ';
        var return_array = [];

        //////////////////////////////////////////  CHECK FOR RECORD & UPDATE! /////////////////////////////////////////////////
        
       for(var j=2; j <= lastRow; j++) {
      
            if(student_id == rs.getRange(j, STUID).getValue()) {
               
               rs.getRange(j, TECHTICKETNO).setValue(ticket_no);
               
               error = 'success'; 
            
            };
          
        }; // end for loop

        ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

        return_array.push([error, student_id, ticket_no]);
        
        return return_array;

}; // end function


/* =========================================================================================================================================== */


function ListRepairsHeaders() {

    // DEFINE ALL ACTIVE SHEETS
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // DEFINE MAIN SHEET          
    var ms = ss.getSheetByName(REPAIRSHEET);

    // GET DATA
    var data = ms.getRange(1, 3, 1, ms.getMaxColumns()-2).getValues();  // to hide ID & DATE fields, add [3] & [-2] to cols
        data = data.toString().toUpperCase().split(",");

    // EXPORT DATA
    var titleHTML = '<thead><tr>';
    
    // styling for last column [reason] to stretch 40% than all other cells
    for(var k=0, y=1; k < data.length; k++, y++) {

        titleHTML += ( (y % REPAIRNUMOFCOLS) == 0 ) ? '<th class="text-center w-10">' + data[k] + '</th>'
                                                    + '<th class="text-center">PAID IN FULL?</th>'
                                                    + '<th class="text-center">TICKET NO.</th></tr><tr>' 
                                                 : '<th class="text-center">' + data[k] + '</th>'; 

    }; // end loop

    return titleHTML + '</tr></thead>';

}; // end function


function ListRepairsList() {
    
    // DEFINE ALL ACTIVE SHEETS
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // DEFINE MAIN SHEET          
    var ms = ss.getSheetByName(REPAIRSHEET);

    // CHECK IF DATASHEET IS BLANK 
    if(ms.getLastRow()-1 < 1) { 
                                var colspan = REPAIRNUMOFCOLS+2; // advance colspan to 2 
                                return '<tbody><tr><td colspan="'+ colspan +'">'
                                        + '<span class="mif-checkmark fg-green"></span>'
                                        + '&nbsp;No Awaiting Repairs!&nbsp;'
                                        + '<span class="mif-checkmark fg-green"></span>'
                                        + '</td></tr></tbody>' };

    // GET DATA
    var data = ms.getRange(2, 3, ms.getLastRow()-1, ms.getMaxColumns()-2).getValues();  // to hide ID & DATE fields, add [3] & [-2] to cols
    
    // CONVERT TIMES & SPLIT DATA   --- useGrouping removes commas from numeric values [such as ID]----
    data = data.toLocaleString("en-US", { month:"numeric", day:"numeric", year:"numeric", useGrouping:false }).split(",");

    // EXPORT DATA
    var listHTML = '<tbody><tr>';
    
    for(var j=0, x=1; j < data.length; j++, x++) {
      
        listHTML += ( (x % REPAIRNUMOFCOLS) == 0 ) ? '<td>' + data[j] + '</td>'
                                                      + '<td><button type="button" class="button info" '
                                                      + 'value="'+ data[j-6] +'|'+ data[j-5] +'|'+ data[j-4] +'" ' 
                                                      + 'onclick="return confirm(\'Confirm student is paid in full?\')'
                                                      + '?Delete(this.value):alert(\'Update cancelled\');">PAID</button></td>'
                                                      + '<td><button type="button" class="button info" '
                                                      + 'onclick="window.top.location.replace(\''+ REPAIRSURL + ''
                                                      + '&id='+ data[j-6] +'\');">EDIT</button></td></tr><tr>'
                                                : '<td>' + data[j] + '</td>';
      
    };  // end loop

    return listHTML + '</tr></tbody>';

}; // end function


/* ====================================================== END CUSTOM FUNCTIONS =============================================================== */



/* =========================================================================================================================================== */


function SortFirstName() {
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var ms = ss.getSheetByName(SHEETNAME).sort(STUFIRSTNAME);
    SpreadsheetApp.flush();
    
    return ms;

}; // end function


function SortStaffFirstName() {
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var ms = ss.getSheetByName(STAFFSHEET).sort(STUFIRSTNAME);
    SpreadsheetApp.flush();
    
    return ms;

}; // end function


function SortFirstNameHistory() {
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var ms = ss.getSheetByName(LOGSSHEET).sort(STUFIRSTNAME);
    SpreadsheetApp.flush();
    
    return ms;

}; // end function


function SortLastName() {
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var ms = ss.getSheetByName(SHEETNAME).sort(STULASTNAME);
    SpreadsheetApp.flush();
    
    return ms;

}; // end function


function SortStaffLastName() {
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var ms = ss.getSheetByName(STAFFSHEET).sort(STULASTNAME);
    SpreadsheetApp.flush();
    
    return ms;

}; // end function


function SortLastNameHistory() {
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var ms = ss.getSheetByName(LOGSSHEET).sort(STULASTNAME);
    SpreadsheetApp.flush();
    
    return ms;

}; // end function


function SortToDate() {

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var ms = ss.getSheetByName(SHEETNAME).sort(DATE);
    SpreadsheetApp.flush();
    
    return ms;

}; // end function


function SortStaffToDate() {

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var ms = ss.getSheetByName(STAFFSHEET).sort(DATE);
    SpreadsheetApp.flush();
    
    return ms;

}; // end function


function SortToDateHistory() {

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var ms = ss.getSheetByName(LOGSSHEET).sort(DATE);
    SpreadsheetApp.flush();
    
    return ms;

}; // end function


function SortToLoaner() {

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var ms = ss.getSheetByName(SHEETNAME).sort(LOANERNO);
    SpreadsheetApp.flush();
    
    return ms;

}; // end function


function SortStaffToLoaner() {

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var ms = ss.getSheetByName(STAFFSHEET).sort(LOANERNO);
    SpreadsheetApp.flush();
    
    return ms;

}; // end function


function SortToLoanerHistory() {

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var ms = ss.getSheetByName(LOGSSHEET).sort(LOANERNO);
    SpreadsheetApp.flush();
    
    return ms;

}; // end function


function SortToID() {

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var ms = ss.getSheetByName(SHEETNAME).sort(STUID);
    SpreadsheetApp.flush();
    
    return ms;

}; // end function


function SortToIDHistory() {

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var ms = ss.getSheetByName(LOGSSHEET).sort(STUID);
    SpreadsheetApp.flush();
    
    return ms;

}; // end function


/* ============================================================================================================================================== */


function ReloadPage() {

    $('#all-spinner').toggleClass("active", true);

    setTimeout(function(){
      google.script.run.withSuccessHandler(function(url){
                                            window.open(url,'_top');
                                          }).GetScriptURL();
    }, 1000);

}; // end function


/* =============================================================================================================================================== */


function AddZero(i) {
  return (i < 10) ? "0" + i : i;
};


function GetDate(date_in) {
  var currentDate = date_in;
  var currentMonth = currentDate.getMonth() + 1;
  var currentYear = currentDate.getFullYear();
  var currentHours = (AddZero(currentDate.getHours()) > 12) ? AddZero(currentDate.getHours()) - 12 
                                                            : AddZero(currentDate.getHours());
  var currentMinutes = AddZero(currentDate.getMinutes());
  var currentSeconds = AddZero(currentDate.getSeconds());
  var suffix = (AddZero(currentDate.getHours()) >= 12)? 'PM' : 'AM';
  var date = currentMonth.toString() + '/' + currentDate.getDate().toString() + '/' + 
             currentYear.toString() + ' ' + currentHours.toString() + ':' +
             currentMinutes.toString() + ':' + currentSeconds.toString() + ' ' + suffix;
  
  return date;
}; 


function GetDateNoTime(date_in) {
  var currentDate = date_in;
  var currentMonth = currentDate.getMonth() + 1;
  var currentYear = currentDate.getFullYear();
  var date = currentMonth.toString() + '/' + 
              currentDate.getDate().toString() + '/' + 
              currentYear.toString();
  
  return date;
};

/* ================================================================================================================================== */


/* =================================================    TRIGGER CODING    =========================================================== */

function ExportAndSend() {
  
      const date = new Date();
      const currentDay = date.toLocaleString('en-US', { month:'numeric', day:'numeric', year:'numeric' });
      
      // prepare a PDF of the sheet & add to the Google Drive as a backup file
      let checkoutSheet = DriveApp.getFileById(SHEETID);
      let blob = checkoutSheet.getAs('application/pdf');
      let pdf = DriveApp.getFolderById(FOLDERID)
                        .createFile(blob)
                        .setName(`${currentDay} - BHS Student Loaners Past Due`);

      // send email with PDF attached
      // sendEmail(recipient, subject, body, options)   
      MailApp.sendEmail(EMAILADMIN, `BHS Student Loaners Past Due - ${currentDay}`, 'Hello!', 
                        { attachments: [pdf], 
                          bcc: 'mstapleton@bentonvillek12.org',
                          name: 'No Reply',
                          noReply: true,
                          htmlBody: `PDF attached shows remaining loaners past due` 
                        });
      var emailQuotaRemaining = MailApp.getRemainingDailyQuota();
      Logger.log("Remaining email quota: " + emailQuotaRemaining);

}; // end function


function ClearOutFiles() {
  
      // variables
      const files = DriveApp.getFolderById(FOLDERID).getFiles();
      
      // search for files
      while (files.hasNext()) {
        var file = files.next();
        Logger.log(file.getName());
        file.setTrashed(true); // delete file
      }; // end while loop

      // send email confirmation
      // sendEmail(recipient, subject, body, options)
      MailApp.sendEmail('mstapleton@bentonvillek12.org',
                        'Export PDF BHS Student Loaners Past Due Deletion',
                        'BHS Student Loaners Past Due PDFs - Weekly Deletion completed!',
                        { noReply: true, name: 'No Reply' });

}; // end function


function ClearOldEntries() {
  
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sh = ss.getSheetByName(SHEETNAME);

      SpreadsheetApp.setActiveSheet(sh);

      var MyRange = sh.getRange(1, DATE, sh.getLastRow(), 1); // selects B column [2] where the date value is stored based on sheet.
      var dataInColumnB = MyRange.getValues();
      //Logger.log('dataInColumnB: ' + dataInColumnB);

      ////////// ADDED code for special case ////////////
      var MyReturnRange = sh.getRange(1, RETURNED, sh.getLastRow(), 1); // selects RETURNED col 
      var dataInReturnCol = MyReturnRange.getValues();
      ///////////////////////////////////////////////////

      var todaysDate = new Date();
      var todayAsMilliseconds = todaysDate.getTime();
      Logger.log('todayAsMilliseconds: ' + todayAsMilliseconds);

      var WeekAgo = todayAsMilliseconds - (NUMOFDAYSOLD * (24 * 3600 * 1000));  // day * 24 * 60 * 60 * 1000;
      Logger.log('WeekAgo: ' + WeekAgo);

      var valueInCell,
          dateInCell,
          lengthOfData = dataInColumnB.length;
      
      // Insert blank row at end of last row BEFORE deletion loop, otherwise deletion will crash!
      if(lengthOfData > 0)
          sh.insertRowsAfter(sh.getLastRow(), 1);

      // loop thru each entry to detect if old
      for (var i=lengthOfData; i > 0; i -= 1) { //decrement count
                                                  ////////// ADDED for special case /////////
          if (dataInColumnB[i-1] === undefined || dataInReturnCol[i-1] === undefined) { continue };
          valueInCell = dataInColumnB[i-1][0];

          ////////// ADDED for special case /////////
          var returned = dataInReturnCol[i-1][0];
          ///////////////////////////////////////////

          // bypass empty cells and headers
          if (valueInCell === undefined || valueInCell === "" || valueInCell == "Timestamp" ) { continue };

          Logger.log('valueInCell: ' + valueInCell);
          Logger.log('typeof valueInCell: ' + typeof valueInCell);

          dateInCell = valueInCell.getTime();
          Logger.log('dateInCell: ' + dateInCell);

          // Delete rows          // ADDED for special case //
          if(dateInCell < WeekAgo && returned == 'yes'){
            Logger.log('Row - ' + i + ' is more than a day old');
            sh.deleteRow(i);  // DELETE row i
          };

      }; //end for loop
  
}; //end function


/* =============================================   COPY TO DUPLICATE SHEET   =========================================================== */

function CopyToDuplicateSheet() {

      /* TO COPY FROM THIS ACTIVE SHEET TO A COPY ON THE TSS SHARED DRIVE - RUN SCRIPT NIGHTLY AT 3AM CST  */

      // student ckouts
      var source = SpreadsheetApp.openById(SHEETID);
      var sourceSheet = source.getSheetByName(SHEETNAME);
      var sourceRange = sourceSheet.getDataRange();
      var sourceValues = sourceRange.getValues();
      var tempSheet = source.getSheetByName('temp');
      var tempRange = tempSheet.getRange('A1');
      var destination = SpreadsheetApp.openById(DUPLSHEETID);
      var destSheet = destination.getSheetByName(SHEETNAME);
      
      Logger.log('Starting : '+ SHEETNAME);
      destSheet.clear(); // clear any old saved data on duplicate sheet
      sourceRange.copyTo(tempRange);  // paste all formats?, broken references
      tempRange.offset(0, 0, sourceValues.length, sourceValues[0].length).setValues(sourceValues);  // paste all values (over broken refs)
      copydSheet = tempSheet.copyTo(destination);   // now copy temp sheet to another ss
      copydSheet.getDataRange().copyTo(destSheet.getDataRange());
      destination.deleteSheet(copydSheet); //delete copydSheet
      tempSheet.clear();
      Logger.log('Finished : '+ SHEETNAME);

      // staff ckouts
      var source = SpreadsheetApp.openById(SHEETID);
      var sourceSheet = source.getSheetByName(STAFFSHEET);
      var sourceRange = sourceSheet.getDataRange();
      var sourceValues = sourceRange.getValues();
      var tempSheet = source.getSheetByName('temp');
      var tempRange = tempSheet.getRange('A1');
      var destination = SpreadsheetApp.openById(DUPLSHEETID);
      var destSheet = destination.getSheetByName(STAFFSHEET);
      
      Logger.log('Starting : '+ STAFFSHEET);
      destSheet.clear(); // clear any old saved data on duplicate sheet
      sourceRange.copyTo(tempRange);  // paste all formats?, broken references
      tempRange.offset(0, 0, sourceValues.length, sourceValues[0].length).setValues(sourceValues);  // paste all values (over broken refs)
      copydSheet = tempSheet.copyTo(destination);   // now copy temp sheet to another ss
      copydSheet.getDataRange().copyTo(destSheet.getDataRange());
      destination.deleteSheet(copydSheet); //delete copydSheet
      tempSheet.clear();
      Logger.log('Finished : '+ STAFFSHEET);

      // no loaner list
      var source = SpreadsheetApp.openById(SHEETID);
      var sourceSheet = source.getSheetByName(EXCLUDESHEET);
      var sourceRange = sourceSheet.getDataRange();
      var sourceValues = sourceRange.getValues();
      var tempSheet = source.getSheetByName('temp');
      var tempRange = tempSheet.getRange('A1');
      var destination = SpreadsheetApp.openById(DUPLSHEETID);
      var destSheet = destination.getSheetByName(EXCLUDESHEET);
      
      Logger.log('Starting : '+ EXCLUDESHEET);
      destSheet.clear(); // clear any old saved data on duplicate sheet
      sourceRange.copyTo(tempRange);  // paste all formats?, broken references
      tempRange.offset(0, 0, sourceValues.length, sourceValues[0].length).setValues(sourceValues);  // paste all values (over broken refs)
      copydSheet = tempSheet.copyTo(destination);   // now copy temp sheet to another ss
      copydSheet.getDataRange().copyTo(destSheet.getDataRange());
      destination.deleteSheet(copydSheet); //delete copydSheet
      tempSheet.clear();
      Logger.log('Finished : '+ EXCLUDESHEET);

}; //end function


/* ====================================================   MOVE TO BLACKLIST   =============================================================== */


function MoveToBlacklist() {

      /* FOR REPAIRS MODULE - TRIGGER THIS FUNCTION DAILY AT 6AM CST */

      // DEFINE ALL ACTIVE SHEETS
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      
      // DEFINE MAIN SHEET          
      var rs = ss.getSheetByName(REPAIRSHEET);  // repair sheet
      var es = ss.getSheetByName(EXCLUDESHEET); // blacklist sheet

      var rLastRow = rs.getLastRow();  // repairs sheet
      var eLastRow = es.getLastRow(); // exclude list

      // CHECK DATA
      var MyRange = rs.getRange(1, DATE, rs.getLastRow(), 1); // selects B column [2] where the date value is stored based on sheet.
      var dataInColumnB = MyRange.getValues();
      //Logger.log('dataInColumnB: ' + dataInColumnB);

      var todaysDate = new Date();
      var todayAsMilliseconds = todaysDate.getTime();
      Logger.log('todayAsMilliseconds: ' + todayAsMilliseconds);

      var WeekAgo = todayAsMilliseconds - (REPAIRNUMOFDAYS * (24 * 3600 * 1000));  // day * 24 * 60 * 60 * 1000;
      Logger.log('WeekAgo: ' + WeekAgo);

      var valueInCell,
          dateInCell,
          lengthOfData = dataInColumnB.length;

      // loop thru each entry to detect if old
      for (var i=lengthOfData; i > 0; i -= 1) { //decrement count
                                                  
          if (dataInColumnB[i-1] === undefined) continue;
          valueInCell = dataInColumnB[i-1][0];

          // bypass empty cells and headers
          if (valueInCell === undefined || valueInCell === "" || valueInCell == "Timestamp" ) continue;

          Logger.log('valueInCell: ' + valueInCell);
          Logger.log('typeof valueInCell: ' + typeof valueInCell);

          dateInCell = valueInCell.getTime();
          Logger.log('dateInCell: ' + dateInCell);

          // Move to BLACKLIST   
          if(dateInCell < WeekAgo) {
            
                var unique_id = (Math.floor(Math.random() * 2) + Date.now()).toString(36).slice(-5);
                var return_date = GetDate(date);
                
                var student_id  = rs.getRange(i, STUID).getValue();
                var last_name   = rs.getRange(i, STULASTNAME).getValue();
                var first_name  = rs.getRange(i, STUFIRSTNAME).getValue();
                var reason      = "repairs list :: loaner term expired";

                // CHECK RECORD IF TIME HAS EXPIRED FIRST
                for( var k=2; k <= rLastRow; k++ ) {
                      
                      // UPDATE RECORD ON REPAIRS SHEET
                      rs.getRange(i, TIMEEXPIRED).setValue('yes');

                      if(lengthOfData > 0) es.insertRowsAfter(eLastRow,1);
                      
                      // ADD RECORD TO BLACKLIST
                      es.getRange(eLastRow + 1, UNIQUEID).setValue(unique_id);
                      es.getRange(eLastRow + 1, DATE).setValue(return_date).setNumberFormat("MM/dd/yy hh:mm am/pm");
                      es.getRange(eLastRow + 1, STUID).setValue(student_id);
                      es.getRange(eLastRow + 1, STULASTNAME).setValue(last_name);
                      es.getRange(eLastRow + 1, STUFIRSTNAME).setValue(first_name);
                      es.getRange(eLastRow + 1, REASON - 2).setValue(reason); // adjust the REASON value down to 2 [LOANERNO]

                      Logger.log('Row - ' + i + ' has been blacklisted.');
                
                }; // end for loop

          }; // end if

      }; // end for loop

      ////////////////////     CLEAN BLACKLIST FOR DUPLICATES     //////////////////////
      
          es.getDataRange().removeDuplicates([STUID]);

          CleanBlankRows();

      //////////////////////////////////////////////////////////////////////////////////

}; // end function


function CleanBlankRows() {

      // remove blank rows //
      var sheetnames = [EXCLUDESHEET];  // list in array google sheets to be cleaned
      var ss = SpreadsheetApp.getActive();
      var allsheets = ss.getSheets();
      
      for (var s in allsheets) {
          var sheet = allsheets[s];
          if (sheetnames.includes(sheet.getName())) {
              var maxRows = sheet.getMaxRows();
              var lastRow = sheet.getLastRow();
              if (maxRows - lastRow != 0) {
                  sheet.deleteRows(lastRow + 1, maxRows - lastRow);
              };
          };
      };

}; // end function


/* =========================================================================================================================================================== */ 
