var spreadsheet = SpreadsheetApp.getActive();
var masterSheet = spreadsheet.getSheetByName('Master');
var dataSheet = spreadsheet.getSheetByName('Data');
var masterCount = dataSheet.getRange('B2').getValue();
var dayCount = dataSheet.getRange('B4').getValue();
var daySheet = spreadsheet.getSheetByName('Today');
var date = dataSheet.getRange('B1').getDisplayValue();
var blueColumnAtt = 7;
var blueColumnSign = 8;
var commColumnAtt = 9;
var commColumnSign = 10;
var signoutNameBox = 'K3';
var signoutDateBox = 'K8';
var teacherNameBox = 'G1';
const months = ["January","February","March","April","May","June","July","August","September","October","November","December"];
var monthNumeric = dataSheet.getRange('C1').getValue();
var month = months[(monthNumeric-1)];
var day = dataSheet.getRange('D1').getDisplayValue();
var lastSave = dataSheet.getRange('B5').getDisplayValue();
var offDay = dataSheet.getRange('B6').getDisplayValue();
var offReason = dataSheet.getRange('B7').getDisplayValue();
var isSchoolDay = dataSheet.getRange('K1').getValue();
if (isSchoolDay) {
  var dayColumn = masterSheet.createTextFinder(date).findNext().getColumn();
}
var weekDayNum = (dataSheet.getRange('B10').getDisplayValue()-2);
const days = ["Monday","Tuesday","Wednesday","Thursday","Friday"];
var weekDay = days[weekDayNum];
var autoSheet = spreadsheet.getSheetByName('Auto-Signouts');

function generateToday() {

 if (isSchoolDay) {
    newDay();

      //Hide all the columns before today in the master sheet
    if (!(dayColumn == null)) {
      masterSheet.hideColumns(7,dayColumn-7);
    }
    cleanSlate();

    dayFormula();

    copyMastertoDay();

    //rectifySignouts();
 }
}

function copyMastertoDay() {
  //Copy a signouts for the new day from the master sheet into the day sheet

  var masterNames = masterSheet.getRange(4,1,masterCount,6).getValues();
  daySheet.getRange(3,1,masterCount,6).setValues(masterNames);
  dataSheet.getRange('I4').setValue(date);

}

function saveMasterYesterday() {

  //Save all of the current values in the day sheet to yesterday's column in the master sheet, only to be used if data didn't save one day!

  var dayAtt = daySheet.getRange(3,7,masterCount,4).getValues();
  masterSheet.getRange(4,(dayColumn-4),masterCount,4).setValues(dayAtt);
}

function saveMasterToday() {

  //Save all of the current values in the day sheet to yesterday's column in the master sheet, should only run if today's data has not already been saved!
  if(isSchoolDay) {
    if (!(date == lastSave)) {
      var dayAtt = daySheet.getRange(3,7,masterCount,4).getValues();
      masterSheet.getRange(4,(dayColumn),masterCount,4).setValues(dayAtt);
      dataSheet.getRange('B5').setValue(date);
    } else {
      throw new Error("The data was already saved for today.");
    }
    }
}

function rectifySignouts() {
  if(isSchoolDay) {
  for (var i = 4; i <= (masterCount+3); i++) { 
    var student = masterSheet.getRange(i,3).getValue();
    var blueSign = masterSheet.getRange(i,(dayColumn+1)).getValue();
    var commSign = masterSheet.getRange(i,(dayColumn+3)).getValue();
    if (!(blueSign == "")) {
      var teacherSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(blueSign);
      teacherSheet.getRange(signoutNameBox).setValue(student);
      teacherSheet.activate();
      signout("blue");
    }
    if (!(commSign == "")) {
      var teacherSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(commSign);
      teacherSheet.getRange(signoutNameBox).setValue(student);
      teacherSheet.activate();
      signout("comm");
    }
  }
  dataSheet.getRange('I5').setValue(date);
  }
}

function signout(color) {
 
  //variables that we need regardless of which RT we are signing out for, and resetting the signout box

  var teacherSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var student = teacherSheet.getRange(signoutNameBox).getValue();
  var teacher = teacherSheet.getRange(teacherNameBox).getValue();
  var signoutDate = teacherSheet.getRange(signoutDateBox).getDisplayValue();
  teacherSheet.getRange(signoutNameBox).clearContent();
  teacherSheet.getRange(signoutDateBox).setFormula("TODAY()");
  var studentIndex = daySheet.createTextFinder(student).findNext().getRow();
  var checkBlue = daySheet.getRange(studentIndex,blueColumnSign).getValue();
  var checkComm = daySheet.getRange(studentIndex,commColumnSign).getValue();
  var sameSheet = teacherSheet.createTextFinder(student).findNext();

 // variables which are dependent on whether we are signing out for blue/gold or common

  if (color == "blue") {
    var columnAtt = blueColumnAtt;
    var columnSign = blueColumnSign;
    var oColumnSign = commColumnSign;
    var checkThis = checkBlue;
    var checkThat = checkComm;
    var thisRT = "Blue/Gold";
  } else if (color == "comm") {
    var columnAtt = commColumnAtt;
    var columnSign = commColumnSign;
    var oColumnSign = blueColumnSign;
    var checkThis = checkComm;
    var checkThat = checkBlue;
    var thisRT = "Common";
  }

 // If the student appears multiple times in the sheet, tIndex will find that for us

  if (!(sameSheet == null)){
    var tIndex = teacherSheet.createTextFinder(student).findNext().getRow();
  }

 //This case is just for signing out for the current day

  if (signoutDate == date) {  

    //Signout for a student who is either not signed out to anyone, or is signed out to you

    if (checkThis == "" || checkThis == teacher) {

    //Signing out a student who is not normally in your rider time

      if (!(tIndex <= 39)) { 

      //Signout for a student who is not currently anywhere in your RT sheet 

        if (!(checkThis == teacher)&&(sameSheet == null)) {
          daySheet.getRange(studentIndex,columnSign).setFormula('"'+teacher+'"');
          teacherSheet.insertRowsBefore(teacherSheet.getRange('41:41').getRow(),1);
          //daySheet.getRange(41,columnSign).setFormula("VLOOKUP(C41,'Today'!C3:N700,"+(columnSign-2)+",FALSE)");
          teacherSheet.getRange('C41').setValue(student);
          teacherSheet.getRange(41,columnSign).setFormula("VLOOKUP(C41,'Today'!C3:N700,"+(columnSign-2)+",FALSE)");
          teacherSheet.getRange(41,columnAtt).setValue("Present");
          daySheet.getRange(studentIndex,columnAtt).setFormula("=VLOOKUP(C"+studentIndex+", '"+teacher+"'!C41:N700,"+(columnAtt-2)+",FALSE)");
          spreadsheet.toast(student+" was signed out for "+thisRT);
        
      //Signout for a student who you have signed out for the other RT, but not this RT

        } else if ((checkThat == teacher) && !(checkThis == teacher)) {
          daySheet.getRange(studentIndex,columnSign).setFormula('"'+teacher+'"');
          teacherSheet.getRange(tIndex,columnSign).setFormula("VLOOKUP(C"+tIndex+", 'Today'!C3:N700,"+(columnSign-2)+",FALSE)");
          teacherSheet.getRange(tIndex,columnAtt).setValue("Present");
          daySheet.getRange(studentIndex,columnAtt).setFormula("VLOOKUP(C"+studentIndex+", '"+teacher+"'!C41:N72,"+(columnAtt-2)+",FALSE)");
          spreadsheet.toast(student+" was signed out for "+thisRT);
        
      //unsignout a student who is not in your RT

        } else {
          var ogTeach = daySheet.getRange(studentIndex,4).getValue();
          var otherRT = teacherSheet.getRange(tIndex,oColumnSign).getValue();
          teacherSheet.getRange(tIndex,columnAtt,1,2).clearContent();
          if (otherRT == "") {
            teacherSheet.deleteRow(tIndex);
          }
          daySheet.getRange(studentIndex,columnAtt).setFormula("=VLOOKUP(C"+studentIndex+", '"+ogTeach+"'!C4:N200,"+(columnAtt-2)+",FALSE)");
          daySheet.getRange(studentIndex,columnSign).clearContent();
          spreadsheet.toast(student+" was removed for "+thisRT);
        }

    //Signing out a student from your own RT!

      } else if (!(checkThis == teacher)) {
        daySheet.getRange(studentIndex,columnSign).setFormula('"'+teacher+'"');
        spreadsheet.toast(student+" was signed out for "+thisRT);
      
    //Unsigning out a student from your own RT!
    
      } else {
        daySheet.getRange(studentIndex,columnSign).clearContent();
        spreadsheet.toast(student+" was removed for "+thisRT);
      }

  //Student is signed out to another teacher for today

    } else {
      spreadsheet.toast(student+" was already signed out to "+checkThis);
    }

  //Signing a student out for a future date

  } else {
      var dateColumn = masterSheet.createTextFinder(signoutDate).findNext().getColumn();
      var blueCheck = masterSheet.getRange((studentIndex+1),(dateColumn+1)).getValue();
      var commCheck = masterSheet.getRange((studentIndex+1),(dateColumn+3)).getValue();
      if (color == "blue") {
        var thisCheck = blueCheck;
        var thatCheck = commCheck;
        var offSet = 1;
      } else if (color == "comm") {
        var thisCheck = commCheck;
        var thatCheck = blueCheck;
        var offSet = 3;
      }
      if (thisCheck == "") {
        masterSheet.getRange((studentIndex+1),(dateColumn+offSet)).setFormula('"'+teacher+'"');
        spreadsheet.toast(student+" was signed out for "+signoutDate);
      } else if (thisCheck == teacher) {
        masterSheet.getRange((studentIndex+1),(dateColumn+offSet)).clearContent();
        spreadsheet.toast(student+" was removed for "+signoutDate);
      } else {
        spreadsheet.toast(student+" was already signed out by "+thisCheck+" for "+signoutDate);
      }

  }
}

function signoutbg() {
  signout("blue");
}

function signoutcommon() {
  signout("comm");
}

function signoutboth() {

  //This only works if the student has not been signed out for either RT

  var teacherSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var student = teacherSheet.getRange(signoutNameBox).getValue();
  var teacher = teacherSheet.getRange(teacherNameBox).getValue();
  var signoutDate = teacherSheet.getRange(signoutDateBox).getDisplayValue();
  var studentIndex = daySheet.createTextFinder(student).findNext().getRow();
  if (date == signoutDate) {
     var checkBlue = daySheet.getRange(studentIndex,blueColumnSign).getValue();
     var checkComm = daySheet.getRange(studentIndex,commColumnSign).getValue();
  } else {
      var dateColumn = masterSheet.createTextFinder(signoutDate).findNext().getColumn();
      var checkBlue = masterSheet.getRange((studentIndex+1),(dateColumn+1)).getValue();
      var checkComm = masterSheet.getRange((studentIndex+1),(dateColumn+3)).getValue();
  }

 if ((checkBlue == "" && checkComm == "")||(checkBlue == teacher)&&(checkComm == teacher)) {

    signout("blue");

    teacherSheet.getRange(signoutNameBox).setValue(student);
    teacherSheet.getRange(signoutDateBox).setValue(signoutDate);

    signout("comm");
  } else {
    if(!(checkBlue == "")) {
      spreadsheet.toast(student+" was already signed out to "+checkBlue+" for Blue/Gold RT");
    }
    if(!(checkComm == "")) {
      spreadsheet.toast(student+" was already signed out to "+checkComm+" for Common RT");
    }
  }
}

function clubSignout(clubName, color) {
  var roster = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(clubName);
  var rostercount = roster.getRange('L11').getValue();
    if (rostercount != 0) {
      if(clubName != "OML") {
      var teacher = roster.getRange('L10').getValue();
      var teacherSheet = spreadsheet.getSheetByName(teacher);
        for (var c = 4; c<=(rostercount+3); c++) {
          var student = roster.getRange(c,3).getValue();
          teacherSheet.getRange(signoutNameBox).setValue(student);
          teacherSheet.activate();
          if (color == "B/G") {
            signout("blue");
          } else if (color == "Comm") {
            signout ("comm")
          } else if (color == "Both") {
            signoutboth();
          }
        }
      } else if (clubName == "OML") {
        var omlArray = roster.getRange(4,1,rostercount,10).getValues();
        for (var c = 0; c < rostercount; c++) {
          spreadsheet.getSheetByName(omlArray[c][9]).getRange(signoutNameBox).setValue(omlArray[c][2]);
          spreadsheet.getSheetByName(omlArray[c][9]).activate();
          signout("comm");
        }
      }
      }
}

function runOML() {
  clubSignout("OML","comm");
}

function clubCalendar() {
  if (isSchoolDay) {
    var calendar = spreadsheet.getSheetByName(month);
    var dateIndex = [calendar.createTextFinder(day).findNext().getRow(),calendar.createTextFinder(day).findNext().getColumn()];
    for (var c = 1; c <= 3; c++) {
      var clubName = calendar.getRange(dateIndex[0]+c,dateIndex[1]).getValue();
      var color = calendar.getRange(dateIndex[0]+c,dateIndex[1]+1).getValue();
      if (!(clubName == "")&&!(color == "")) {
        clubSignout(clubName,color);
      }
    }
    dataSheet.getRange('I6').setValue(date);
  }
}

function clubSignoutButton() {
  var clubSign = dataSheet.getRange('B8').getDisplayValue();
  var clubColor = dataSheet.getRange('B9').getDisplayValue();
  clubSignout(clubSign,clubColor);
}
/*
function clubSignoutButton2() {
  var clubSign = dataSheet.getRange('B8').getDisplayValue();
  var clubColor = dataSheet.getRange('B9').getDisplayValue();
  clubSignout2(clubSign,clubColor);
}
*/
function importRoster() {
 var roster = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
 var rostercount = roster.getRange('L11').getValue();
 var clubName = roster.getRange('G1').getValue();
 for (var c = 4; c <= rostercount+3; c++) {
   var student = roster.getRange(c,3).getValue();
   var studentexists = masterSheet.createTextFinder(student).findNext();
   if (!(studentexists == null)) {
      var index = masterSheet.createTextFinder(student).findNext().getRow();
      var clubs = masterSheet.getRange(index, 6).getValue();
      var clubcheck = clubs.includes(clubName);
      if(!(clubcheck == true)) {
        if (clubs == "") {
         clubs = clubName;
        } else {
          clubs = clubs + ", " + clubName;
        }
        masterSheet.getRange(index,6).setValue(clubs);
      } 
    } else {
      spreadsheet.toast(student + " is not in the master sheet!");
      Logger.log(student+ " is not in the master sheet!")
    }
  }
}

function generateMaster() {
  masterSheet.insertColumnsAfter(10,dayCount*4);
  var masterTemplate = masterSheet.getRange(1,7,masterCount+3,4);
  for (var i = 1; i <= dayCount; i++) {
    var loc = 7+4*i;
    masterTemplate.copyTo(masterSheet.getRange(1,loc));
    masterSheet.getRange(1,loc).setFormula("=WORKDAY(G1,"+i+")");
  }

  cleanSlate();

  //Updated MasterFormula Program

  var masterarray = new Array(masterCount);
  for (var y = 0; y < masterCount; y++) {
    masterarray[y] = new Array(4);
    var teach = daySheet.getRange(y+3,4).getValue();
    var ycor = y+3;
      masterarray[y][0]='=VLOOKUP(C'+ycor+", '"+teach+"'!C4:J200, 5,FALSE)"
      //masterarray[y][1]= masterSheet.getRange((ycor+1),(dayColumn+1)).getDisplayValue();
      masterarray[y][2]='=VLOOKUP(C'+ycor+", '"+teach+"'!C4:J200, 7,FALSE)"
      //masterarray[y][3]= masterSheet.getRange((ycor+1),(dayColumn+1)).getDisplayValue();
  }
  daySheet.getRange(3,7,masterCount,4).setFormulas(masterarray);
}

function cleanSlate(){
  var template = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Teacher Template'); //Get the template for teacher rosters 
  var teacherRoster = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Teacher List'); //Get the sheet with a list of teacher names 
  var teacherCount = teacherRoster.getRange('B2').getValue(); //Get the number of teachers from the master sheet 
  for (var t = teacherCount; t >= 1; t = t - 1) { //This loop will create a sheet for every teacher on the list 
    name = teacherRoster.getRange(t+1, 1).getValue(); //gets the name of the teacher 
    if (!(spreadsheet.getSheetByName(name) == null)) {
    spreadsheet.getSheetByName(name).activate();
    spreadsheet.deleteActiveSheet();
    }
    template.activate(); //Go to the template sheet 
    spreadsheet.duplicateActiveSheet(); //Duplicate the template sheet 
    spreadsheet.getActiveSheet().setName(name); //Name that new sheet to match the teacher's name 
    var teacherSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name); //Get the new sheet 
    teacherSheet.getRange('A4').setFormula('=SORT(FILTER(Today!A3:F700,Today!D3:D700="'+name+'"),5,TRUE,1,TRUE)'); //set the filter formula so the students populate 
    teacherSheet.getRange(teacherNameBox).setValue(name); // Put the teacher's name at the top 
    var studentCount = teacherSheet.getRange('L13').getValue(); //Find out how many students are on the normal roster 
      if (!(studentCount == 0)) {
      teacherSheet.getRange('H4').setFormula('=FILTER(Today!H3:H700,Today!D3:D700="'+name+'")');
      teacherSheet.getRange('J4').setFormula('=FILTER(Today!J3:J700,Today!D3:D700="'+name+'")');
      teacherSheet.getRange('G4:G'+(studentCount+3)).setValue('Present');
      teacherSheet.getRange('I4:I'+(studentCount+3)).setValue('Present');
    }
    if (studentCount < 14) {
      var studentrows = 18;
    } else {
      var studentrows = studentCount+4;
    }
    var rowcount = 40 - studentrows;
    teacherSheet.hideRows(studentrows, rowcount);
  }
  dataSheet.getRange('I2').setValue(date);
}

function newDay() {
  
  //Clear all content in the daysheet

  daySheet.getRange(3,blueColumnSign,masterCount,1).clearContent();
  daySheet.getRange(3,commColumnSign,masterCount,1).clearContent();
  daySheet.getRange(3,blueColumnAtt,masterCount,1).setValue("Present");
  daySheet.getRange(3,commColumnAtt,masterCount,1).setValue("Present");
  dataSheet.getRange('I1').setValue(date);
}

function dayFormula() {

  //Updated MasterFormula Program

  var masterarray = new Array(masterCount);
  for (var y = 0; y < masterCount; y++) {
    masterarray[y] = new Array(4);
    var teach = daySheet.getRange(y+3,4).getValue();
    var ycor = y+3;
      masterarray[y][0]='=VLOOKUP(C'+ycor+", '"+teach+"'!C4:J200, 5,FALSE)"
      //masterarray[y][1]= masterSheet.getRange((ycor+1),(dayColumn+1)).getDisplayValue();
      masterarray[y][2]='=VLOOKUP(C'+ycor+", '"+teach+"'!C4:J200, 7,FALSE)"
      //masterarray[y][3]= masterSheet.getRange((ycor+1),(dayColumn+1)).getDisplayValue();
  }
  daySheet.getRange(3,7,masterCount,4).setFormulas(masterarray);
  dataSheet.getRange('I3').setValue(date);
}

function closeSchool() {
  var closeColumn = masterSheet.createTextFinder(offDay).findNext().getColumn();
  var closureArray = new Array(masterCount);
  for (var i = 0; i < masterCount; i++) {
    closureArray[i] = new Array(4);
    closureArray[i][0]= offReason;
    closureArray[i][2]= offReason;
  }
  masterSheet.getRange(4,closeColumn,masterCount,4).setValues(closureArray);
}

function mySheet() {
  var email = Session.getEffectiveUser().getEmail();
  var teacherRoster = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Teacher List'); //Get the sheet with a list of teacher names 
  var row = teacherRoster.createTextFinder(email).findNext().getRow();
  if(!(row==null)) {
    var userName = teacherRoster.getRange(row,1).getValue();
  } else {
    throw new Error("You don't seem to have a RT Sheet, contact John Hartman if this is a mistake");
  }
  spreadsheet.getSheetByName(userName).activate();
} 

function absentSheet() {
  spreadsheet.getSheetByName("Absent Today").activate();
}

function reportSheet() {
  spreadsheet.getSheetByName("Reports").activate();
}

function calendarCall() {
  spreadsheet.getSheetByName(month).activate();
}

function openMaster() {
  spreadsheet.getSheetByName('Master').activate();
}

function openToday() {
  spreadsheet.getSheetByName('Today').activate();
}

function openData() {
  spreadsheet.getSheetByName('Data').activate();
}

function openBulkSheet() {
  spreadsheet.getSheetByName('Bulk Signout').activate();
}

function openAutoSheet() {
  spreadsheet.getSheetByName('Auto-Signouts').activate();
}
/*
function clubSignout2(clubName, color) {
  var roster = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(clubName);
  var status = roster.getRange('L12').getValue();
  var rostercount = roster.getRange('L11').getValue();
  var rosterArray = new Array(0);
  var teacher = roster.getRange('L10').getValue();
  var teacherSheet = spreadsheet.getSheetByName(teacher);
  if (color == "B/G") {
    var columnSign = blueColumnSign;
    var oColumnSign = commColumnSign;
    var thisRT = "Blue/Gold";
  } else if (color == "Comm") {
    var columnSign = commColumnSign;
    var oColumnSign = blueColumnSign;
    var thisRT = "Common";
  }
  if ((status != color) && (status != "Both")) {
    for (var i = 4; i < rostercount+4; i++) {
      if (roster.getRange(i,columnSign).getValue() == "") {
        rosterArray.push(roster.getRange(i,3).getValue());
      }
    }
  } else if ((status == color) || (status == "Both")) {
    if (roster.getRange(i,columnSign).getValue() == teacher) {
        rosterArray.push(roster.getRange(i,3).getValue());      
    }
  }
  Logger.log(rosterArray);
  Logger.log(rosterArray.length);
}
*/
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  var email = Session.getEffectiveUser().getEmail();
  ui.createMenu('RT App')
      .addItem('My Sheet', 'mySheet')
      .addItem('Absent Sheet','absentSheet')
      .addItem('Report Sheet','reportSheet')
      .addItem('RT Calendar','calendarCall')
      .addToUi();
  if(email == "john.hartman@smriders.net") {
      ui.createMenu('RT Admin')
      .addItem('Master Sheet', 'openMaster')
      .addItem('Today Sheet','openToday')
      .addItem('Data Sheet','openData')
      .addItem('Auto Sheet','openAutoSheet')
      .addItem('Bulk Sheet','openBulkSheet')
      .addItem('Signout Club','clubSignoutButton')
      .addItem('Bulk Signout','bulkSignout')
      .addToUi();
  }
}

function autoSignOut() {
  if (isSchoolDay) {
    var weekColumn = 1+(3*weekDayNum);
    var numSign = autoSheet.getRange(1,weekColumn+2).getValue();
    var signArray = autoSheet.getRange(3,weekColumn,numSign,3).getValues();
    for (var i = 0; i < signArray.length; i++) {
      var teacherSheet = spreadsheet.getSheetByName(signArray[i][1]);
      teacherSheet.activate();
      teacherSheet.getRange(signoutNameBox).setValue(signArray[i][0]);
      if (signArray[i][2] == "B/G") {
        signout("blue");
      } else if (signArray[i][2] == "Common") {
        signout("comm");
      } else if (signArray[i][2] == "Both") {
        signoutboth();
      }
    }
  }    
}

function bulkSignout() {
  var bulkSheet = spreadsheet.getSheetByName('Bulk Signout');
  var studentCount = bulkSheet.getRange('J2').getValue();
  var teacher = bulkSheet.getRange('K2').getValue();
  var teacherSheet = spreadsheet.getSheetByName(teacher);
  var rt = bulkSheet.getRange('L2').getValue();
  var roster = bulkSheet.getRange(2,9,studentCount).getValues();
  teacherSheet.activate();
  teacherSheet.getRange('K8').setFormula("=TODAY()");
  for(var c = 0; c < roster.length; c++) {
    Logger.log(roster[c]);
    teacherSheet.getRange(signoutNameBox).setValue(roster[c]);
    switch(rt) {
      case "Comm":
        signout("comm");
        break;
      case "B/G":
        signoutbg("blue");
        break;
    }
  }
}
