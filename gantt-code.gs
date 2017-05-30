/************************************************************
  * For internal use in Asian Tech only.
  * @author Huy Dinh <huydq@asiantech.vn>
  * @copyright PMO Team at Asian Tech.
  * Good luck understanding the hacks and shits and the sugoiness
  * you gonna find in this file.
  ***********************************************************/
var DateDiff = load_DateUtilities();

/************************************************************
  *
  * Settings
  *
  ***********************************************************/
var SYM_OVERDUE = "⚠";
var SYM_DONE = "✔";
var SYM_HOLIDAY_VN = "★";
var SYM_HOLIDAY_JP = "❖";
var SYM_PMO_PROJECT_START = "♳";
var SYM_PMO_IMPLEMENTATION_END = "♵";
var SYM_PMO_DESIGN_END = "♴";
var SYM_PMO_INTERNAL_TEST_END = "♶";
var SYM_PMO_ACCEPTANCE_TEST_END = "♷";
var SYM_PMO_PROJECT_CLOSE = "♸";
var SYM_DELIVER_DATE = "♛";

// Remembers that Month is different to other fields, you need to minus the value by 1.
var PUBLIC_HOLIDAYS_JP = {};
PUBLIC_HOLIDAYS_JP[DateDiff.toString(new Date(2016,0,1,12,0,0,0))] = true; 
PUBLIC_HOLIDAYS_JP[DateDiff.toString(new Date(2016,0,11,12,0,0,0))] = true; 
PUBLIC_HOLIDAYS_JP[DateDiff.toString(new Date(2016,1,11,12,0,0,0))] = true; 
PUBLIC_HOLIDAYS_JP[DateDiff.toString(new Date(2016,2,21,12,0,0,0))] = true; 
PUBLIC_HOLIDAYS_JP[DateDiff.toString(new Date(2016,3,29,12,0,0,0))] = true;  //Golden week
PUBLIC_HOLIDAYS_JP[DateDiff.toString(new Date(2016,4,3,12,0,0,0))] = true;  //Golden week Sun -> Mon
PUBLIC_HOLIDAYS_JP[DateDiff.toString(new Date(2016,4,4,12,0,0,0))] = true;  //Golden week
PUBLIC_HOLIDAYS_JP[DateDiff.toString(new Date(2016,4,5,12,0,0,0))] = true;  //Golden week
PUBLIC_HOLIDAYS_JP[DateDiff.toString(new Date(2016,6,18,12,0,0,0))] = true; 
PUBLIC_HOLIDAYS_JP[DateDiff.toString(new Date(2016,7,11,12,0,0,0))] = true; 
PUBLIC_HOLIDAYS_JP[DateDiff.toString(new Date(2016,8,19,12,0,0,0))] = true; 
PUBLIC_HOLIDAYS_JP[DateDiff.toString(new Date(2016,8,22,12,0,0,0))] = true; 
PUBLIC_HOLIDAYS_JP[DateDiff.toString(new Date(2016,9,10,12,0,0,0))] = true; 
PUBLIC_HOLIDAYS_JP[DateDiff.toString(new Date(2016,10,3,12,0,0,0))] = true; 
PUBLIC_HOLIDAYS_JP[DateDiff.toString(new Date(2016,10,23,12,0,0,0))] = true; 
PUBLIC_HOLIDAYS_JP[DateDiff.toString(new Date(2016,11,23,12,0,0,0))] = true; 

var PUBLIC_HOLIDAYS_VN = {};
PUBLIC_HOLIDAYS_VN[DateDiff.toString(new Date(2016,3,29,12,0,0,0))] = true; //gio to Hung Vuong
PUBLIC_HOLIDAYS_VN[DateDiff.toString(new Date(2016,4,2,12,0,0,0))] = true; //Sat -> Mon
PUBLIC_HOLIDAYS_VN[DateDiff.toString(new Date(2016,4,3,12,0,0,0))] = true; //Sun -> Tue
PUBLIC_HOLIDAYS_VN[DateDiff.toString(new Date(2016,8,2,12,0,0,0))] = true;                          
                   
var BG_PROGRESS = "#55493c";
var BG_PLANNED = "#bcaa89";
var BG_DEFAULT = "#FFFFFF";
var BG_WEEKEND = "#666666";
var BG_WEEKDAY = "#6d645f";
var BG_TODAY = "#FFFFFF";
var COLOR_TRACKDATES = "#bcaa89";
var BG_METADATA = "#6d645f";
var BG_TODAY = "#000000";
var BG_PMO = "#83c8d0";
var BG_PMO_META = "#073763";
var BG_HOLIDAY_META = "#660000";
var BG_HOLIDAY = "#ea9999";
var BG_PARENT = "#dddddd";

/************************************************************
  *
  * Advanced Settings
  *
  ***********************************************************/
var trackDateRowId = 5;
var initRowId = 8;
var initColId = 9;
var holidayDateRowId = 7;
var ROW_ID_PMO = 6;
var ROW_ID_INPUT_PMO = 3;
var TODAY = new Date();
TODAY.setHours(12,0,0,0);

var PMO_MTG1_AXIS = {row:3, col:initColId + 5};
var PMO_MTG2_AXIS = {row:3, col:initColId + 10};
var PMO_MTG3_AXIS = {row:3, col:initColId + 15};
var PMO_MTG4_AXIS = {row:3, col:initColId + 20};
var PMO_MTG5_AXIS = {row:3, col:initColId + 25};
var PMO_MTG6_AXIS = {row:3, col:initColId + 30};
var PRJ_DELIVER_AXIS = {row:3, col:initColId + 39};
var PMO_DATES = {};




/****************************************** 
  *
  * Event Listeners 
  *
  *****************************************/
function onOpen(){
  _loadupPMOSettings();  
  renderTrackDates();
  renderPMODates();
  renderHolidayDates();
  renderGantt();
  
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
  .createMenu('Asian Tech GANTT Chart')
  .addItem('Help', 'displayHelp')
  .addItem('Changelogs', 'displayChangelogs')
  .addToUi();
}

function onEdit(e){
  var sheet = _getSheet();
  var sheetName = e.source.getActiveSheet().getSheetName();

  if(sheet.getSheetName() == sheetName) {
    var range = e.source.getActiveRange();
    _loadupPMOSettings();
    
    if(range.getRow() >= initRowId && range.getColumn() < initColId) {
      renderTasks();
      renderGantt();
    } else if(range.getRow() == ROW_ID_INPUT_PMO) {
      renderPMODates();
      renderGantt();
    } else if(range.getColumn() == initColId - 2 && range.getRow() == 1) {
      renderTrackDates();
      renderTasks();    
      renderPMODates();
      renderHolidayDates();
      renderGantt();
    } else if(range.getColumn() <= initColId && range.getRow() >= initRowId) {
      //Edit on Meta-data columns, let's refresh the gantt for this row
      renderGantt();
    }
  }
}

function displayHelp() {
  var html = HtmlService.createHtmlOutputFromFile('Help')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setTitle('AT Gantt chart Help')
      .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}

function displayChangelogs() {
  var html = HtmlService.createHtmlOutputFromFile('Changelogs')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setTitle('AT GANTT Chart Changelogs')
      .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}

/****************************************** 
  *
  * MAIN LOGICS 
  *
  *****************************************/
function isOverdued(startDate, plannedEndDate, trackDate, percent) {
  return plannedEndDate >= startDate && percent < 1 && trackDate > plannedEndDate && trackDate <= TODAY;
}


function isShowedAsProgressed(startDate, endDate, trackDate, percent, iteratedWeekendsNumber, totalWeekendsNumber) {
  if(trackDate>=startDate && trackDate<=endDate) {
    // Always return first date as progressed if the percent is > 0
    if(DateDiff.inDays(startDate,trackDate) == 0 && percent > 0) {
      return true; 
    }
    
    //Otherwise
    var percentProgress = Math.floor(percent*(DateDiff.inDays(startDate,endDate)+1-totalWeekendsNumber)); //+1 days because duration should be inclusive (not exclusive). Also minus total weekends.
    var trackProgress = DateDiff.inDays(startDate,trackDate)+1-iteratedWeekendsNumber;//+1 days because duration should be inclusive (not exclusive). Also minus partial weekends.
    return percentProgress >= trackProgress && percentProgress != 0;
  }
  return false;
}

function isDone(startDate, endDate, trackDate, percent) {
  return percent >= 1 && trackDate<=endDate && DateDiff.inDays(trackDate,endDate) <= 0
}

function isHoliday(d) {
  return (PUBLIC_HOLIDAYS_JP[DateDiff.toString(d)] 
              || PUBLIC_HOLIDAYS_VN[DateDiff.toString(d)]);
}

function isPMODay(d) {
  return PMO_DATES[DateDiff.toString(d)];
}

function isParent(bg) {
  var rgb = hexToRgb(bg);
/*  if(rgb.r == rgb.b == rgb.g == 255) //if background == gray
  {
    return false;
  }
  */
  if(bg == "White" || bg == "#FFFFFF" || bg == "#ffffff" || bg == "#FFF" || bg == "white")
    return false;
  return true;
}

function _getSheet(){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Gantt Chart");
  if (sheet != null) {
    return sheet;
  }   
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  if(sheets.length == 1) {
    return sheets[0];
  } 
  
  return sheet;
}

function _loadupPMOSettings(){
  var sheet = _getSheet();
  
  var pmo1 = sheet.getRange(PMO_MTG1_AXIS.row,PMO_MTG1_AXIS.col).getValue(); 
  if(typeof pmo1 == 'Date') {
    pmo1.setHours(12,0,0,0);
    PMO_DATES[DateDiff.toString(pmo1)] = 1; 
  }
  var pmo2 = sheet.getRange(PMO_MTG2_AXIS.row,PMO_MTG2_AXIS.col).getValue();   
  if(typeof pmo2 == 'Date') {
    pmo2.setHours(12,0,0,0);
    PMO_DATES[DateDiff.toString(pmo2)] = 1; 
  }
  var pmo3 = sheet.getRange(PMO_MTG3_AXIS.row,PMO_MTG3_AXIS.col).getValue();
  if(typeof pmo3 == 'Date') {
    pmo3.setHours(12,0,0,0);
    PMO_DATES[DateDiff.toString(pmo3)] = 3; 
  }
  var pmo4 = sheet.getRange(PMO_MTG4_AXIS.row,PMO_MTG4_AXIS.col).getValue();  
  if(typeof pmo4 == 'Date') {
    pmo4.setHours(12,0,0,0);
    PMO_DATES[DateDiff.toString(pmo4)] = 4; 
  }
  var pmo5 = sheet.getRange(PMO_MTG5_AXIS.row,PMO_MTG5_AXIS.col).getValue();  
  if(typeof pmo5 == 'Date') {
    pmo5.setHours(12,0,0,0);
    PMO_DATES[DateDiff.toString(pmo5)] = 5; 
  }  
  var pmo6 = sheet.getRange(PMO_MTG6_AXIS.row,PMO_MTG6_AXIS.col).getValue();
  if(typeof pmo6 == 'Date') {
    pmo6.setHours(12,0,0,0);
    PMO_DATES[DateDiff.toString(pmo6)] = 6; 
  }    
  var deadline = sheet.getRange(PRJ_DELIVER_AXIS.row,PRJ_DELIVER_AXIS.col).getValue();
  if(typeof deadline == 'Date') {
    deadline.setHours(12,0,0,0);
    PMO_DATES[DateDiff.toString(deadline)] = 99; 
  }    
}


/****************************************** 
  *
  * UI 
  *
  *****************************************/


/**
  *
  */
function renderTrackDates(){
  var sheet = _getSheet();

  var rangeTrackDates = sheet.getRange(trackDateRowId,initColId,1,sheet.getMaxColumns()-initColId+1);
  var trackDatesCount = rangeTrackDates.getNumColumns();

  var trackDates = rangeTrackDates.getValues();
  var trackDateBgs = rangeTrackDates.getBackgrounds();  
  
  for (var i = 0; i < trackDatesCount; i++) {
    trackDates[0][i] = (new Date(trackDates[0][i]));
    trackDates[0][i].setHours(12,0,0,0);
    if (DateDiff.inDays(trackDates[0][i],TODAY) == 0) { //Today
       trackDateBgs[0][i] = BG_TODAY;
       renderTodayVerticalLine(i+initColId); //to save us from another loop somewhere else later in the app we invoke the rendering of "today" column here.
    }
    else if(DateDiff.isWeekend(trackDates[0][i])) { //Weekends
       trackDateBgs[0][i] = BG_WEEKEND;
    }  
    else { //Weekdays
       trackDateBgs[0][i] = BG_WEEKDAY;
    }
  }
  rangeTrackDates.setFontColor(COLOR_TRACKDATES);
  rangeTrackDates.setBackgrounds(trackDateBgs);
  SpreadsheetApp.flush();
}

/**
  *
  */
function renderPMODates(){
  var sheet = _getSheet();
  var pmoDefinedDates = _loadupPMOSettings();
  var rangeTrackDates = sheet.getRange(trackDateRowId,initColId,1,sheet.getMaxColumns()-initColId+1);
  var trackDates = rangeTrackDates.getValues();
  var trackDatesCount = rangeTrackDates.getNumColumns();

  var rangePMORow = sheet.getRange(ROW_ID_PMO,initColId,1,sheet.getMaxColumns()-initColId+1);
  var pmos = rangePMORow.getValues();
  var pmoBackgrounds = rangePMORow.getBackgrounds();
  
  for (var i = 0; i < trackDatesCount; i++) {
    trackDates[0][i].setHours(12,0,0,0);
    pmos[0][i] = "";
    pmoBackgrounds[0][i] = BG_METADATA;
    if(PMO_DATES[DateDiff.toString(trackDates[0][i])] == 1) {
      pmos[0][i] = SYM_PMO_PROJECT_START;
      pmoBackgrounds[0][i] = BG_PMO_META;      
    } else if(PMO_DATES[DateDiff.toString(trackDates[0][i])] == 2) {
      pmos[0][i] = SYM_PMO_DESIGN_END;
      pmoBackgrounds[0][i] = BG_PMO_META;      
    } else if(PMO_DATES[DateDiff.toString(trackDates[0][i])] == 3) {
      pmos[0][i] = SYM_PMO_IMPLEMENTATION_END;
      pmoBackgrounds[0][i] = BG_PMO_META;      
    } else if(PMO_DATES[DateDiff.toString(trackDates[0][i])] == 4) {
      pmos[0][i] = SYM_PMO_INTERNAL_TEST_END;
      pmoBackgrounds[0][i] = BG_PMO_META;      
    } else if(PMO_DATES[DateDiff.toString(trackDates[0][i])] == 5) {
      pmos[0][i] = SYM_PMO_ACCEPTANCE_TEST_END;
      pmoBackgrounds[0][i] = BG_PMO_META;      
    } else if(PMO_DATES[DateDiff.toString(trackDates[0][i])] == 6) {
      pmos[0][i] = SYM_PMO_PROJECT_CLOSE;
      pmoBackgrounds[0][i] = BG_PMO_META;      
    } else if(PMO_DATES[DateDiff.toString(trackDates[0][i])] == 99) {
      pmos[0][i] = SYM_DELIVER_DATE;
      pmoBackgrounds[0][i] = BG_PMO_META;  
    }    
  }
  rangePMORow.setValues(pmos);
  rangePMORow.setBackgrounds(pmoBackgrounds);   
  SpreadsheetApp.flush();  
}
  
/**
  *
  */
function renderHolidayDates(){
  var sheet = _getSheet();
  var rangeTrackDates = sheet.getRange(trackDateRowId,initColId,1,sheet.getMaxColumns()-initColId+1);
  var trackDates = rangeTrackDates.getValues();
  var trackDatesCount = rangeTrackDates.getNumColumns();

  var rangeHolidaysRow = sheet.getRange(holidayDateRowId,initColId,1,sheet.getMaxColumns()-initColId+1);
  var holidays = rangeHolidaysRow.getValues();
  var holidayBackgrounds = rangeHolidaysRow.getBackgrounds();
  for (var i = 0; i < trackDatesCount; i++) {
    trackDates[0][i].setHours(12,0,0,0);
    holidayBackgrounds[0][i] = BG_METADATA;
    holidays[0][i] = "";
    if(PUBLIC_HOLIDAYS_JP[DateDiff.toString(trackDates[0][i])]) {
      holidays[0][i] = SYM_HOLIDAY_JP; 
      holidayBackgrounds[0][i] = BG_HOLIDAY_META;
    } else if (PUBLIC_HOLIDAYS_VN[DateDiff.toString(trackDates[0][i])]) {
      holidays[0][i] += SYM_HOLIDAY_VN;
      holidayBackgrounds[0][i] = BG_HOLIDAY_META;      
    }
  }
  rangeHolidaysRow.setValues(holidays);
  rangeHolidaysRow.setBackgrounds(holidayBackgrounds); 
  SpreadsheetApp.flush();  
}

/**
  * @_coldId Column ID of today
  */
function renderTodayVerticalLine(_colId) {
  var sheet = _getSheet();
  var rangeGantt = sheet.getRange(initRowId,initColId,sheet.getMaxRows()-initRowId+1,sheet.getMaxColumns()-initColId+1);
  rangeGantt.setBorder(false, false, false, false, false, false);
  
  //Set border for today column
  var rangeTodayColumn = sheet.getRange(initRowId,_colId,sheet.getMaxRows()-trackDateRowId+1, 1);
  rangeTodayColumn.setBorder(null, true, null, true, true, false, "black", null);
  
  //Reset backgrounds for dates meta rows
  var rangeMetaDateRows = sheet.getRange(trackDateRowId+1,initColId,initRowId-trackDateRowId,sheet.getMaxColumns()-initColId+1);
  rangeMetaDateRows.setBackground(BG_METADATA);
  
  //set background for today column within meta rows
  var rangeTodayMetaDatesCol = sheet.getRange(trackDateRowId,_colId,initRowId-trackDateRowId, 1);
  rangeTodayMetaDatesCol.setBackground(BG_TODAY);

  SpreadsheetApp.flush();
}

function renderTasks() {
  var sheet = _getSheet();  
  var rangeTask = sheet.getRange(initRowId,1,sheet.getMaxRows()-initRowId+1,initColId-1);
  var taskValues = rangeTask.getValues();
  var taskBgs = rangeTask.getBackgrounds();

  var dataForNextParent = {'startDate':null,'endDate':null,'sumProductDuration':0, 'sumDuration':0,'childrenCount':0};
  var taskRows = rangeTask.getNumRows();
  
// Loops through Task from bottom up
  for (var i = taskRows-1; i >= 0; i--) {
    if(!isParent(taskBgs[i][0])) {
      var startDate = taskValues[i][initColId-7];
      var duration = Math.ceil(parseFloat(taskValues[i][initColId-3]));     
      if(!startDate || !duration)
        continue; //skip this row.
      else 
        startDate.setHours(12,0,0,0);  
      
      var percent = parseFloat(taskValues[i][initColId-2]);
      if(!percent)
        percent = 0.0;
      if(percent > 1)
        percent = 1.0;
      //Always recalculate the End Date of a task based on Duration and Start Date
      var endDate = new Date(startDate);
      endDate.setHours(12,0,0,0);  
      DateDiff.addDays(endDate,duration-1);
      
      //Check Start Date and make sure it's not during the weekend
      if(DateDiff.isSaturday(startDate)){
        DateDiff.addDays(startDate,2);
      } else if(DateDiff.isSunday(startDate)){
        DateDiff.addDays(startDate,1);
      }
      
      var weekendOverflowBuffer = 0; 
      var weekendOverflowTotalCount = 0;
      var overflowedEndDate = new Date(endDate);
      var overflowedStartDate = new Date(startDate);
      var totalWeekendsBetween = 0; 
      var overflowedWeekendDays = 0;
      var iteratingTotalWeekendsBetween = 0;
      
      do {
        overflowedWeekendDays = DateDiff.numberOfWeekendsBetween(overflowedStartDate, overflowedEndDate);
        overflowedStartDate = new Date(overflowedEndDate);
        totalWeekendsBetween += overflowedWeekendDays;
        
        if(DateDiff.isSaturday(overflowedEndDate)) {
          DateDiff.addDays(overflowedEndDate,2);
          overflowedWeekendDays -= 1;
        } else if (DateDiff.isSunday(overflowedEndDate)) {
          DateDiff.addDays(overflowedEndDate,2);
          overflowedWeekendDays -= 2;
        }
        if(DateDiff.isSaturday(overflowedStartDate)) {
          DateDiff.addDays(overflowedStartDate,2);
        } else if (DateDiff.isSunday(overflowedStartDate)) {
          DateDiff.addDays(overflowedStartDate,1);
        }
        
        DateDiff.addDays(overflowedEndDate, overflowedWeekendDays);  
      } while(overflowedWeekendDays > 0);
      endDate = new Date(overflowedEndDate); endDate.setHours(12,0,0,0);
      taskValues[i][3] = endDate;
      taskValues[i][2] = startDate;
      taskValues[i][7] = percent;
    }
    // Save data for next parent task
    if(isParent(taskBgs[i][0])) {
      taskValues[i][2] = new Date(dataForNextParent.startDate);
      taskValues[i][3] = new Date(dataForNextParent.endDate);
      taskValues[i][7] = dataForNextParent.sumProductDuration / dataForNextParent.sumDuration;
      taskValues[i][6] = DateDiff.inDays(dataForNextParent.startDate,dataForNextParent.endDate) - DateDiff.numberOfWeekendsBetween(dataForNextParent.startDate, dataForNextParent.endDate) + 1;
      dataForNextParent = {'startDate':null,'endDate':null,'sumProductDuration':0,'sumDuration':0,'childrenCount':0}; //Reset
    } else {
      dataForNextParent.startDate =  (dataForNextParent.startDate === null || startDate < dataForNextParent.startDate) ? startDate : dataForNextParent.startDate; //get Min for Start Date
      dataForNextParent.endDate =  (dataForNextParent.endDate === null || endDate > dataForNextParent.endDate) ? endDate : dataForNextParent.endDate; //get Max for End Date
      dataForNextParent.sumProductDuration += duration * percent; 
      dataForNextParent.sumDuration += duration; 
      dataForNextParent.childrenCount += 1;
    }
  }
  rangeTask.setValues(taskValues); 
  rangeTask.setBackgrounds(taskBgs); 
  SpreadsheetApp.flush();
}

/**
  * Meet and Juice of this file.
  */
function renderGantt() {
  var sheet = _getSheet();
  var rangeTrackDates = sheet.getRange(trackDateRowId,initColId,1,sheet.getMaxColumns()-initColId+1);
  var rangeTask = sheet.getRange(initRowId,1,sheet.getMaxRows()-initRowId+1,sheet.getMaxColumns()-initColId+1);
  var rangeGantt = sheet.getRange(initRowId,initColId,sheet.getMaxRows()-initRowId+1,sheet.getMaxColumns()-initColId+1);
  //Convert data dictionary from range to values to reduce IO during loops
  var taskValues = rangeTask.getValues();
  var taskBgs = rangeTask.getBackgrounds();
  var valuesTrackDates = rangeTrackDates.getValues();
  var ganttSymbols = rangeGantt.getValues();
  var ganttBgs = rangeGantt.getBackgrounds();
  var ganttCols = rangeGantt.getNumColumns();
  var ganttRows = rangeGantt.getNumRows();

  // Loops through Task from bottom up
  for (var i = ganttRows-1; i >= 0; i--) {
    var startDate = taskValues[i][initColId-7]; 
    var endDate = taskValues[i][initColId-6]; 
    var skipRow = false;
    if(!startDate || !(startDate instanceof Date) 
      ||!endDate || !(endDate instanceof Date))
      skipRow = true; //skip this row.
    else {
      startDate.setHours(12,0,0,0); 
      endDate.setHours(12,0,0,0);  
    }
    var percent = parseFloat(taskValues[i][initColId-2]);
    if(!percent)
      percent = 0.0;
    
    var iteratingTotalWeekendsBetween = 0;
    // Loops through Days
    for (var j = 0; j <= ganttCols-1; j++) {    
      var symbol = "";
      var background = BG_DEFAULT;
      if(isParent(taskBgs[i][0])) {
        background = BG_PARENT;
      }

      var trackDate = valuesTrackDates[0][j]; 
      if(!trackDate || !(trackDate instanceof Date)) {
        skipRow = true;
      } else {
        trackDate.setHours(12,0,0,0);
      }
      
      if(DateDiff.isWeekend(trackDate)) {
        background = BG_WEEKEND;
      }
      
      // Render background colors to show progress and planned duration
      if(isHoliday(trackDate)) {
        background = BG_HOLIDAY;  
      }
      
      // Render background color for PMO Dates 
      if(isPMODay(trackDate)){
        background = BG_PMO;
      }
      
      if(!skipRow) {
         var totalWeekendsBetween = DateDiff.numberOfWeekendsBetween(startDate, endDate);

        //Render Overdue
        if(isOverdued(startDate, endDate, trackDate, percent)) {
          symbol = SYM_OVERDUE;
        }
        
        // When track date falls BETWEEN Start and End dates
        if(trackDate>=startDate && trackDate<=endDate) {
          if(DateDiff.isWeekend(trackDate)) {
            iteratingTotalWeekendsBetween++;
          } else {
            //If this task fall over weekend, we remember this and still render it as weekend and fill another day later
            background = BG_PLANNED;
            
            if (isDone(startDate, endDate, trackDate, percent)) {          
              symbol = SYM_DONE;
            }
            
            if(symbol == SYM_DONE || isShowedAsProgressed(startDate, endDate, trackDate, percent, iteratingTotalWeekendsBetween, totalWeekendsBetween)) {
              background = BG_PROGRESS;
            }
          }
        }
      }
      ganttSymbols[i][j] = symbol;
      ganttBgs[i][j] = background;
    }
  }
  rangeTask.setValues(taskValues);
  rangeGantt.setValues(ganttSymbols); 
  rangeGantt.setBackgrounds(ganttBgs); 
  SpreadsheetApp.flush();
}
