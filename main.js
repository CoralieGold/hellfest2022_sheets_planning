// *****************************************************************************************
// ********************************** GLOBAL VARIABLES *************************************
// *****************************************************************************************

var UI = SpreadsheetApp.getUi(); 
var HELLFEST_SPREADSHEET = SpreadsheetApp.getActiveSpreadsheet();
var PLANNING_SHEET = HELLFEST_SPREADSHEET.getSheetByName('Planning');
var SETTINGS_SHEET = HELLFEST_SPREADSHEET.getSheetByName('Réglages');

var MIN_NOTE = SETTINGS_SHEET.getRange(7, 4).getValue(); 
var MAX_NOTE = SETTINGS_SHEET.getRange(6, 4).getValue();
var CALENDAR_ID = SETTINGS_SHEET.getRange(8, 4).getValue();

var NUMBER_OF_COL = 18;
var NUMBER_OF_ROW = 20;
var FIRST_ROW = 2;

var MIN_DAY_HOUR = '00:00';
var MAX_DAY_HOUR = '09:00';


var DAYS = {
  "Vendredi 1": {
    date: "2022-06-17",
    isShow: true,
    sheet: HELLFEST_SPREADSHEET.getSheetByName("Vendredi 1")
  },
  "Samedi 1": {
    date: "2022-06-18",
    isShow: true,
    sheet: HELLFEST_SPREADSHEET.getSheetByName("Samedi 1")
  },
  "Dimanche 1": {
    date: "2022-06-19",
    isShow: true,
    sheet: HELLFEST_SPREADSHEET.getSheetByName("Dimanche 1")
  },
  "Lundi 1": {
    date: "2022-06-20",
    isShow: false
  },
  "Jeudi": {
    date: "2022-06-23",
    isShow: true,
    sheet: HELLFEST_SPREADSHEET.getSheetByName("Jeudi")
  },
  "Vendredi 2": {
    date: "2022-06-24",
    isShow: true,
    sheet: HELLFEST_SPREADSHEET.getSheetByName("Vendredi 2")
  },
  "Samedi 2": {
    date: "2022-06-25",
    isShow: true,
    sheet: HELLFEST_SPREADSHEET.getSheetByName("Samedi 2")
  },
  "Dimanche 2": {
    date: "2022-06-26",
    isShow: true,
    sheet: HELLFEST_SPREADSHEET.getSheetByName("Dimanche 2")
  },
  "Lundi 2": {
    date: "2022-06-27",
    isShow: false
  }
}

var STAGES = {
              "Mainstage 1": {
                name: 'Mainstage 1',
                color: '#0000ff',
                event: '9'
              },
              "Mainstage 2": {
                name: 'Mainstage 2',
                color: '#999999',
                event: '8'
              },
              "Altar": {
                name: 'Altar',
                color: '#ff0000',
                event: '11'
              },
              "Valley": {
                name: 'Valley',
                color: '#b45f06',
                event: '6'
              },
              "Temple": {
                name: 'Temple',
                color: '#4a86e8',
                event: '1'
              },
              "Warzone": {
                name: 'Warzone',
                color: '#38761d',
                event: '10'
              }
            }


// *****************************************************************************************
// ************************************* CALLBACKS *****************************************
// *****************************************************************************************

function onEdit(e) {
  /* onEdit on N1 cell, on Planning sheet. Replace menu on mobile devices.
   */

  if(e.source.getActiveSheet().getName() != 'Planning')
    return

  if(e.range.getA1Notation() != 'N1')
    return
    
  var functionName;
  if(e.value == 'Planning : Mettre à jour')
    functionName = 'updatePlanning';
  else if(e.value == 'Planning : Remettre à zéro')
    functionName = 'clearPlanning';
  else if(e.value == 'Google Agenda : Mettre à jour')
    functionName = 'addToCalendar';
  else if(e.value == 'Google Agenda : Remettre à zéro')
    functionName = 'clearCalendar';
  else if(e.value == 'Exporter un PDF')
    functionName = 'generatePdf';
  else if(e.value == 'Template : Créer un tableau')
    functionName = 'createSpreadsheet';
  else if(e.value == 'Template : Mettre à jour')
    functionName = 'fixPlanningWithTemplate';
  
  if(functionName) {        
    eval(functionName)();
    e.range.clear();
  }

}


function onOpen() {
  /* Add menu with pdf export and planning update on file open.
   */

  var planningMenu = [{name: 'Planning : Mettre à jour', functionName: 'updatePlanning'},
                      {name: 'Planning : Remettre à zéro', functionName: 'clearPlanning'},
                      {name: 'Google Agenda : Mettre à jour', functionName: 'addToCalendar'},
                      {name: 'Google Agenda : Remettre à zéro', functionName: 'clearCalendar'},
                      {name: 'Exporter un PDF', functionName: 'generatePdf'},
                      {name: 'Template : Créer un tableau', functionName: 'createSpreadsheet'},
                      {name: 'Template : Mettre à jour', functionName: 'fixPlanningWithTemplate'}];
  HELLFEST_SPREADSHEET.addMenu('Planning', planningMenu);  
}


// *****************************************************************************************
// ************************************* FUNCTIONS *****************************************
// *****************************************************************************************


// ***************************************** UTILS *****************************************

function getSheetsData() {
  var sheetsData = {};
  var stageIdx = 0;
  var stage;
  
  for(const [day, day_value] of Object.entries(getShowDays())) {
    sheetsData[day] = {};

    // Get stages
    day_value.sheet.getRange(day + "!A1:R1").getValues().forEach(function(row) {
      row.forEach(function(stage) {
        if(stage !== "") sheetsData[day][stage] = [];
      });
    });

    // Get bands, hours and notation
    day_value.sheet.getRange(day + "!A2:R20").getValues().forEach(function(row) {
      stageIdx = 0; 
      row.forEach(function(_, colIdx, colArray) {
        stage = Object.keys(STAGES)[stageIdx];
        if(colIdx % 3 == 0) {
          sheetsData[day][stage].push({
            hour: colArray[colIdx],
            band: colArray[colIdx + 1],
            note: colArray[colIdx + 2],
            stage: stage
          });
          stageIdx ++;
        }   
      });
    });
  }

  return sheetsData;
}


function getShowDays() {
  var show_days = {};
  for (const [key, value] of Object.entries(DAYS)) {
    if(value.isShow == true)
      show_days[key] = value;
  }

  return show_days;
}


function getStageFromCellColor(cellColor) {
  /* Get stage's name from an hexa color.
   * 
   * @param   {String}       cellColor  Hexa color of a stage
   * 
   * @return  {String}                  Name of stage corresponding to given color
   */

  for(const [stage, value] of Object.entries(STAGES)) {
    if(value.color == cellColor) return stage; 
  }
  
}


function getFolder() {
  var folder;

  var hellfestSpreadSheetId = DriveApp.getFileById(HELLFEST_SPREADSHEET.getId());

  // Get folder containing spreadsheet to save pdf in.
  var parents = hellfestSpreadSheetId.getParents();
  if (parents.hasNext()) {
    folder = parents.next();
  }
  else {
    folder = DriveApp.getRootFolder();
  }

  return folder;
}


// ***************************************** MENU *****************************************

function updatePlanning() {
  /* Find bands with max notes for each hours in each days
   * and fill planning spreadsheet.
   */

  var planningValues = [];
  var planningBackgrounds = [];
  var planningColors = [];
  var perHour = {};
  var first, second;
  var hourIdx = 0;
  var dayIdx = 0;

  var sheetsData = getSheetsData();

  for(var day in sheetsData) {
    perHour = {};
    planningValues = [];
    planningBackgrounds = [];
    planningColors = [];
    planningWeights = [];

    // Get all bands at the same hour
    for(var stage in sheetsData[day]) {
      sheetsData[day][stage].forEach(function(band) {
        if(perHour[band.hour] == undefined)
          perHour[band.hour] = [];
        if(band.band != "") {
          perHour[band.hour].push(band);
        }
      });
    }

    // Get two bands with higher notes
    hourIdx = 0;
    for(var hour in perHour) {
      first = {
        note: 0,
        hour: "",
        band: "",
        stage: ""
      };
      second = {
        note: 0,
        hour: "",
        band: "",
        stage: ""
      };

      perHour[hour].forEach(function(band) {
        if(band.note != "" && band.note >= MIN_NOTE) {
          if(band.note > first.note)
            first = band;
    
          else if(band.note <= first.note && band.note >= second.note)
            second = band;
        }
      });
      
      // Set planning values
      planningValues[hourIdx] = [
        first.hour,
        (first.band != "") ? first.band + " [" + first.note + "]" : "",
        second.hour,
        (second.band != "") ? second.band + " [" + second.note + "]" : ""
      ]

      planningBackgrounds[hourIdx] = [
        "#FFFFFF",
        (first.stage != "") ? STAGES[first.stage].color : "#FFFFFF",
        "#FFFFFF",
        (second.stage != "") ? STAGES[second.stage].color : "#FFFFFF"
      ]

      planningColors[hourIdx] = [
        "#000000",
        "#FFFFFF",
        "#000000",
        "#FFFFFF"
      ]

      hourIdx ++;

    }  
  
    // Fill planning
    PLANNING_SHEET.getRange(2, dayIdx*4 + 1, hourIdx, 4)
      .setValues(planningValues)
      .setBackgrounds(planningBackgrounds)
      .setFontFamily("Open Sans")
      .setFontSize(11)
      .setFontColors(planningColors)
      .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle")

    dayIdx ++;
  }
 
  UI.alert('Planning mis à jour !');
}


function generatePdf() {
  /* From: https://gist.github.com/primaryobjects/6370689c6f5fd3799ea53f89551eced7
   * 
   * Create a pdf file in 
   */

  // Get folder containing spreadsheet to save pdf in.
  var folder = getFolder();
  
  // Copy whole spreadsheet.
  var destSpreadsheet = SpreadsheetApp.open(DriveApp.getFileById(HELLFEST_SPREADSHEET.getId()).makeCopy('tmp_convert_to_pdf', folder))

  // Delete redundant sheets.
  var sheets = destSpreadsheet.getSheets();
  for (i = 0; i < sheets.length; i++) {
    if (sheets[i].getSheetName() != 'Planning')
      destSpreadsheet.deleteSheet(sheets[i]);
  }
  
  var destSheet = destSpreadsheet.getSheets()[0];

  // Repace cell values with text (to avoid broken references).
  var sourceRange = PLANNING_SHEET.getRange(1, 1, PLANNING_SHEET.getMaxRows(), 12);
  var sourcevalues = sourceRange.getValues();
  var destRange = destSheet.getRange(1, 1, destSheet.getMaxRows(), 12);
  destRange.setValues(sourcevalues);

  // Save to pdf.
  var theBlob = destSpreadsheet.getBlob().getAs('application/pdf').setName(HELLFEST_SPREADSHEET.getName());
  var newFile = folder.createFile(theBlob);
  
  // Delete the temporary sheet.
  DriveApp.getFileById(destSpreadsheet.getId()).setTrashed(true);
  
  UI.alert('PDF créé : ' + HELLFEST_SPREADSHEET.getName() + '.pdf');
}


function clearPlanning() {
  /* Empty planning and remove all cells and style.
   */
  PLANNING_SHEET.getRange("Planning!A2:AC20")
    .clearContent()
    .setBackground("#FFFFFF")
    .setFontWeight("normal");
}


function addToCalendar() {
  /* Add all events in planning sheet in a Google Calendar.
   */

  if(!CALENDAR_ID) {
    UI.alert('Merci d\'ajouter un ID de calendrier dans la feuille "Réglages".');
    return;
  }

  var calendar = CalendarApp.getCalendarById(CALENDAR_ID);
  
  if(!calendar) {
    UI.alert('Le calendrier ' + CALENDAR_ID + ' n\'a pas été trouvé.');
    return;
  }
  
  var day, summary, startTime, endTime, description, location, eventOptions, event, hours, rangeValue, eventDay;
  var show_days = getShowDays();

  var planningRange = PLANNING_SHEET.getRange(FIRST_ROW, 1, NUMBER_OF_ROW, NUMBER_OF_COL);
  var planningDays = PLANNING_SHEET.getRange(1, 1, 1, NUMBER_OF_COL);
  var planningValues = planningRange.getValues();
  var planningBackgrounds = planningRange.getBackgrounds();

  for(var row = 0; row < NUMBER_OF_ROW - FIRST_ROW; row++) {
    for(var col = 0; col < NUMBER_OF_COL; col ++) {
      rangeValue = planningValues[row][col];
      if(!rangeValue) continue;

      if(col % 2 != 0) {
        summary = 'Concert : ' + rangeValue;
        location = getStageFromCellColor(planningBackgrounds[row][col]); 
        description = 'GROUPE : ' + rangeValue + '\nSTAGE : ' + location;
        
        eventOptions = {
          'location': location,
          'description': description,
        }
        
        event = calendar.createEvent(summary, startTime, endTime, eventOptions);
        event.setColor(STAGES[location].event);
      }
      
      // Get hour info
      else {
        day = col % Object.keys(show_days).length;

        hours = rangeValue.split(' > ');
        if(hours[0] >= MIN_DAY_HOUR && hours[0] <= MAX_DAY_HOUR) eventDay = day + 1;
        else eventDay = day;
        startTime = new Date(show_days[Object.keys(show_days)[eventDay]].date + 'T' + hours[0]);
        endTime = new Date(show_days[Object.keys(show_days)[eventDay]].date + 'T' + hours[1]);

      }
    }
  }        
  
  UI.alert('Le calendrier ' + CALENDAR_ID + ' a bien été mis à jour !');
}


function clearCalendar() {
  /* Remove all events between friday and sunday in Google Calendar.
   */

  if(!CALENDAR_ID) {
    UI.alert('Merci d\'ajouter un ID de calendrier dans la feuille "Réglages".');
    return;
  }
  
  var calendar = CalendarApp.getCalendarById(CALENDAR_ID);
  
  if(!calendar) {
    UI.alert('Le calendrier ' + CALENDAR_ID + ' n\'a pas été trouvé.');
    return;
  }
   
  var firstDay = new Date(DAYS[Object.keys(DAYS)[0]].date + 'T00:00:00');
  var lastDay = new Date(DAYS[Object.keys(DAYS)[Object.keys(DAYS).length - 1]].date + 'T23:59:59');
  var events = calendar.getEvents(firstDay, lastDay, {search: 'Concert :'});
  
  for(var e = 0; e < events.length; e++) {
    events[e].deleteEvent();
  }

  UI.alert('Le calendrier ' + CALENDAR_ID + ' a bien été vidé !');
}

                     
function createSpreadsheet() {
  /* Show a pop up menu to enter a user name and create a new Hellfest spreadsheet
   */
  var ui = SpreadsheetApp.getUi(); // Same variations.

  var result = ui.prompt('Entre ton nom :', ui.ButtonSet.OK_CANCEL);

  // Process the user's response.
  var button = result.getSelectedButton();
  var text = result.getResponseText();

  // User clicked "OK".
  if (button == ui.Button.OK) {    
    var hellfestSpreadSheetId = DriveApp.getFileById(HELLFEST_SPREADSHEET.getId());
    var folder = getFolder();

    var documentId = hellfestSpreadSheetId.makeCopy().getId();
  
    // Rename the copied file and move it inside correct folder                  
    var copyFile = DriveApp.getFileById(documentId);
    copyFile.setName('Hellfest_' + text);
    folder.addFile(copyFile);
    DriveApp.getRootFolder().removeFile(copyFile);
  }
}

function getHellfestTemplateId() {

  var folder = getFolder();
  var files = folder.getFilesByName('Hellfest_Template');
  if (files.hasNext()) {
    var file = files.next();
    return file.getId();
  }
}


function fixPlanningWithTemplate() {

  // Récup l'info template
  // Récupe l'info current
  // Remplacer info current par template, mais mettre la bonne note

  var sheet, sheetTpl;
  var templateSpreadsheet = SpreadsheetApp.openById(getHellfestTemplateId());
  var show_days = getShowDays();
  var name = "";
  var sheetValues = [];
  var templateBands = [];
  var currentBands = {};
  
  for(const [day, day_value] of Object.entries(show_days)) {
    sheet = day_value.sheet;
    sheetTpl = templateSpreadsheet.getSheetByName(day);
    if(!sheet || !sheetTpl) continue;

    sheetTplValues = sheetTpl.getRange(1, 1, NUMBER_OF_ROW, NUMBER_OF_COL).getValues();
    sheetValues = sheet.getRange(1, 1, NUMBER_OF_ROW, NUMBER_OF_COL).getValues(); 

    for(var row = 1; row < NUMBER_OF_ROW; row++) {
      for(var col = 0; col < NUMBER_OF_COL; col += 3) {     
        templateBands.push({hour:  sheetTplValues[row][col],
                            name:  sheetTplValues[row][col + 1],
                            note:  sheetTplValues[row][col + 2], 
                            stage: sheetTplValues[0][col],
                            day: day,
                            col: col + 1,
                            row: row + 1});

        name = sheetValues[row][col + 1];
        currentBands[name] = {hour:  sheetValues[row][col],
                              name:  name,
                              note:  sheetValues[row][col + 2], 
                              stage: sheetValues[0][col],
                              day: day,
                              col: col + 1,
                              row: row + 1};
      }
    }

    sheet.getRange(FIRST_ROW, 1, NUMBER_OF_ROW, NUMBER_OF_COL).setValue("");
  }
  
  sheetValues = [];
  for(var index = 0; index < templateBands.length; index ++) {
    var templateBand = templateBands[index];
    sheet = show_days[templateBand.day].sheet;
    sheetValues = [
      [templateBand.hour, templateBand.name],
    ]

    if(templateBand.name == '' && !(templateBand.name in currentBands)) sheetValues[0].push("");
    else sheetValues[0].push(currentBands[templateBand.name].note);
    
    sheet.getRange(templateBand.row, templateBand.col, 1, 3).setValues(sheetValues);
  }
  
  print('Fichier mis à jour par rapport au template !');

}


function print(toPrint) {
  SpreadsheetApp.getUi().alert(toPrint);
}
