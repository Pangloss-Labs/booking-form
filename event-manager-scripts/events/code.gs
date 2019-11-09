var triggerDelay = 5; // by min
var notificationEmailIT = "sample@sample.org"; //IT 
var notificationEmailWorkflow = "conciergerie@sample.org"; //conciergerie 
var columnIndexForState = 0;
var criterias = ["submitted","validated","processed","published","deleted"];
var eventSpreadsheetUrl = "https://docs.google.com/spreadsheets/d/XXXXXXspreadsheetID/edit#gid=0"; // To define
var topic = ["Online","Online-Event"];
var publicationMsg =["<p>Votre réservation est prise en compte</p>","<p>Votre réservation est en cours de validation par la conciergerie</p>"];
var signatureEmail = "<p>Attention: un paiement Helloasso est parfois necessaire pour finaliser l'utilisation d'un équipement ou l'accès à une salle de réunion</p><p>Pour toute information complémentaire ou modification, veuillez contracter la conciergerie : conciergerie@sample.org ou par téléphone au +33 (0)4 50 XXXXXXXX</p><br><p>L'équipe conciergie</p>";

//----------------------------
// Create spreadsheet menu on open
//----------------------------
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('Pangloss Event Manager').addItem('Start publication', 'publicationProcess')
    .addSeparator()
    .addSubMenu(ui.createMenu('Sub function').addItem('Delete events', 'deleteEvents').addItem('Re-activate triggers', 'starterFunction').addItem('Kill Triggers', 'deleteTrigger').addItem('get event id', 'getEventsId')).addToUi();
}

//----------------------------
// Starter function (to start)
//----------------------------
function starterFunction() {
  refreshUserProps(); // create properties
  Logger.log("properties created");
  createTrigger(); // create trigger to run program automatically
  Logger.log("Trigger created");
  // @notification to IT
  
  MailApp.sendEmail({
   to: notificationEmailIT,
   subject: "PANGLOSS - Event Bot & his Trigger - Google Script Apps is running under the spreadsheet Events",
   htmlBody: "<p>properties created</p><p>Trigger created</p><p>spreadsheet: <a href="+eventSpreadsheetUrl+">"+eventSpreadsheetUrl+"</a></p><p>All the function to manage this bot is under the main menu of the spreadsheet (Kill/restart...)</p>",
   });
  
}
//----------------------------
// TESTER
//----------------------------
function starterFunction3() {
  refreshUserProps(); // create properties
  Logger.log("properties created");
  controlProcess();
}
//----------------------------
// Main function
//----------------------------
function controlProcess() {
  var userProperties = PropertiesService.getUserProperties();
  var startTiggerDate = userProperties.getProperty('startTiggerDate');
  var endTiggerDate = userProperties.getProperty('endTiggerDate');
  var currentTime = Date.parse(new Date());
  var endTigger = Date.parse(endTiggerDate);
  
  Logger.log("controlProcess run at: "+ new Date()+"_______STARTED at :"+startTiggerDate+"_______END at :"+endTiggerDate);
  if (currentTime < endTigger) {
    publicationProcess(); //Post-traitment
    Logger.log("Post-traitment launch");
  } else {
    // End of the automatic spreadsheet process
    deleteTrigger();
    Logger.log("Tiggers deleted");
    // Delete all user properties in the current script.
    var userProperties = PropertiesService.getUserProperties();
    userProperties.deleteAllProperties();
    Logger.log("userProperties deleted");
    //todo add @notification to IT
  }
}

//----------------------------
// Post-traitment
//----------------------------
function publicationProcess() {

  clearFilter();
  Logger.log("Filter reseted at: "+ new Date());
  //STEP 1 : Sub process of SUBMITTED ROW (Go to Published state or validated state)
  setFilter(columnIndexForState, criterias[0])
  Logger.log("Filter activated 1: Submitted data"+ new Date());

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
  var lastColumn = sheet.getLastColumn();


  var submittedListOfIndex = getIndexesOfFilteredRows();
  var submittedLengthList = submittedListOfIndex.length;
  if (submittedLengthList > 0) {
    // LOGICAL TESTS
    for (var j = 0; j < submittedLengthList; j++) {
      var currentLigne = submittedListOfIndex[j] + 1;
      var row = proceedByRow(currentLigne, sheet, lastColumn);
      row = rowToObject(row[0]);
      Logger.log(row);
      var attendee = [];
      //1st TEST ID
      if (row.id.length <= 0) {
        Logger.log('ERROR_Event Id missing row: ' + currentLigne);
      }
      //2sd TEST emails
      else {
        // Built an array of emails and checking if it is an email else next line
        // Note:1st email will be the leader
        attendee = parseEmails(row.emails);
        if (attendee[1].length > 0) {
        var errorMsg = 'ERROR_row: ' + currentLigne + " _Emails errors"
          Logger.log(errorMsg);
        } 
          row.emails = attendee[0];
        
        if (row.type.length <= 0 || row.startTimeInput.length <= 0 || row.endTimeInput.length <= 0 || row.agendaId.length <= 0 || row.title.length <= 0 || row.topic <= 0) {
          Logger.log('ERROR_Data missing (Type/ Time/title/agenda Id): ' + currentLigne);
        }else{
        // CONDITIONNAL WORKFLOW
          // For public event from the form as onligne-event
          if (row.type == 'Event'&& row.topic == topic[1]) {
            Logger.log('SEND NOTIFICATION EMAIL to conciergerie/row: ' + currentLigne);
            row.workflowState = criterias[1]; // Update state to validation
            sheet.getRange((currentLigne), 1).setValue(row.workflowState);
            var htmlBodyForEventOwner = "<p>Bonjour "+row.ownerName+"</p>"+publicationMsg[1];
            htmlBodyForEventOwner += "<p>Salle: "+getRoomNameOfAnEvent(row.id)+" _  le "+row.startTimeInput.getDate()+"/"+(row.startTimeInput.getMonth()+1)+" de : "+row.startTimeInput.getHours()+"h"+row.startTimeInput.getMinutes()+" à : "+row.endTimeInput.getHours()+":"+row.endTimeInput.getMinutes()+" heure </p>";
            htmlBodyForEventOwner += signatureEmail;
             
            var htmlBodyForValidationOwner = "<p>Bonjour à l'équipe Conciergerie</p><p>Votre Event Bot vous informe que:</p>";
            htmlBodyForValidationOwner += "<p>Une demande d'évènement a été demandée par "+row.ownerName+"("+row.emails[0]+") pour le "+row.startTimeInput.getDate()+"/"+(row.startTimeInput.getMonth()+1)+" de : "+row.startTimeInput.getHours()+"h"+row.startTimeInput.getMinutes()+" à : "+row.endTimeInput.getHours()+"h"+row.endTimeInput.getMinutes()+"</p>";
            htmlBodyForValidationOwner += "<p>Salle: "+getRoomNameOfAnEvent(row.id)+"</p>";
            htmlBodyForValidationOwner += "<p>Le N° de réservation temporaire est le: "+row.id+". L'ensemble des informations relatives à cette évènement est disponible dans la spreadsheet<a href="+eventSpreadsheetUrl+"> Event 2018 ligne: "+currentLigne+"</a></p>";
            htmlBodyForValidationOwner += "<ul><li><a href='https://www.helloasso.com/associations/pangloss/evenements/les-services-co-working'>Paiement de vos services co-working via helloasso (privatisation, stockage, vid&eacute;oprojecteur...)</a></li>";
            htmlBodyForValidationOwner += "<li><a href='https://www.helloasso.com/associations/pangloss/evenements/les-services-fablab'>Paiement de vos services FabLab via helloasso (location machines)</a></li></ul><br/>";
            htmlBodyForValidationOwner += "<p>[Option] un lien Helloasso peut être ajouté pour assurer la réservation en ligne.</p><p>Pour toute information complémentaire ou anomalie, veuillez contracter l'équipe IT.</p><br><p>Un des Bots Pangloss ¯\_(ツ)_/¯ </p>";
; 
             
              // Send an e-mail to the user.
            MailApp.sendEmail({
              to: row.emails[0],
              subject: "PANGLOSS -"+row.type+" - Your event is waiting for conciergerie check - N°"+row.id,
              htmlBody: htmlBodyForEventOwner,
               });
             // Send an e-mail to conciergerie.
            MailApp.sendEmail({
              to: notificationEmailWorkflow,
              subject: "PANGLOSS - Event bot - A Conciergerie validation is requested for the event N°"+row.id+" submitted by: "+row.emails[0],
              htmlBody: htmlBodyForValidationOwner,
               });  
          }
          // For mass event generation from the spreadsheet
          if (row.type == 'Event'&& row.topic !== topic[1]) {
            row = generateGoogleEvent(row,currentLigne);
            row.workflowState = criterias[3]; // Update state
            sheet.getRange((currentLigne), 1).setValue(row.workflowState);
            sheet.getRange((currentLigne), 18).setValue(row.eventUrl);
            sheet.getRange((currentLigne), 19).setValue(row.eventId);
            sheet.getRange((currentLigne), 20).setValue(criterias[3]+": "+new Date());      
          }
          
          
     
          // waiting validation of an event with room attached.
          if ((row.type == 'Room'|| row.type == 'Machine')&& row.topic == topic[1]) {
            Logger.log('SEND NOTIFICATION EMAIL to conciergerie/row: ' + currentLigne);
            row.workflowState = criterias[1]; // Update state
            sheet.getRange((currentLigne), 1).setValue(row.workflowState);
          }
          // Auto booking room & machines
          if ((row.type == 'Room'|| row.type == 'Machine'|| row.type == 'Internal')&& row.topic !== topic[1]) {
            row = generateGoogleEvent(row,currentLigne);
            row.workflowState = criterias[3]; // Update state
            sheet.getRange((currentLigne), 1).setValue(row.workflowState);
            sheet.getRange((currentLigne), 18).setValue(row.eventUrl);
            sheet.getRange((currentLigne), 19).setValue(row.eventId);
            sheet.getRange((currentLigne), 20).setValue(criterias[3]+": "+new Date());
           
            var htmlBody = publicationMsg[0];
            htmlBody += "<p>"+row.type+" "+row.agendaName+" a bien été réservée pour le "+row.startTimeInput.getDate()+"/"+(row.startTimeInput.getMonth()+1)+" de : "+row.startTimeInput.getHours()+"h"+row.startTimeInput.getMinutes()+" à : "+row.endTimeInput.getHours()+":"+row.endTimeInput.getMinutes()+"</p>";
            htmlBody += signatureEmail;
            
            
              // Send an e-mail to the user.
            MailApp.sendEmail({
              to: row.emails[0],
              subject: "PANGLOSS -"+row.type+" "+row.agendaName+" is booked - N°"+row.eventId,
              htmlBody: htmlBody,
            });
            
            
            Logger.log('SEND NOTIFICATION EMAIL to USER/row: ' + currentLigne);
          }
          // Make sure the cell is updated right away in case the script is interrupted
			SpreadsheetApp.flush();
        }
      }
    }
   Logger.log ('number of processed row :'+ submittedLengthList+ 'at :'+ new Date());
  }
  //PHASE 2 in work

 clearFilter();
 setFilter(columnIndexForState, criterias[2])
 Logger.log("Filter activated 2: Submitted data"+ new Date());

 lastColumn = sheet.getLastColumn();

 submittedListOfIndex = getIndexesOfFilteredRows();
 submittedLengthList = submittedListOfIndex.length;
   if (submittedLengthList > 0) {
    // LOGICAL TESTS
    for (var j = 0; j < submittedLengthList; j++) {
      var currentLigne = submittedListOfIndex[j] + 1;
      var row = proceedByRow(currentLigne, sheet, lastColumn);
      row = rowToObject(row[0]);
      Logger.log(row);
      var attendee = [];
      //1st TEST ID
      if (row.id.length <= 0) {
        Logger.log('ERROR_Event Id missing row: ' + currentLigne);
      }
      //2sd TEST emails
      else {
        // Built an array of emails and checking if it is an email else next line
        // Note:1st email will be the leader
        attendee = parseEmails(row.emails);
        if (attendee[1].length > 0) {
          var errorMsg2 = 'ERROR_row: ' + currentLigne + " _Emails errors"
          Logger.log(errorMsg2);
        } 
          row.emails = attendee[0];
        
        if (row.type.length <= 0 || row.startTimeInput.length <= 0 || row.endTimeInput.length <= 0 || row.agendaId.length <= 0 || row.title.length <= 0 || row.topic <= 0) {
          Logger.log('ERROR_Data missing (Type/ Time/title/agenda Id): ' + currentLigne);
        }else{ 
          
                    // Auto booking room & machines
          if (row.topic == topic[1]) {
            row = generateGoogleEvent(row,currentLigne);
            row.workflowState = criterias[3]; // Update state
            sheet.getRange((currentLigne), 1).setValue(row.workflowState);
            sheet.getRange((currentLigne), 18).setValue(row.eventUrl);
            sheet.getRange((currentLigne), 19).setValue(row.eventId);
            sheet.getRange((currentLigne), 20).setValue(criterias[3]+": "+new Date());
            
            var htmlBodyForEventOwner2 = publicationMsg[0];
            htmlBodyForEventOwner2 += "<p>"+row.type+" "+row.agendaName+" a bien été réservée pour le "+row.startTimeInput.getDate()+"/"+(row.startTimeInput.getMonth()+1)+" de : "+row.startTimeInput.getHours()+"h"+row.startTimeInput.getMinutes()+" à : "+row.endTimeInput.getHours()+"h"+row.endTimeInput.getMinutes()+"</p>";
            htmlBodyForEventOwner2 += signatureEmail;
            
            var htmlBodyForValidationOwner2 = "<p>Bonjour à l'équipe Conciergerie</p><p>Votre Event Bot vous informe que:</p>";
            htmlBodyForValidationOwner2 += "<p>Vous avez validé l'évènement demandée par "+row.ownerName+"("+row.emails[0]+") pour le "+row.startTimeInput.getDate()+"/"+(row.startTimeInput.getMonth()+1)+" de : "+row.startTimeInput.getHours()+"h"+row.startTimeInput.getMinutes()+" à : "+row.endTimeInput.getHours()+"h"+row.endTimeInput.getMinutes()+"</p>";
            htmlBodyForValidationOwner2 += "<p>Salle: "+getRoomNameOfAnEvent(row.id)+"</p>";
            htmlBodyForValidationOwner2 += "<p>Le N° de réservation est le: "+row.id+". L'ensemble des informations relatives à cette évènement est disponible dans la spreadsheet<a href="+eventSpreadsheetUrl+"> Event 2018 ligne: "+currentLigne+"</a></p>";
            htmlBodyForValidationOwner2 += "<ul><li><a href='https://www.helloasso.com/associations/pangloss/evenements/les-services-co-working'>Paiement de vos services co-working via helloasso (privatisation, stockage, vid&eacute;oprojecteur...)</a></li>";
            htmlBodyForValidationOwner2 += "<li><a href='https://www.helloasso.com/associations/pangloss/evenements/les-services-fablab'>Paiement de vos services FabLab via helloasso (location machines)</a></li></ul><br/>";
            htmlBodyForValidationOwner2 += "<p>[Option] un lien Helloasso peut être ajouté pour assurer la réservation en ligne.</p><p>Pour toute information complémentaire ou anomalie, veuillez contracter l'équipe IT.</p><br><p>Un des Bots Pangloss ¯\_(ツ)_/¯ </p>"; 
            
            
              // Send an e-mail to the user.
            MailApp.sendEmail({
              to: row.emails[0],
              subject: "PANGLOSS -"+row.type+" "+row.agendaName+" is booked - N°"+row.eventId,
              htmlBody: htmlBodyForEventOwner2,
            });
            // Send an e-mail to conciergerie.
            MailApp.sendEmail({
              to: notificationEmailWorkflow,
              subject: "PANGLOSS - Event bot - A Conciergerie validation DONE for the event N°"+row.id+" submitted by: "+row.emails[0],
              htmlBody: htmlBodyForValidationOwner2,
               });
            
            Logger.log('SEND NOTIFICATION EMAIL to USER/row: ' + currentLigne);
          }else{
              row = generateGoogleEvent(row,currentLigne);
            row.workflowState = criterias[3]; // Update state
            sheet.getRange((currentLigne), 1).setValue(row.workflowState);
            sheet.getRange((currentLigne), 18).setValue(row.eventUrl);
            sheet.getRange((currentLigne), 19).setValue(row.eventId);
            sheet.getRange((currentLigne), 20).setValue(criterias[3]+": "+new Date());
            
          }
          // Make sure the cell is updated right away in case the script is interrupted
			SpreadsheetApp.flush();
          
          
        }}}}
  clearFilter();
}
//----------------------------
// Row process (analysis line by line)
//----------------------------
function proceedByRow(index, sheet, lastColumn) {
  var range = sheet.getRange(index, 1, 1, lastColumn);
  var values = range.getValues();
  return values;
}
//----------------------------
// Event generator (google agenda)
//----------------------------
function generateGoogleEvent(eventObject,currentLigne) {

if (eventObject.eventUrl.length <= 0) {// Prevents sending duplicates
                  
			//Manage Event description - Fr-Eng
			var divider = "<br>___________________________________________________<br>";
          
			if (eventObject.urlHelloAsso != "" || eventObject.urlHelloAsso.length >0) {
				var linkFr = '<a href="' + eventObject.urlHelloAsso + '">Inscription via le lien : Pangloss Helloasso</a>';
				var linkEn = '<a href="' + eventObject.urlHelloAsso + '">Subscribe using the link : Pangloss Helloasso</a>';
				var eventDescription = eventObject.eventDescriptionFr + linkFr + divider + eventObject.eventDescriptionEn + linkEn;
			} else {
				var eventDescription = eventObject.eventDescriptionFr + divider + eventObject.eventDescriptionEn;
			}
          
			// Gets the calendar by ID.
			var calendar = CalendarApp.getCalendarById(eventObject.agendaId);
            //Logger.log('The calendar is named "%s".', calendar.getName());
          
			//Check if calendar exist
			if (calendar == null) {
				Logger.log("ligne :" + currentLigne + " _calendar id errors");
			}
            
			//Creates an event
			var event = calendar.createEvent(eventObject.title, eventObject.startTimeInput, eventObject.endTimeInput, {
				location : eventObject.adress,
				guests : eventObject.emails[0],
				description : eventDescription
			});
          
			//Event custom metadata
			event.setGuestsCanModify(false);
			event.setGuestsCanSeeGuests(true);
			var creationDate = new Date();
			event.setTag('creationDate', creationDate);
			event.setTag('type', eventObject.type);
			if (eventObject.topic != "") {
				event.setTag('topic', eventObject.topic);
			}
			if (eventObject.urlHelloAsso != "") {
				event.setTag('helloassoUrl', eventObject.urlHelloAsso);
			}
			var eventId = event.getId();
			var splitEventId = eventId.split('@');
			var eventUrl = "https://www.google.com/calendar/event?eid=" + Utilities.base64Encode(splitEventId[0] + " " + eventObject.agendaId);
			//Logger.log('Event url: ' + eventUrl);
            eventObject.eventUrl = eventUrl;
            eventObject.eventId = eventId;
			// Notify guest only for event and open doors
			if ( type = "Event") {
				event.addEmailReminder(240);
				//reminder 4h
				event.addEmailReminder(1440);
				//reminder 1J
				event.addEmailReminder(8640);
				//reminder 6J
			}
            if ( type = "CA") {
				event.addEmailReminder(240);
				//reminder 4h
				event.addEmailReminder(1440);
				//reminder 1J
                event.addEmailReminder(2880);
				//reminder 2J
				event.addEmailReminder(8640);
				//reminder 6J
			}
			if ( type = "Open-doors") {
				event.addEmailReminder(240);
				//reminder 4h
				event.addEmailReminder(1440);
				//reminder 1J
				event.addEmailReminder(2880);
				//reminder 2J
				event.addEmailReminder(8640);
				//reminder 6J
			}
            
            //Add guests
            if (eventObject.emails.length >= 2) {
				for (var e = 1; e < eventObject.emails.length; ++e) {
					event.addGuest(eventObject.emails[e]);
				}
			}
          
		}else{
        Logger.log('Event URL already exist: ' + currentLigne);
        }
return eventObject;
}
//----------------------------
// Convert Row data into Object
//----------------------------
function rowToObject(row) {
  var rowObject = {
    workflowState: row[0], //[A]
    urlHelloAsso: row[1], //[B]
    type: row[2], //[C] Mandatory*
    topic: row[3], //[D]
    startTimeInput: row[7], //[H] Mandatory*
    endTimeInput: row[8], //[I] Mandatory*
    ownerName: row[9], //[J] Mandatory*
    emails: row[10], //[K] Mandatory*
    eventDescriptionFr: row[11], //use html
    eventDescriptionEn: row[12], //use html
    agendaName: row[13], //[N] 
    agendaId: row[14], //[O] Mandatory*
    adress: row[15],
    title: row[16], //Mandatory*
    eventUrl: row[17], //[R]
    id: row[18], //[S] Mandatory*
    eventCreationDate: row[19], //[T]
  };
  return rowObject;
}
//---------------------------------------------
// Import a string with coma separator and generate an array of emails and an array of errors
//---------------------------------------------
function parseEmails(input) {
  var emails = input;
  var arrayList = [];
  var errors = [];
  var isEmail = /^[a-zA-Z0-9._-]+@[a-z0-9._-]{2,}\.[a-z]{2,4}$/;
  emails = emails.replace(/;/gi, ",");
  var emailArray = emails.split(',');
  var j = emailArray.length;
  for (var i = 0; i < j; ++i) {
    isEmail.lastIndex = 0;
    // reset state 
    var email = emailArray[i].trim();
    var testEmail = isEmail.test(email);
    if (testEmail == false) {
      errors.push(email);
    } else {
      arrayList.push(email);
    }
  }
  return [arrayList, errors];
}
//---------------------------------------------
// Get EventId of an event line to extract the room name linked with 
//---------------------------------------------
function getRoomNameOfAnEvent(eventId) {
    var roomName = eventId.split("@");
    return roomName[1];
}
//---------------------------------------------
// Get filtered data 
//---------------------------------------------
function getIndexesOfFilteredRows() {
  var hiddenRows = [];
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ssId = ss.getId();
  var sheetId = ss.getActiveSheet().getSheetId();

  // limit what's returned from the API
  var fields = "sheets(data(rowMetadata(hiddenByFilter)),properties/sheetId)";
  var sheets = Sheets.Spreadsheets.get(ssId, {
    fields: fields
  }).sheets;

  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].properties.sheetId == sheetId) {
      var data = sheets[i].data;
      var rows = data[0].rowMetadata;
      for (var j = 0; j < rows.length; j++) {
        if (rows[j].hiddenByFilter) hiddenRows.push(j);
      }
    }
  }
  Logger.log("Filter applied on row: " + hiddenRows);
  return hiddenRows;
}
//---------------------------------------------
// Test ID is event 
//---------------------------------------------
function testEventId() {
  var str = "2018-06-10T17:12:51.961Z"; 
  var n = str.search("eventRoom");
  var isEventRow = "";
  if( n > 0 ){
  isEventRow = true;
  }else{
  isEventRow = false;
  }
  return isEventRow;
}
// -----------------------------------------------------------------------------
// reset date counter to now and end tigger date to now+1year in properties
// -----------------------------------------------------------------------------
function refreshUserProps() {
  var userProperties = PropertiesService.getUserProperties();
  userProperties.deleteAllProperties();
  var startTiggerDate = new Date();
  var year = startTiggerDate.getFullYear();
  var month = startTiggerDate.getMonth();
  var day = startTiggerDate.getDate();
  var hour = startTiggerDate.getHours();
  var minute = startTiggerDate.getMinutes();
  var endTiggerDate = new Date(year+1, month, day, hour, minute)
  userProperties.setProperty('startTiggerDate', startTiggerDate);
  userProperties.setProperty('endTiggerDate', endTiggerDate);
  userProperties.setProperty('runInit', true);
  Logger.log(startTiggerDate + " =>|=> " + endTiggerDate);
  return [startTiggerDate, endTiggerDate];
}
// -----------------------------------------------------------------------------
// SUPPORT FUNCTION: dispay all userproperties in properties
// -----------------------------------------------------------------------------
function callUserProps() {
  // Get multiple script properties in one call, then log them all.
  var userProperties = PropertiesService.getUserProperties();
  var data = userProperties.getProperties();
  for (var key in data) {
    Logger.log('Key: %s, Value: %s', key, data[key]);
  }
}

// -----------------------------------------------------------------------------
// create trigger to run publicationProcess every hour
// -----------------------------------------------------------------------------
function createTrigger() {

  // Trigger every 1 minute
  ScriptApp.newTrigger('controlProcess')
    .timeBased()
    .everyMinutes(1)
    .create();
}
// -----------------------------------------------------------------------------
// function to delete triggers
// -----------------------------------------------------------------------------
function deleteTrigger() {

  // Loop over all triggers and delete them
  var allTriggers = ScriptApp.getProjectTriggers();

  for (var i = 0; i < allTriggers.length; i++) {
    ScriptApp.deleteTrigger(allTriggers[i]);
  }
    MailApp.sendEmail({
   to: notificationEmailIT,
   subject: "PANGLOSS - Event Bot & his Trigger - Google Script Apps delete all triggers under this account",
   htmlBody: "<p>TriggerS KILLED</p><p> using a GAS under spreadsheet: <a href="+eventSpreadsheetUrl+">"+eventSpreadsheetUrl+"</a></p><p>All the function to manage this bot is under the main menu of the spreadsheet (Kill/restart...)</p>",
   });
   //Logger.log ('It notification send');
}
//---------------------------------------------
// Clear / remove all filters 
//---------------------------------------------
function clearFilter() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ssId = ss.getId();
  var sheetId = ss.getActiveSheet().getSheetId();
  var requests = [{
    "clearBasicFilter": {
      "sheetId": sheetId
    }
  }];
  Sheets.Spreadsheets.batchUpdate({
    'requests': requests
  }, ssId);
}
//---------------------------------------------
// Create a new filter
//---------------------------------------------
function setFilter(columnIndex, criteria) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var filterSettings = {};

  // The range of data on which you want to apply the filter.
  // optional arguments: startRowIndex, startColumnIndex, endRowIndex, endColumnIndex
  filterSettings.range = {
    sheetId: ss.getActiveSheet().getSheetId()
  };

  // Criteria for showing/hiding rows in a filter
  // https://developers.google.com/sheets/api/reference/rest/v4/FilterCriteria
  filterSettings.criteria = {};

  filterSettings['criteria'][columnIndex] = {
    'hiddenValues': [criteria]
  };

  var request = {
    "setBasicFilter": {
      "filter": filterSettings
    }
  };
  Sheets.Spreadsheets.batchUpdate({
    'requests': [request]
  }, ss.getId());
}
//---------------------------------------------
// Delete a list of event using event id.
//---------------------------------------------
function deleteEvents() {
	var cancel = "delete";
	var deletedMsg = "deleted";

	var ss = SpreadsheetApp.getActiveSpreadsheet();
	//var sheet = ss.getSheets()[2];
	var sheet = ss.getSheetByName('load event');
	//load event
	var lastRow = sheet.getLastRow();
	var startRow = 2;
	// First row of data to process
	var range = sheet.getRange(startRow, 1, lastRow - 1, 21);
	var values = range.getValues();

	SpreadsheetApp.getUi()// Or DocumentApp or FormApp.
	.alert("Warning, this function will delete all events with a delete string in the correct cell");

	for (var i = 0; i < values.length; ++i) {

		var row = values[i];
		var currentLigne = i + 2;
		var toDelete = row[21];
		var eventId = row[19];
		var calendarId = row[15];

		if (toDelete.length > 0 && toDelete != cancel) {
			//Logger.log("ligne :" + currentLigne + " _errors");
			SpreadsheetApp.getUi()// Or DocumentApp or FormApp.
			.alert("ligne :" + currentLigne + " _errors");
			continue;
		} else {
			if (toDelete == "") {
				continue;
			} else {
				//Delete event
				var calendar = CalendarApp.getCalendarById(calendarId);
				var event = calendar.getEventSeriesById(eventId);
                //Check if calendar exist
			      if (event == null) {
                    //Logger.log("ligne :" + currentLigne + " _event id or calendar id errors");
                    SpreadsheetApp.getUi()// Or DocumentApp or FormApp.
                    .alert("ligne :" + currentLigne + " _event id or calendar id errors");
                    continue;
                  }
				event.deleteEventSeries();
				//Logger.log("ligne :"+currentLigne+" _deleted");
              
                // Set metadata to the spreadsheet
				sheet.getRange(startRow + i, 21).setValue("");
				sheet.getRange(startRow + i, 20).setValue(criterias[3]+": "+new Date());
				sheet.getRange(startRow + i, 19).setValue("");
                sheet.getRange(startRow + i, 18).setValue("");
				sheet.getRange(startRow + i, 1).setValue(deletedMsg);
				// Make sure the cell is updated right away in case the script is interrupted
				SpreadsheetApp.flush();
			}
		}
	}
    SpreadsheetApp.getUi()
	.alert('Script end : Google Agenda event(s) deleted');
}
