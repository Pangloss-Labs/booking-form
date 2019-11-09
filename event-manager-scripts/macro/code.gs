/* Create spreadsheet menu
function onOpen() {
	var ui = SpreadsheetApp.getUi();
	// Or DocumentApp or FormApp.
	ui.createMenu('Pangloss Event Manager').addItem('Run event creation macro', 'loadEventData')
    .addSeparator()
    .addSubMenu(ui.createMenu('Sub function').addItem('Delete events', 'deleteEvents').addItem('get event id', 'getEventsId')).addToUi();    
}
*/
// Main function
function loadEventData() {

	//Access to data
	var ss = SpreadsheetApp.getActiveSpreadsheet();
	//var sheet = ss.getSheets()[2];
	var sheet = ss.getSheetByName('load event');//load event
	var lastRow = sheet.getLastRow();
	var startRow = 2;// First row of data to process
	var range = sheet.getRange(startRow, 1, lastRow - 1, 16);
	var values = range.getValues();

	for (var i = 0; i < values.length; ++i) {
		var currentLigne = i + 2;
		var row = values[i];
		var url = row[0];//check if url (TODO)
		var type = row[1];//mandatory, used into event tags
		var topic = row[2];// used into event tags
		var date = String(row[3]);//mandatory
		var startTimeInput = row[4];//mandatory
		var endTimeInput = row[5];//mandatory
		var startTime = row[6];//mandatory - from spreadsheet function
		var endTime = row[7];//mandatory - from spreadsheet function
		var emails = row[8];//mandatory
		var eventDescriptionFr = row[9];//use html
		var eventDescriptionEn = row[10];//use html
		var agendaId = row[12];//mandatory
		var adress = row[13];
		var title = row[14];//mandatory
		var eventDescription = "";
		var outputEventUrl = row[15];
		// row[16] = event id
		// row[17] = event creation date
        
        var attendee = [];
      
        if (outputEventUrl.length <= 0) { //Added on 24/01/2018 by AU. If there is a "X" in the event output url, the script is ignored (the line is not treated)
          //Checking mandatory data
          if (type.length <= 0 || date.length <= 0 || startTimeInput.length <= 0 || endTimeInput.length <= 0 || startTime.length <= 0 || emails.length <= 0 || endTime.length <= 0 || agendaId.length <= 0 || title.length <= 0) {
              //Logger.log("ligne number: "+currentLigne+" _Mandatory data empty");
              SpreadsheetApp.getUi()// Or DocumentApp or FormApp.
              .alert("ligne number: " + currentLigne + " _Mandatory data empty");
              continue;
          }
        

          // Built an array of emails and checking if it is an email else next line
          // Note:1st email will be the leader
          attendee = parseEmails(emails);
          if (attendee[1].length > 0) {
              //Logger.log("ligne :"+currentLigne+" _Emails errors");
              SpreadsheetApp.getUi()// Or DocumentApp or FormApp.
              .alert("ligne :" + currentLigne + " _Emails errors");
              continue;
          } else {
              emails = attendee[0];
          }
  
          //var testdate = dateCheck([date,startTime,endTime],i+2);
          //if( testdate === false){
          // return;
          //}
        }
		if (outputEventUrl.length <= 0) {// Prevents sending duplicates
           
          //(TODO) check if description
          
			//Manage Event description - Fr-Eng
			var divider = "<br>___________________________________________________<br>";
          
			if (url != "") {
				var linkFr = '<a href="' + url + '">Inscription via le lien : Pangloss Helloasso</a>';
				var linkEn = '<a href="' + url + '">Subscribe using the link : Pangloss Helloasso</a>';
				var eventDescription = eventDescriptionFr + linkFr + divider + eventDescriptionEn + linkEn;
			} else {
				var eventDescription = eventDescriptionFr + divider + eventDescriptionEn;
			}
          
			// Gets the calendar by ID.
			var calendar = CalendarApp.getCalendarById(agendaId);
            //Logger.log('The calendar is named "%s".', calendar.getName());
          
			//Check if calendar exist
			if (calendar == null) {
				Logger.log("ligne :" + currentLigne + " _calendar id errors");
				SpreadsheetApp.getUi()// Or DocumentApp or FormApp.
				.alert("ligne :" + currentLigne + " _calendar id errors");
				continue;
			}
            
			//Creates an event
			var event = calendar.createEvent(title, startTime, endTime, {
				location : adress,
				guests : emails[0],
				description : eventDescription
			});
          
			//Event custom metadata
			event.setGuestsCanModify(false);
			event.setGuestsCanSeeGuests(true);
			var creationDate = new Date();
			event.setTag('creationDate', creationDate);
			event.setTag('type', type);
			if (topic != "") {
				event.setTag('topic', topic);
			}
			if (url != "") {
				event.setTag('helloassoUrl', url);
			}
			var eventId = event.getId();
			var splitEventId = eventId.split('@');
			var eventUrl = "https://www.google.com/calendar/event?eid=" + Utilities.base64Encode(splitEventId[0] + " " + agendaId);
			//Logger.log('Event url: ' + eventUrl);
            
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
            if (emails.length >= 2) {
				for (var e = 1; e < emails.length; ++e) {
					event.addGuest(emails[e]);
				}
			}
            
            //addGuests(emails[0], event);
            
			//Save event metadata into the spreadsheet
			sheet.getRange(startRow + i, 16).setValue(eventUrl);
			sheet.getRange(startRow + i, 17).setValue(eventId);
			sheet.getRange(startRow + i, 18).setValue(creationDate);
			// Make sure the cell is updated right away in case the script is interrupted
			SpreadsheetApp.flush();
		}
	}
	SpreadsheetApp.getUi()
	.alert('Script end : Google Agenda event(s) created only if an agenda url is available');
}

//in work
function dateCheck(dates, ligneNumber) {

	var pattern = /GMT+/g;
	for (var i = 0; i < dates.length; ++i) {
		var testDate = undefined;
		var tt = dates[i].toString().length;
		pattern.lastIndex = 0;
		// reset state g
		if (dates[i].length > 0) {
			var test = dates[i].toString();
			testDate = pattern.test(test);
			Logger.log(testDate + "--" + i);
		} else {
			var DateFormatMsg = "ligne :" + ligneNumber + "Date, startTime or endTime must be completed";
		}

	}
	var timeslotConsistency = dates[2] - dates[1];
	//must >0
	if (timeslotConsistency <= 0) {
		var DateConsistencyMsg = "ligne :" + ligneNumber + " The endtime cannot be placed before the start time";
	}
	if (DateFormatMsg !== undefined || DateConsistencyMsg !== undefined) {
		SpreadsheetApp.getUi()// Or DocumentApp or FormApp.
		.alert(DateFormatMsg + ' - ' + DateConsistencyMsg);
		return false;
	}
	return true;
}

// import a string with coma separator and generate an array of emails and an array of errors
function parseEmails(input) {
	var emails = input;
	var arrayList = [];
	var errors = [];
	var isEmail = /^[a-zA-Z0-9._-]+@[a-z0-9._-]{2,}\.[a-z]{2,4}$/;
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

// Check event tags
function getEventTags() {
	var calendarId = 'sample@group.calendar.google.com';
	var eventId = 'sample@google.com';
	var calendar = CalendarApp.getCalendarById(calendarId);
	var event = calendar.getEventSeriesById(eventId);
	var eventTags = event.getAllTagKeys();
	Logger.log(eventTags);
	for (var i = 0; i < eventTags.length; ++i) {
		var tag = event.getTag(eventTags[i]);
		Logger.log(eventTags[i] + ": " + tag);
	}
	return;
}

// Delete a list of event using event id.
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
	var range = sheet.getRange(startRow, 1, lastRow - 1, 19);
	var values = range.getValues();

	SpreadsheetApp.getUi()// Or DocumentApp or FormApp.
	.alert("Warning, this function will delete all events with a delete string in the correct cell");

	for (var i = 0; i < values.length; ++i) {

		var row = values[i];
		var currentLigne = i + 2;
		var toDelete = row[18];
		var eventId = row[16];
		var calendarId = row[12];

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
				sheet.getRange(startRow + i, 19).setValue(deletedMsg);
				sheet.getRange(startRow + i, 18).setValue("");
				sheet.getRange(startRow + i, 17).setValue("");
				sheet.getRange(startRow + i, 16).setValue(deletedMsg);
				// Make sure the cell is updated right away in case the script is interrupted
				SpreadsheetApp.flush();
			}
		}
	}
    SpreadsheetApp.getUi()
	.alert('Script end : Google Agenda event(s) deleted');
}

// Get event id using Event data.
function getEventsId() {

	var ss = SpreadsheetApp.getActiveSpreadsheet();
	//var sheet = ss.getSheets()[2];
	var sheet = ss.getSheetByName('load event');
	//load event
	var lastRow = sheet.getLastRow();
	var startRow = 2;
	// First row of data to process
	var range = sheet.getRange(startRow, 1, lastRow - 1, 19);
	var values = range.getValues();

	for (var i = 0; i < values.length; ++i) {
      
		var row = values[i];
		var currentLigne = i + 2;
        var startTime = row[6];//mandatory - from spreadsheet function
		var endTime = row[7];//mandatory - from spreadsheet function
		var calendarId = row[12];
        var eventUrl = row[15];
        var eventId = row[16];
        var type = row[1];//mandatory, used into event tags
		var topic = row[2];// used into event tags
      
      if (eventUrl.length <= 0){
       	//Logger.log("ligne number: "+currentLigne+" _Mandatory data empty");
		continue; 
      }else{ 
        if (eventId.length > 0 ){
        //Logger.log("ligne number: "+currentLigne+" _event id exist");
		continue; 
      }else{
        
        //Get calendar
		var calendar = CalendarApp.getCalendarById(calendarId);
      
        //Check if calendar exist
		if (calendar == null) {
            //Logger.log("ligne :" + currentLigne + " _calendar id errors");
            SpreadsheetApp.getUi()// Or DocumentApp or FormApp.
             .alert("ligne :" + currentLigne + " _calendar id errors");
            continue;
        }
        
        //Get events
		var events = calendar.getEvents(startTime, endTime);
        if (events.length >1) {
            //Logger.log("ligne :" + currentLigne + " _multiple events at the same slot: errors");
            SpreadsheetApp.getUi()// Or DocumentApp or FormApp.
            .alert("ligne :" + currentLigne + " _multiple events at the same slot: errors");
            continue;          
        }
        var event = events[0]
        eventId = event.getId(); 
        //Logger.log(eventId);
        
        //Event custom metadata
			event.setGuestsCanModify(false);
			event.setGuestsCanSeeGuests(true);
			var creationDate = new Date();
            var eventTags = event.getAllTagKeys();
        if(eventTags.length = 0){
			event.setTag('creationDate', creationDate);
			event.setTag('type', type);
			if (topic != "") {
				event.setTag('topic', topic);
			}
			if (url != "") {
				event.setTag('helloassoUrl', url);
			}
        }
   
        // Set metadata to the spreadsheet
		sheet.getRange(startRow + i, 17).setValue(eventId);
		sheet.getRange(startRow + i, 18).setValue(creationDate);
		// Make sure the cell is updated right away in case the script is interrupted
		SpreadsheetApp.flush();
      }
      }
    }
    SpreadsheetApp.getUi()
	.alert('Script end : Google Agenda event(s) updated');
}

function addGuests(emails, event){
var emailList = emails;
  for (var i = 0; i < emailList.length; ++i) {
        Logger.log("Sent to: " + emailList[i]);
		var newGuestAdded = event.addGuest(emailList[i]);
		Logger.log("Guest added: " + emailList[i]);
	}
  return;
}
