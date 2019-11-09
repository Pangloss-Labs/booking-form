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



// Check event tags
function getEventTags() {
	var calendarId = 'ln1h82ecd06sr4dakdtpl60k2o@group.calendar.google.com';
	var eventId = 'nj4ldgnu0iiin9bota2m8vk5ak@google.com';
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
	var range = sheet.getRange(startRow, 1, lastRow - 1, 21);
	var values = range.getValues();

	SpreadsheetApp.getUi()// Or DocumentApp or FormApp.
	.alert("Warning, this function will delete all events with a delete string in the correct cell");

	for (var i = 0; i < values.length; ++i) {

		var row = values[i];
		var currentLigne = i + 2;
		var toDelete = row[20];
		var eventId = row[18];
		var calendarId = row[14];

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
