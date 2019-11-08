function doGet() {

  var template = HtmlService.createTemplateFromFile('Form.html');
  var moment = Moment.load();

  // load parameters
  template.pubUrl = ScriptApp.getService().getUrl();
  template.appsTitle = appsTitle;
  template.bookingTypes = bookingTypes;
  template.dateMax = dateMax;
  // generate html template
  var html = template.evaluate();
  var output = HtmlService.createHtmlOutput(html);
  output.setTitle(appsTitle);
  output.setFaviconUrl(appsIcon);
  return output; 
}
function processForm(formObject) {
  //load moment library to deal with dates
  Logger.log(formObject);
    var moment = Moment.load();
  //This is what we get : {eventDescFr=, =[, FileUpload, more info from YL], 
  //date=, lastName=last, roomList1=ZEBRA, roomList2=, title=, 
  //extraAttendees=attendee1@pouet.com;attendee2@pouet.com, firstName=first, 
  //eventDescription=blablabla event descirptiotnt, startTime=13:00, endTime=15:00, 
  //extraAttendees2=, eventDescEng=, machineList=, email=email@email.com}
  //Logger.log(formObject);
  //form validation (could be done on the browser side but I'm too lazy)
  if(formObject.endTime == '' || formObject.endTime == undefined){
    result = "Please fill in the end time";
    return result;
  }
  if(formObject.startTime == '' || formObject.startTime == undefined){
    result = "Please fill in the start time";
    return result;
  }
  if(formObject.date == undefined || formObject.date == ''){
    result = "Please fill in the date";
    return result;
  }
  var result = "youpi";
  var type;
  var bookedItem;
  var attendees ="";
  var descFr ="";
  var descEng ="";
  var eventId = moment().toISOString();
  var creationDate = moment().format("DD[/]MM[/]YYYY h:mm:ss");
  //check which extra Attendees field was filled
  if(formObject.extraAttendees.length >1){
    //check each item to see if it's a valid email
    var testInput = parseEmails(formObject.extraAttendees);
    if(testInput[1].length>0){
       result = testInput[1]+" is an invalid email, sorry";
       return result;
       }else{
       attendees = formObject.email+","+formObject.extraAttendees;
       }
  } else if(formObject.extraAttendees2.length >1){
      //check each item to see if it's a valid email
    var testInput = parseEmails(formObject.extraAttendees2);
    if(testInput[1].length>0){
      result = testInput[1]+" is an invalid email, sorry";
       return result;
       }else{
       attendees = formObject.email+","+formObject.extraAttendees2;
       }
  } else {
    attendees = formObject.email;
  }
  Logger.log("attendees="+attendees);
  //check which description we're going to use
  if(formObject.eventDescEng.length >1){
    descEng = formObject.eventDescEng;
  } else if(formObject.eventDescEng2.length >1){
    descEng = formObject.eventDescEng2;
  }
  //Logger.log(descEng);
   if(formObject.eventDescFr.length >1){
    descFr = formObject.eventDescFr;
  } else if(formObject.eventDescFr2.length >1){
    descFr = formObject.eventDescFr2;
  }
   //Logger.log(descFr);
  //figure out what we're trying to book
  if(formObject.roomList1.length > 1){
    type = "Room";
    bookedItem = findObjectByKey(roomList,"name",formObject.roomList1.toLowerCase());
  }else if(formObject.roomList2.length > 1){
    type = "Event";
    bookedItem = findObjectByKey(roomList,"name",formObject.roomList2.toLowerCase());
  } else if(formObject.machineList.length > 1){
    type = "Machine";
    bookedItem = findObjectByKey(machineList,"name",formObject.machineList.toLowerCase());
  } else {
     //Make sure I have something otherwise get back to the user and insult him
    result = "Please select a machine or a room";
    return result;
  }
 
  Logger.log("Booked item= "+bookedItem.name);
  //Checking if room is already booked
  //find the room in the roomList array
  //Get the item calendar by its ID
  var calendar = CalendarApp.getOwnedCalendarById(bookedItem.id);
  
  //get the events for that room within the selected time interval
  var eventStart = new Date(formObject.date+'T'+formObject.startTime+':00');
  var eventEnd = new Date(formObject.date+'T'+formObject.endTime+':00');
  var existingEvents = calendar.getEvents(eventStart,eventEnd);
  //Logger.log(existingEvents);
  //if there is already one, send message
  if(existingEvents.length>0){
    result = bookedItem.name+" is already booked at that time, sorry";
    return result;
  }
  
  //Figuring out if there is already a conflicting event in google sheets that hasn't been processed yet (this is a pain in the ass)
  // get the values from the last 10 rows of the excel sheet
  var row = sheet.getLastRow()-10 > 0 ? sheet.getLastRow()-10: 1;
  var values = sheet.getRange(row,1,11,13).getDisplayValues();
  //Logger.log("unfiltered values = "+values);
  //Filter out everything that doesn't overlap
  var filtered = values.filter(function (row) {
    //Logger.log("cellule = "+row[1]);
    if(row[1]==''){
      return false;
    }
    if(row[12] == bookedItem.id){
      if(moment(eventStart).format("DD[/]MM[/]YYYY") == row[3]){
        var existingStart = row[4];
        var existingEnd = row[5];
        var eventStartTime = moment(eventStart).format("H:mm");
        var eventEndTime = moment(eventEnd).format("H:mm");
        if(existingEnd <= eventStartTime){
          Logger.log(existingEnd +"<="+ eventStartTime);
          return false;
        }
        if(existingStart >= eventEndTime){
          Logger.log(existingStart +">="+ eventEndTime);
          return false;
        }
      }else{
        return false;
      }
    }else{
      return false;
    }
    return true;
  });
  //Logger.log("filtered values = "+filtered);
  //If there are any items overlapping, insult the user
   if(filtered.length>0){
     Logger.log(filtered);
    result = bookedItem.name+" is already booked at that time, sorry, an overlapping event needs to be approved and doesn't appear on the calendar yet";
    return result;
  }
  //Logger.log(filtered.length);
  
  //if it's an event, set a different topic so the script can recognize it in google sheets
  var topic;
  if(type == "Event"){
    topic = "Online-Event";
    type = "Room";
  }else{
    topic = "Online";
  }
  
  //create an object with all the info we'll send to google sheets ( technically useless but it makes things a bit easier to read)
  
  var eventData = {type:type
              , eventDate: moment(formObject.date).format("DD[/]MM[/]YYYY")
                   , topic: topic
              , startTime: formObject.startTime
              , endTime: formObject.endTime
              , attendees: attendees
              , idAgenda: bookedItem.id
              , eventTitle: formObject.title
              ,descriptionFr: descFr
                   ,descriptionEng: descEng
                   ,agendaName : bookedItem.name
              ,firstName : formObject.firstName.toUpperCase()
              ,lastName : formObject.lastName.toUpperCase()
                   , address : "PanglossLabs - FabLab 12 bis rue de Gex, 01210 Ferney Voltaire"
                   ,eventId : eventId
                   ,creationDate : creationDate
                   ,information : formObject.moreInfo
             }
       
  //Logger.log(eventData);

//append the data to it
Logger.log(eventData.eventDate+" "+eventData.startTime);
 sheet.appendRow(["submitted","",eventData.type,eventData.topic,eventData.eventDate,eventData.startTime,eventData.endTime,eventData.eventDate+" "+eventData.startTime,eventData.eventDate+" "+eventData.endTime,eventData.firstName+" "+eventData.lastName,eventData.attendees,eventData.descriptionFr,eventData.descriptionEng,eventData.agendaName,eventData.idAgenda,eventData.address,eventData.eventTitle,"",eventData.eventId,eventData.creationDate]);
  if(topic == "Online-Event"){
    eventData.type = "Event";
    eventData.agendaName = "Events";
    eventData.idAgenda = eventAgendaId;
    eventData.eventId = eventData.eventId+"@"+bookedItem.name;
    Logger.log(eventAgendaId);
    sheet.appendRow(["submitted","",eventData.type,eventData.topic,eventData.eventDate,eventData.startTime,eventData.endTime,eventData.eventDate+" "+eventData.startTime,eventData.eventDate+" "+eventData.endTime,eventData.firstName+" "+eventData.lastName,eventData.attendees,eventData.descriptionFr,eventData.descriptionEng,eventData.agendaName,eventData.idAgenda,eventData.address,eventData.eventTitle,"",eventData.eventId,eventData.creationDate]);
    
  }  
  result = bookedItem.name+" has been submitted ! A confirmation email will be sent. You're all set.";
  return result;
}

//function to find an object in an Array because Google script.
function findObjectByKey(array, key, value) {
    for (var i = 0; i < array.length; i++) {
        if (array[i][key] === value) {
            return array[i];
        }
    }
    return null;
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
