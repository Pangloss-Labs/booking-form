//Header
var appsTitle = "Test Booking Form";
var appsIcon = "https://upload.wikimedia.org/wikipedia/commons/thumb/c/c6/Google_favicon.png/220px-Google_favicon.png";


//Booking type
var room = {item:"Room", style:"uk-tile-default"};
var machine = {item:"Machines", style:"uk-tile-muted"};
var event = {item:"Event", style:"uk-tile-primary"};
var bookingTypes = [room,machine,event];

//Room
var sampleRoom = {name:"sample",description:"Meeting room équiped with TV or projector" ,people:5 , location:"RdC", cost:25, costUnit:"€/h", id:"sampleRoom@group.calendar.google.com"};
var sampleroom2 = {name:"sample2",description:"Co-working room available after 18:00" ,people:20 , location:"RdC", cost:80, costUnit:"€/h", id:"sampleroom2@group.calendar.google.com"};

//Machines : ALL NAMES must be lower case : 
var sampleMachine = {name:"",description:" - owner: email@sample.org" ,type:"" , location:"sampleroom", cost:8, costUnit:"€/h", id:"sample@group.calendar.google.com"};
var sample2 = {name:"",description:" - owner: email@sample.org" ,type:"" , location:"sampleroom", cost:8, costUnit:"€/h", id:"sample@group.calendar.google.com"};


var machineList = [sampleMachine,sample2];
//Form
var dateMax = "2019-01-31"; //use a today + X days function ?

//event list sheet
//event calendar ID
var eventAgendaId = "sample@group.calendar.google.com";
//google sheet ID
  var googleSheet = SpreadsheetApp.openById("XXXXXXXXXX");
  var sheet = googleSheet.getSheets()[0];
