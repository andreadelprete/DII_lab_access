// script configuration
/*var scriptProperties = PropertiesService.getScriptProperties();
scriptProperties.setProperty('SEND_EMAILS', 'false');
scriptProperties.setProperty('ADD_GUESTS', 'false');*/
var SEND_EMAILS = false;   // flag to enable/disable sending emails (for debug)
var ADD_GUESTS = false;   // flag to enable/disable adding guests (for debug)
var emails_to_send = {};  // dictionary of emails to send. recipient is the key, value is a dictionary with fields title, body

function validateEmail(email) {
    const re = /^(([^<>()\[\]\\.,;:\s@"]+(\.[^<>()\[\]\\.,;:\s@"]+)*)|(".+"))@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\])|(([a-zA-Z\-0-9]+\.)+[a-zA-Z]{2,}))$/;
    return re.test(String(email).toLowerCase());
}

function event_to_string(event){
  var event_info_string = "Event title: "+event.getTitle()+"\n";
  event_info_string += "Event description: "+event.getDescription()+"\n";
  var start = event.getStartTime();
  event_info_string += "Event date: "+ start.getDate() + "/"+ (1+start.getMonth())+"/"+start.getFullYear()+"\n";
  event_info_string += "\nYou can find more information at the website: https://sites.google.com/unitn.it/diiaccesscovid-19/home-page";
  return event_info_string;
}

function sendEmail(recipient, title, body){
  if(!SEND_EMAILS){
    return;
  }
  
  var email = {};
  if(recipient in emails_to_send){
    email = emails_to_send[recipient];
    email.body += "\n\n================= "+title+" ================\n\n";
  }
  else{
    email = {
      title: title,
      body: "Dear "+recipient+"\n\n"
    };
    emails_to_send[recipient] = email;
  }
  
  email.body += body;
}


function cannot_read_calendar(workstation){
  Logger.log("Could not read calendar: "+workstation.name+". Notify lab manager "+workstation.manager_email);
  var email_title = "DII Lab Access - Could not read calendar associated to your laboratory";
  var email_body = "The script could not read the calendar "+workstation.name+" associated to your laboratory "+workstation.lab_name+".\n";
  email_body += "Please make sure that the calendar ID specified in the following spreadsheet is correct:\n";
  email_body += "https://docs.google.com/spreadsheets/d/1lrUj-WIXsqInLcQ3GBK9CIf10ZzQ3zngIeMJqV75tXs/edit#gid=753719539.\n";
  email_body += "Calendar ID:\n"+workstation.calendar_ID+"\n";
  email_body += "\nMoreover, make sure that andrea.delprete@unitn.it has access to the calendar.\n";
  email_body += "You can find more information at the website: https://sites.google.com/unitn.it/diiaccesscovid-19/home-page";
  sendEmail(workstation.manager_email, email_title, email_body);
}


function invite_lab_manager(event, workstation, applicant){
  try {
    Logger.log("Invite Lab manager "+workstation.manager_email+" and send email to applicant "+applicant);
    if(ADD_GUESTS){
      event.addGuest(workstation.manager_email);
    }
    var email_title = "DII Lab Access - Lab manager invited to event";
    var email_body = "The lab manager has been invited to the event you have created (or someone has created on your behalf) for the workstation "+workstation.name+" of the laboratory "+workstation.lab_name+".\n";
    email_body += event_to_string(event);
    sendEmail(applicant, email_title, email_body);
  }
  catch(err) {
    Logger.log(err.toString()+" Error trying to invite lab manager "+workstation.manager_email);
  }
}


function invite_supervisor(applicant, user, event, workstation, cal){
  Logger.log("Invite Supervisor "+user.supervisor_email+" and send emails to both supervisor and applicant "+applicant);
  if(ADD_GUESTS){
    event.addGuest(user.supervisor_email);
  }
  
  if(user.supervisor_email != workstation.manager_email){
    var event_info_string = event_to_string(event);
    
    var email_title = "DII Lab Access - Request from "+applicant;
    var email_body = applicant+" has requested to access "+cal.getName()+" on "+event.getStartTime().toDateString()+". \n";
    email_body += "You can allow or deny this access on your personal calendar: https://calendar.google.com/calendar/r \n";
    email_body += event_info_string;
    sendEmail(user.supervisor_email, email_title, email_body);
    
    email_title = "DII Lab Access - Supervisor invited to event";
    email_body = "Your supervisor "+user.supervisor_email+" has been invited to the event you have created (or someone has created on your behalf) for the workstation ";
    email_body += workstation.name+" of the laboratory "+workstation.lab_name+".\n";
    email_body += event_info_string;
    sendEmail(applicant, email_title, email_body);
  }
}

function user_not_found(creator, applicant, workstation, cal, event){
  // if researcher not found then do not authorize
  Logger.log("Applicant not found in user list: "+applicant+". Send email to event creator "+creator+" and cancel event.");
  var email_title = "DII Lab Access - User "+applicant+" not found";
  var email_body = "You have created an event in the calendar "+cal.getName()+" on "+event.getStartTime().toDateString()+" on behalf of "+applicant+". \n";
  email_body += "However, the email address "+applicant+" was not found in the tab Users of this spreadsheet:\n";
  email_body += "https://docs.google.com/spreadsheets/d/1lrUj-WIXsqInLcQ3GBK9CIf10ZzQ3zngIeMJqV75tXs/edit#gid=753719539.\n";
  email_body += "For this reason the event has been cancelled.\n" + event_to_string(event);
  sendEmail(creator, email_title, email_body);
}

/** Check whether the specified guest is invited to the specified event. 
If so it returns the answer, otherwise it returns false. */
function is_in_guest_list(event, guest){
  var guests = event.getGuestList();
  for (var j=0; j<guests.length; j++) {
    //Logger.log("Check guest "+guests[j].getEmail());
    if(guests[j].getEmail()==guest){
      var answer = guests[j].getGuestStatus();
      //Logger.log("Answer of "+guest+": "+answer)
      return answer;
    }
  }
  return false;
}


function export_cal_to_sheet(workstation, date_start, date_end, first_row, users, position_dict, output_sheet, db_sheet){
  // Export Google Calendar Events to a Google Spreadsheet
  // 
  // Reference Websites:
  // https://developers.google.com/apps-script/reference/calendar/calendar
  // https://developers.google.com/apps-script/reference/calendar/calendar-event
  
  var cal = CalendarApp.getCalendarById(workstation.calendar_ID);
  Logger.log("Reading calendar ID: "+workstation.calendar_ID+" of "+workstation.manager_email);
  if(!cal){
    cannot_read_calendar(workstation);
    return first_row;
  }
  
  // Optional variations on getEvents
  // var events = cal.getEvents(new Date("January 3, 2014 00:00:00 CST"), new Date("January 14, 2014 23:59:59 CST"), {search: 'word1'});
  var events = cal.getEvents(date_start, date_end);
  
  // Loop through all calendar events found and write them out starting on specified ROW
  var row = first_row;
  for (var i=0;i<events.length;i++) {
    var e = events[i];
    var creator = e.getCreators()[0];
    var applicant = creator;
    var title = e.getTitle().trim();
    var guests = e.getGuestList();
    var start = e.getStartTime();
    var access_date = start.getDate() + "/"+ (1+start.getMonth())+"/"+start.getFullYear();
    var authorized_lab = "No";            // authorization of the lab manager
    var authorized_supervisor = "YES";    // authorization of the supervisor
    
    Logger.log("Process event "+title+" created by "+creator);
    // if the event title is "Closed" then it's not a real reservation, so don't insert it in the spreadsheet
    if(title=="Closed"){
      continue;
    }
    
    // if the event title is an email address => consider it as the applicant
    if(validateEmail(title)){
      Logger.log("Event title is an email address so consider it as applicant: "+title);
      applicant = title;
    }    
    
    // first look for authorization from lab manager
    if(creator==workstation.manager_email){
      // if the person asking is the lab manager => authorize
      Logger.log("The even creator is lab manager => authorize")
      authorized_lab = "YES";
    }
    else{
      // otherwise check for the authorization of the lab head
      authorized_lab = is_in_guest_list(e, workstation.manager_email);
      Logger.log("Authorization of lab manager "+workstation.manager_email+": "+authorized_lab);
      // if lab head has not been invited => invite him/her
      if(!authorized_lab){
        invite_lab_manager(e, workstation, applicant);
      }
    }
    
    // look for researcher in user list
    var user = users[applicant];
    if(!user){
      user_not_found(creator, applicant, workstation, cal, e);
      e.deleteEvent();
      continue;
    }
    
    // if user has a supervisor => look for authorization from supervisor
    if(user.supervisor_email!="" && user.supervisor_email!=creator){
      authorized_supervisor = is_in_guest_list(e, user.supervisor_email);
      Logger.log("Authorization of supervisor "+user.supervisor_email+": "+authorized_supervisor);
      // if supervisor has not been invited => invite him/her
      if(!authorized_supervisor){
        invite_supervisor(applicant, user, e, workstation, cal);
      }
    }
    
    if(authorized_lab=="YES" || authorized_supervisor=="YES"){
      var db_last_row = 1+db_sheet.getDataRange().getLastRow();
      var row_values=[[access_date, cal.getId(), e.getId(), authorized_lab, authorized_supervisor]];
      db_sheet.getRange(db_last_row,1,1,5).setValues(row_values);
    }
    
    // insert event in spreadsheet
    var end = e.getEndTime();
    // start.toLocaleString('it-IT', { timeZone: 'Europe/Berlin' })
    // "Applicant", "Position", "Title", "Room", "Block", "Access Date", "Time", "Director Auth", "Auth",
    //"Titolo", "Descrizione", "Ora", "Fine", "Ora", "Responsabile Lab", "Autorizzato (Resp. Lab)", "Supervisore", "Autorizzato (supervisore)"
    var position_code = position_dict[user.position];
    var room = cal.getName();
    //var access_date = start.toDateString();

    var time = "";
    var authorization = "";
    if(authorized_lab=="YES" && authorized_supervisor=="YES"){
      authorization = "x";
    }
    var details=[[applicant, user.position, position_code, room, workstation.block, access_date, time, authorization, workstation.manager_name,
                  title, e.getDescription(), start.toLocaleTimeString(), end.toDateString(), end.toLocaleTimeString(), 
                  workstation.manager_email, authorized_lab, user.supervisor_email, authorized_supervisor]];
    var range=output_sheet.getRange(row,1,1,18);
    range.setValues(details);
    row += 1;
  }
  return row;
}


function main(){
  // Export Google Calendar Events to a Google Spreadsheet
  //var date_start = new Date("May 20, 2020 00:00:00 GMT");
  var date_start = new Date();
  date_start.setDate(date_start.getDate() - 1); // start looking at events from yesterday
  var date_end = new Date();
  date_end.setDate(date_end.getDate() + 14); // until 14 days into the future
  Logger.log("Start date: "+ date_start);
  Logger.log("End date: "+ date_end);
  
  //var sheet = SpreadsheetApp.getActiveSheet();
  var out_sprsheet = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1hcVgmhjE3uGkbK0EK5v3lwM4rqdWskTqhcGpQ_qMGgw/edit#gid=0");
  var output_sheet = out_sprsheet.getSheetByName("Richieste");
  
  var in_sprsheet = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1lrUj-WIXsqInLcQ3GBK9CIf10ZzQ3zngIeMJqV75tXs/edit#gid=0");
  var locations_sheet = in_sprsheet.getSheetByName("Workstations");
  var supervisors_sheet = in_sprsheet.getSheetByName("Users");
  
  var positions_sprsheet = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1LN8oJSspiu51_cYZrzlB54mqgHAgXOEQb6doEDdnP8o/edit#gid=997932956");
  var positions_sheet = positions_sprsheet.getSheetByName("PositionCodes");
  
  var db_sprsheet = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1me-zaf5H-zK7fPlXh4Jzbho55fyHrlAL7DLxJXii0ZY/edit#gid=0");
  var db_sheet = db_sprsheet.getSheetByName("Events");
  
  
  // Uncomment this next line if you want to always clear the spreadsheet content before running - Note people could have added extra columns on the data though that would be lost
  output_sheet.clearContents();  
  
  // Create a header record on the current spreadsheet in cells A1:N1
  var header = [["Applicant", "Position", "Title", "Room", "Block", "Access Date", "Time", "Director Auth", "Auth",
                 "Titolo Evento", "Descrizione", "Ora Inizio", "Giorno Fine", "Ora Fine",  
                 "Responsabile Lab", "Autorizzato (Resp. Lab)", "Supervisore", "Autorizzato (supervisore)"]]
  var range = output_sheet.getRange(1,1,1,18);
  range.setValues(header);
  
  // Read locations sheet
  // 0 Postazione	1 Laboratorio	2 Edificio	3 Responsabile	4 Email responsabile postazione	5 ID Calendario
  var data = locations_sheet.getDataRange().getValues();
  var workstations = [];
  for (var i = 1; i < data.length; i++) {
    //Logger.log('Manager: ' + data[i][3] + '; Calendar ID: ' + data[i][5]);
    var ws = {};
    ws.name = data[i][0];
    ws.lab_name = data[i][1];
    ws.block = data[i][2];
    ws.manager_name = data[i][3];
    ws.manager_email = data[i][4].trim();
    ws.calendar_ID = data[i][5];
    workstations.push(ws);
  }
  
  // Read supervisors sheet
  // 0 Ricercatore	1 Email ricercatore	2 Posizione	3 Supervisore	4 Email supervisore
  data = supervisors_sheet.getDataRange().getValues();
  var users = {};
  for (var i = 1; i < data.length; i++) {
    //Logger.log('Researcher: ' + data[i][1] + '; supervisor: ' + data[i][4]);
    var user_email = data[i][1].trim();
    var supervisor_email = data[i][4].trim();
    if(validateEmail(user_email) && (validateEmail(supervisor_email) || supervisor_email=="")){
      var u = {};
      u.supervisor_email = supervisor_email;
      u.position = data[i][2];
      users[user_email] = u;
    }
  }
  
  // read position codes
  // 0 Position Code	1 Position App Unitn
  data = positions_sheet.getDataRange().getValues();
  var position_dict = {};
  for(var i=1; i<data.length; i++){
    var position_name = data[i][1];
    var position_code = data[i][0];
    position_dict[position_name] = position_code;
    //Logger.log("Position name: "+position_name+"; Position code: "+position_code);
  }

  var current_row = 2;
  for(var i=0; i<workstations.length; i++){
    current_row = export_cal_to_sheet(workstations[i], date_start, date_end, current_row, users, position_dict, output_sheet, db_sheet);  
  }
  
  Logger.log("Start sending emails");
  var email_counter = 0;
  for(var recipient in emails_to_send){
    var email = emails_to_send[recipient];
    Logger.log("Send email to "+recipient);
    //Logger.log(email.body);
    MailApp.sendEmail(recipient.trim(), email.title, email.body, {name: 'DII Access Emailer Script', noReply: true});
    email_counter += 1;
  }
  Logger.log("Total number of emails sent: "+email_counter);
  
  // save log as text file
  try{
    var folder = DriveApp.getRootFolder().getFoldersByName("DII_accessi_covid").next().getFoldersByName("logs").next();
    var now = new Date();
    var filename = now.getFullYear() + "/"+ (1+now.getMonth())+"/"+now.getDate()+" "+now.toTimeString()+".txt";
    folder.createFile(filename, Logger.getLog());
  }
  catch(err){
    Logger.log("Error while trying to save log as text file: "+err.toString());
  }
}

function onOpen() {
  Browser.msgBox('Istruzioni', 'Questo file contiene le richieste di accesso alle postazioni di lavoro del DII che sono lette automaticamente dai calendari elencati in questo file: https://docs.google.com/spreadsheets/d/1lrUj-WIXsqInLcQ3GBK9CIf10ZzQ3zngIeMJqV75tXs/edit#gid=0.', Browser.Buttons.OK);
}
