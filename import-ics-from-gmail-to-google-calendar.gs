function myFunction() {
    if( typeof SitesApp === "undefined") {

        CalendarApp = function() {};
        CalendarApp.getCalendarsByName = function(name) {
            var calenders = [];
            calenders.push( new Calendar());
            return calenders;
        };

        Calendar = function() {
            var events = [];
            events.push( new CalendarEvent());
            this.events = events;
            return this;
        };
        Calendar.events = [];
        Calendar.prototype.getEvents = function(start,end) {
            return this.events;
        }
        Calendar.prototype.createEvent = function(title, startTime, endTime, options) {
            var event = new CalendarEvent();
            event.title = title;
            event.startTime = startTime;
            event.endTime = endTime;
            event.options = options;
            this.events.push( event );
            return event;
        }
        Calendar.prototype.createAllDayEvent = function(title, date, options) {
            var event = new CalendarEvent();
            event.title = title;
            event.date = date;
            event.options = options;
            this.events.push( event );
            return event;
        }

        CalendarEvent = function() {};
        CalendarEvent.prototype.deleteEvent = function() {};
        CalendarEvent.prototype.addPopupReminder = function(minutesBefore) {
            this.popupremainder = minutesBefore;
            return this;
        }
        CalendarEvent.prototype.setVisibility = function(visibility) {
            this.visibility = visibility;
            return this;
        }

        CalendarApp.Visibility = function() {};
        CalendarApp.Visibility.CONFIDENTIAL = 1;
        CalendarApp.Visibility.DEFAULT = 2;
        CalendarApp.Visibility.PRIVATE = 3;
        CalendarApp.Visibility.PUBLIC = 4;

        Logger = function() {};
        Logger.log = console.log;
        console.log("Run in node.js");
    } else {
        Logger.log("Run in Google Apps");
    }

    var Timezone = function() {};
    var Event = function() {};
    var Tag = function() {};

    function extractICSProperty(eventstr,tag) {
        var match;
        var val = new Tag();
        val.value = "";
        val.parameter = "";
        val.paramValue = "";
        if( match = eventstr.match(new RegExp("^"+tag+"(;?)([^=:]*?)([=]?)([^=:]*?):(.*?)(\\\\n)*$","im")) ) {
            val.value = match[5].replace(/\\n/g,"\n");
            val.parameter = match[2];
            val.paramValue = match[4];
        }
        return val;
    }

    function convertDate( str,tz ) {
        var match;
        if( match = str.match(/(\d{4})(\d{2})(\d{2})$/) ) {
            str = match[1]+"/"+match[2]+"/"+match[3];
        }else{
            if( str.match(/Z$/) ) {
                str = str.replace(/(\d{4})(\d{2})(\d{2})T(\d{2})(\d{2})(\d{2})Z/,"$1/$2/$3 $4:$5:$6 +0000");
            } else {
                str = str.replace(/(\d{4})(\d{2})(\d{2})T(\d{2})(\d{2})(\d{2})/,"$1/$2/$3 $4:$5:$6 "+tz);
            }
        }
        return new Date( str );
    }

    var formatDate = function (date, format) {
      if (!format) format = 'YYYY-MM-DD hh:mm:ss';
      format = format.replace(/YYYY/g, date.getFullYear());
      format = format.replace(/MM/g, ('0' + (date.getMonth() + 1)).slice(-2));
      format = format.replace(/DD/g, ('0' + date.getDate()).slice(-2));
      format = format.replace(/hh/g, ('0' + date.getHours()).slice(-2));
      format = format.replace(/mm/g, ('0' + date.getMinutes()).slice(-2));
      format = format.replace(/ss/g, ('0' + date.getSeconds()).slice(-2));
      if (format.match(/S/g)) {
        var milliSeconds = ('00' + date.getMilliseconds()).slice(-3);
        var length = format.match(/S/g).length;
        for (var i = 0; i < length; i++) format = format.replace(/S/, milliSeconds.substring(i, i + 1));
      }
      return format;
    };

    // ****** DEFINITION *******
    var icsFileName = "outlook.ics";
    var calendarName = "Outlook";

    var ics;
    if( typeof SitesApp === "undefined") {
        var fs = require('fs');
        ics = fs.readFileSync('./'+icsFileName, 'utf8');
    } else {
        var messages = new Array();
        var thds = GmailApp.search("filename:"+icsFileName);
        for(var nt in thds){
            var meses = thds[nt].getMessages();
            for(var nm in meses){
            messages.push( meses[nm] );
            }
        }

        if( messages.length == 0 ) {
          Logger.log("No Messages found.");
          return;
        }

        var newestMessage;
        var newestDate = new Date("2000/1/1");
        for(var n in messages){
            if( messages[n].getDate() > newestDate){
                newestDate = messages[n].getDate();
                newestMessage = messages[n];
            }
        }
        Logger.log("Newest message '"+newestMessage.getSubject()+"'@"+newestDate);

        for(var n in messages){
            if( messages[n] !== newestMessage){
                messages[n].moveToTrash();
                Logger.log("Old message '"+messages[n].getSubject()+"'@"+messages[n].getDate()+" was trashed.");
            }
        }

        var attachment = newestMessage.getAttachments()[0];
        Logger.log("Newest attachment name and size : "+attachment.getName() + " has "+attachment.getSize()+" bytes");

        ics = attachment.getDataAsString();

        newestMessage.moveToTrash();
        Logger.log("Newest message '"+newestMessage.getSubject()+"'@"+newestDate+" was trashed.");
    }

    ics = ics.replace(/[\n\r]+[ \t]+/g,"");

    var timezone = new Timezone();
    timezone.id = "";
    timezone.offset = "";
    var regTimezone=/EGIN:VTIMEZONE[\s\S]*?END:VTIMEZONE/g;
    var timezoneDef;
    while ((timezoneDef = regTimezone.exec(ics)) !== null) {
        timezone.id = extractICSProperty( timezoneDef[0],"TZID").value;
        timezone.offset = extractICSProperty( timezoneDef[0],"TZOFFSETTO").value;
    }
    Logger.log("TIMEZONE="+timezone.offset);

    var events = [];
    var regEvent=/BEGIN:VEVENT[\s\S]*?END:VEVENT/g;
    var eventDef;
    while ((eventDef = regEvent.exec(ics)) !== null) {
        var event = new Event();
        var st = extractICSProperty( eventDef[0],"DTSTART");
        var et = extractICSProperty( eventDef[0],"DTEND");
        if( st.paramValue == "DATE") {
            event.allDay = true;
            event.date = convertDate(st.value,timezone.offset);
        } else {
            event.allDay = false;
            event.startTime = convertDate(st.value,timezone.offset);
            event.endTime = convertDate(et.value,timezone.offset);
        }
        event.title = extractICSProperty( eventDef[0],"SUMMARY").value;
        event.location = extractICSProperty( eventDef[0],"LOCATION").value;
        event.description = extractICSProperty( eventDef[0],"DESCRIPTION").value;
        event.visibility = extractICSProperty( eventDef[0],"CLASS").value;
        var reminderValue = extractICSProperty( eventDef[0],"TRIGGER").value;
        var match;
        var minutesBefore;
        if( match = reminderValue.match(/([+-])P(\d+W|)(\d+D|)T?(\d+H|)(\d+M|)(\d+S|)/) ) {
            minutesBefore = Number(match[2].replace(/W/,""))*7*24*60
                + Number(match[3].replace(/D/,""))*24*60
                + Number(match[4].replace(/H/,""))*60
                + Number(match[5].replace(/M/,""))
                + Number(match[6].replace(/S/,""))/60.0;
            if( match[1]=="+" ) minutesBefore = -minutesBefore;
            event.reminder = minutesBefore;
        } else {
            event.reminder = "";
        }
        events.push(event);
        // Logger.log(event.visibility+" "+formatDate(event.startTime)+" - "+formatDate(event.endTime)+" : "+event.title+" @"+event.location);
    }

    var cals = CalendarApp.getCalendarsByName(calendarName);
    var cal = null;
    if (cals.length == 0){
        Logger.log("Calendar '"+calendarName+"' not found.");
    } else {
        cal = cals[0];
    }
    var deleteEvents = cal.getEvents( new Date("1900/01/01"),new Date("2199/12/31"));
    Logger.log("DELETING "+deleteEvents.length+" event(s)");
    for(var n in deleteEvents){
        deleteEvents[n].deleteEvent();
    }
    var now = new Date();
    cal.createEvent("Import @"+formatDate(now), now, now);

    Logger.log("INSERTING "+events.length+" event(s)");
    for(var n in events){
        var event;
        if( events[n].allDay ) {
            event = cal.createAllDayEvent( events[n].title,events[n].date,
                {location:events[n].location, description:events[n].description}
            );
        } else {
            event = cal.createEvent( events[n].title,events[n].startTime,events[n].endTime,
                {location:events[n].location, description:events[n].description}
            );
        }
        if( events[n].reminder!="" ) {
            event.addPopupReminder( events[n].reminder );
        }
        switch (events[n].visibility) {
            case "PUBLIC":
                event.setVisibility(CalendarApp.Visibility.PUBLIC);
                break;
            case "PRIVATE":
                event.setVisibility(CalendarApp.Visibility.PRIVATE);
                break;
            default:
            event.setVisibility(CalendarApp.Visibility.DEFAULT);
        }
    }

    if( typeof SitesApp === "undefined") {
        var insertedEvents = cal.getEvents( new Date("1900/01/01"),new Date("2199/12/31"));
        Logger.log("Inserted Events = "+insertedEvents.length);
        for(var n in insertedEvents){
            Logger.log(insertedEvents[n]);
        }
    }
}
if( typeof SitesApp === "undefined") {
    myFunction();
}
