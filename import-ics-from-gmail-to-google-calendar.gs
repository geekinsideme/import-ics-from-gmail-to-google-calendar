function myFunction() {
    var Timezone = function() {};
    var Event = function() {};
    var Tag = function() {};
    var n;

    function extractICSProperty(eventstr, tag) {
        var match;
        var val = new Tag();
        val.value = "";
        val.parameter = "";
        val.paramValue = "";
        if ((match = eventstr.match(new RegExp("^" + tag + "(;?)([^=:]*?)([=]?)([^=:]*?):(.*?)(\\\\n)*$", "im"))) !== null) {
            val.value = match[5].replace(/\\n/g, "\n");
            val.parameter = match[2];
            val.paramValue = match[4];
        }
        return val;
    }

    function convertDate(str, tz) {
        var match;
        if ((match = str.match(/(\d{4})(\d{2})(\d{2})$/)) !== null) {
            str = match[1] + "/" + match[2] + "/" + match[3];
        } else {
            if (str.match(/Z$/)) {
                str = str.replace(/(\d{4})(\d{2})(\d{2})T(\d{2})(\d{2})(\d{2})Z/, "$1/$2/$3 $4:$5:$6 +0000");
            } else {
                str = str.replace(/(\d{4})(\d{2})(\d{2})T(\d{2})(\d{2})(\d{2})/, "$1/$2/$3 $4:$5:$6 " + tz);
            }
        }
        return new Date(str);
    }

    var formatDate = function(date, format) {
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
    var messages = [];
    var thds = GmailApp.search("filename:" + icsFileName);
    for (var nt in thds) {
        var meses = thds[nt].getMessages();
        for (var nm in meses) {
            messages.push(meses[nm]);
        }
    }

    if (messages.length === 0) {
        Logger.log("No Messages found.");
        return;
    }

    var newestMessage;
    var newestDate = new Date("2000/1/1");
    for (n in messages) {
        if (messages[n].getDate() > newestDate) {
            newestDate = messages[n].getDate();
            newestMessage = messages[n];
        }
    }
    Logger.log("Newest message '" + newestMessage.getSubject() + "'@" + newestDate);

    for (n in messages) {
        if (messages[n] !== newestMessage) {
            messages[n].moveToTrash();
            Logger.log("Old message '" + messages[n].getSubject() + "'@" + messages[n].getDate() + " was trashed.");
        }
    }

    var attachment = newestMessage.getAttachments()[0];
    Logger.log("Newest attachment name and size : " + attachment.getName() + " has " + attachment.getSize() + " bytes");

    ics = attachment.getDataAsString();

    newestMessage.moveToTrash();
    Logger.log("Newest message '" + newestMessage.getSubject() + "'@" + newestDate + " was trashed.");

    ics = ics.replace(/[\n\r]+[ \t]+/g, "");

    var timezone = new Timezone();
    timezone.id = "";
    timezone.offset = "";
    var regTimezone = /EGIN:VTIMEZONE[\s\S]*?END:VTIMEZONE/g;
    var timezoneDef;
    while ((timezoneDef = regTimezone.exec(ics)) !== null) {
        timezone.id = extractICSProperty(timezoneDef[0], "TZID").value;
        timezone.offset = extractICSProperty(timezoneDef[0], "TZOFFSETTO").value;
    }
    Logger.log("TIMEZONE=" + timezone.offset);

    var events = [];
    var regEvent = /BEGIN:VEVENT[\s\S]*?END:VEVENT/g;
    var eventDef;
    var matches;
    while ((eventDef = regEvent.exec(ics)) !== null) {
        var event = new Event();
        var st = extractICSProperty(eventDef[0], "DTSTART");
        var et = extractICSProperty(eventDef[0], "DTEND");
        if (st.paramValue == "DATE") {
            event.allDay = true;
            event.date = convertDate(st.value, timezone.offset);
        } else {
            event.allDay = false;
            event.startTime = convertDate(st.value, timezone.offset);
            event.endTime = convertDate(et.value, timezone.offset);
        }
        event.title = extractICSProperty(eventDef[0], "SUMMARY").value;
        event.location = extractICSProperty(eventDef[0], "LOCATION").value;
        event.description = extractICSProperty(eventDef[0], "DESCRIPTION").value;
        event.visibility = extractICSProperty(eventDef[0], "CLASS").value;
        var reminderValue = extractICSProperty(eventDef[0], "TRIGGER").value;
        var minutesBefore;
        if ((matches = reminderValue.match(/([+-])P(\d+W|)(\d+D|)T?(\d+H|)(\d+M|)(\d+S|)/)) !== null) {
            minutesBefore = Number(matches[2].replace(/W/, "")) * 7 * 24 * 60 +
                Number(matches[3].replace(/D/, "")) * 24 * 60 +
                Number(matches[4].replace(/H/, "")) * 60 +
                Number(matches[5].replace(/M/, ""));
            if (matches[1] == "+") minutesBefore = -minutesBefore;
            event.reminder = minutesBefore;
        } else {
            event.reminder = "";
        }
        var rrule;
        if ((rrule = extractICSProperty(eventDef[0], "RRULE").value) !== "") {
            event.rrule = rrule;
        } else {
            event.rrule = "";
        }
        events.push(event);
    }

    var cals = CalendarApp.getCalendarsByName(calendarName);
    var cal = null;
    if (cals.length === 0) {
        Logger.log("Calendar '" + calendarName + "' not found.");
    } else {
        cal = cals[0];
    }
    var deleteEvents = cal.getEvents(new Date("1900/01/01"), new Date("2199/12/31"));
    Logger.log("DELETING " + deleteEvents.length + " event(s)");
    for (n in deleteEvents) {
        deleteEvents[n].deleteEvent();
    }

    Logger.log("INSERTING " + events.length + " event(s)");
    for (n in events) {
        var gcalEvent;
        var recurrence;
        var list = [];
        var regList;
        var lists;
        var m;
        if (events[n].rrule !== "") {
            if ((matches = events[n].rrule.match(/FREQ=(\w+)/)) !== null) {
                switch (matches[1]) {
                    case "YEARLY":
                        recurrence = CalendarApp.newRecurrence().addYearlyRule();
                        break;
                    case "MONTHLY":
                        recurrence = CalendarApp.newRecurrence().addMonthlyRule();
                        break;
                    case "WEEKLY":
                        recurrence = CalendarApp.newRecurrence().addWeeklyRule();
                        break;
                    case "DAILY":
                        recurrence = CalendarApp.newRecurrence().addDailyRule();
                        break;
                }
            }
            if ((matches = events[n].rrule.match(/UNTIL=(\w+)/)) !== null) {
                recurrence.until(convertDate(matches[1], timezone));
            }
            if ((matches = events[n].rrule.match(/COUNT=(\w+)/)) !== null) {
                recurrence.times(Number(matches[1]));
            }
            if ((matches = events[n].rrule.match(/INTERVAL=(\w+)/)) !== null) {
                recurrence.interval(Number(matches[1]));
            }
            if ((matches = events[n].rrule.match(/(BYMONTH|BYWEEKNO|BYYEARDAY|BYMONTHDAY)=([\w,]+)/)) !== null) {
                list = [];
                regList = /\d+/g;
                lists = matches[2];
                while ((m = regList.exec(lists)) !== null) {
                    list.push(Number(m[0]));
                }
                switch (matches[1]) {
                    case "BYMONTH":
                        recurrence.onlyInMonths(list);
                        break;
                    case "BYWEEKNO":
                        recurrence.onlyOnWeeks(list);
                        break;
                    case "BYYEARDAY":
                        recurrence.onlyOnYearDays(list);
                        break;
                    case "BYMONTHDAY":
                        recurrence.onlyOnMonthDays(list);
                        break;
                }
            }
            if ((matches = events[n].rrule.match(/BYDAY=([\w,]+)/)) !== null) {
                var weekday = {
                    "SU": CalendarApp.Weekday.SUNDAY,
                    "MO": CalendarApp.Weekday.MONDAY,
                    "TU": CalendarApp.Weekday.TUESDAY,
                    "WE": CalendarApp.Weekday.WEDNESDAY,
                    "TH": CalendarApp.Weekday.THURSDAY,
                    "FR": CalendarApp.Weekday.FRIDAY,
                    "SA": CalendarApp.Weekday.SATURDAY
                };
                list = [];
                regList = /\w+/g;
                lists = matches[1];
                while ((m = regList.exec(lists)) !== null) {
                    list.push(weekday[m[0]]);
                }
                recurrence.onlyOnWeekdays(list);
            }
            if (events[n].allDay) {
                gcalEvent = cal.createAllDayEventSeries(events[n].title, events[n].date, recurrence, {
                    location: events[n].location,
                    description: events[n].description
                });
            } else {
                gcalEvent = cal.createEventSeries(events[n].title, events[n].startTime, events[n].endTime, recurrence, {
                    location: events[n].location,
                    description: events[n].description
                });
            }
        } else {
            if (events[n].allDay) {
                gcalEvent = cal.createAllDayEvent(events[n].title, events[n].date, {
                    location: events[n].location,
                    description: events[n].description
                });
            } else {
                gcalEvent = cal.createEvent(events[n].title, events[n].startTime, events[n].endTime, {
                    location: events[n].location,
                    description: events[n].description
                });
            }
        }
        if (events[n].reminder !== "") {
            gcalEvent.addPopupReminder(events[n].reminder);
        }
        switch (events[n].visibility) {
            case "PUBLIC":
                gcalEvent.setVisibility(CalendarApp.Visibility.PUBLIC);
                break;
            case "PRIVATE":
                gcalEvent.setVisibility(CalendarApp.Visibility.PRIVATE);
                break;
            default:
                gcalEvent.setVisibility(CalendarApp.Visibility.DEFAULT);
        }
        Utilities.sleep(1000);
    }
    var now = new Date();
    cal.createEvent("Imported @" + formatDate(now), now, now);
}
