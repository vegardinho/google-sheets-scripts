// Create, update or delete calendar event based on parameters and existing events
// Make sure only one event exists; update if something has changed since creation
function updateCalendarEvent(evName, start, end, prvName, dlt=false) {
    var startTime = crtDateStr(start);
    var endTime = crtDateStr(end);
    var calendar = CalendarApp.getCalendarById(CAL_ID);
    var prvName = prvName;
    var evName = evName;
    var allDayEv = false;
    var events;
  
  if (prvName !== null) {
    events = calendar.getEventsForDay(startTime, {search: prvName});
  } else {
    events = calendar.getEventsForDay(startTime, {search: evName});
  }
    

    if (dlt) {
        dltEvMatches(events, evName);
        return;
    }

    // Google Calendar API treats end date as day after event has ended
    if (startTime.getHours() == startTime.getMinutes() == endTime.getHours() == endTime.getMinutes == 0) {
        allDayEv = true;
        endTime.setDate(endTime.getDate() + 1);
    }
 
    switch(events.length) {
        default:
            dltEvMatches(events, evName, keepOne=true);
        case 1:
            if (!hasChanged(events[0], startTime, endTime, evName, prvName)) {
                return;
            }
            postEv(startTime, endTime, allDayEv, calendar, evName, prvName, ev=events[0]);
            break;
        case 0:
            postEv(startTime, endTime, allDayEv, calendar, evName, prvName, ev=null);
    }
}

// Create if no event passed, else update event. 
function postEv(start, end, allDayEv, cal, newName, oldName, ev=null) {
    var oldName = oldName;
    var newName = newName;
    var event = ev;
    var calendar = cal;

    if (event !== null) {
        if (allDayEv) {
            event.setAllDayDates(start, end);
        } else {
            event.setTime(start, end);
        }
        if (newName !== null && oldName !== newName) {
          event.setTitle(newName);
        }
    } else {
        if (allDayEv) {
            calendar.createAllDayEvent(newName, start, end);
        } else {
            calendar.createEvent(newName, start, end);
        }
    }
}

function hasChanged(ev, start, end, oldName, newName) {
  if (ev.getStartTime() !== start || ev.getEndTime() !== end) {
        return true;
    }
    if (newName !== null && oldName !== newName) {
        return true;
    }

    return false;
}

// Delete all or keep one, of event list passed
function dltEvMatches(evs, evName, keepOne=false) {
    var strtInd = 0;
    if (keepOne === true) {
        strInd = 1;
    }

    for (let i = strtInd; i < evs.length; i++) {
        evs[i].deleteEvent();
    }
}

// Return Date-object from string on form 'dd.mm'
function crtDateStr(date, hrs='00', min='00') {
    year = M_END_DATE.substring(0,4);
    month = date.substring(3,5);
    dateNum = date.substring(0,2);
    dateStr = `${year}-${month}-${dateNum}T${hrs}:${min}+02:00`;
    Logger.log(dateStr);
    return new Date(dateStr);
}

/* Test-function:

   function runTest() {
//updateCalendarEvent("Morohendelse", '10.08', '11.08', dlt=true);  
// Create custom button text for Today and Tomorrow
var TodayButton = CardService.newTextButton()
.setText("Today");
var TomorrowButton = CardService.newTextButton()
.setText("Tomorrow");
var DayButtonSet = CardService.newButtonSet()
.addButton(TodayButton)
.addButton(TomorrowButton);

var responseWeatherImpactsDate = ui.alert('Do you want the first day of the Weather Impacts table to be today or tomorrow?', DayButtonSet);
}
 */