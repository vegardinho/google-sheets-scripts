/**
 * Create, update or delete calendar event based on parameters and existing 
 * eventsMake sure only one event exists; update if something 
 * has changed since creation
 * 
 * @param  {String}  evName  
 * @param  {String}  start      Start date
 * @param  {String}  end        End date
 * @param  {String}  prvName    
 * @param  {Boolean} dlt    
 *  
 * @return {undefined}      
 */
function updateCalendarEvent(evName, start, end, prvName, dlt=false) {
    var calendar = CalendarApp.getCalendarById(CAL_ID);
    var prvName = prvName;
    var evName = evName;
    var allDayEv = false;
    var events;
    var startTime = start;
    var endTime = end;
  
    // Google Calendar API treats end date as day after event has ended
    if (startTime.valueOf() === end.valueOf()) {
        endTime.setDate(endTime.getDate() + 1);



    }
  
  if (prvName !== null) {
    events = calendar.getEventsForDay(startTime, {search: prvName});
  } else {
    events = calendar.getEventsForDay(startTime, {search: evName});
  }    

    if (dlt) {
        dltEvMatches(events, evName);
        return;
    }




 
    switch(events.length) {
        default:
            dltEvMatches(events, evName, keepOne=true);
        case 1:
            if (!hasChanged(events[0], startTime, endTime, evName, prvName)) {
                return;
            }
            postEv(startTime, endTime, calendar, evName, prvName, ev=events[0]);
            break;
        case 0:
            postEv(startTime, endTime, calendar, evName, prvName, ev=null);
    }
}


/**
 * Create if no event passed, else update event. 
 * 
 * @param  {Date}       start
 * @param  {Date}       end  
 * @param  {Calendar}   cal 
 * @param  {String}     newName
 * @param  {String}     oldName 
 * @param  {Event}      ev      
 * 
 * @return {undefined}         
 */
function postEv(start, end, cal, newName, oldName, ev=null) {
    var oldName = oldName;
    var newName = newName;
    var event = ev;
    var calendar = cal;

    if (event !== null) {
        event.setAllDayDates(start, end);
        if (newName !== null && oldName !== newName) {
          event.setTitle(newName);
        }
    } else {
        calendar.createAllDayEvent(newName, start, end);
    }
}


/**
 * Chekcs if event is different from dates and names passed
 * 
 * @param  {Event}      ev      
 * @param  {Date}       start   
 * @param  {Date}       end    
 * @param  {String}     oldName 
 * @param  {String}     newName 
 * 
 * @return {Boolean}            True if has changed; false otherwise  
 */
function hasChanged(ev, start, end, oldName, newName) {
  if (ev.getStartTime() !== start || ev.getEndTime() !== end) {
        return true;
    }
    if (newName !== null && oldName !== newName) {
        return true;
    }

    return false;
}


/**
 * Delete all or keep one, of event list passed
 * 
 * @param  {Event[]}    evs     
 * @param  {String}     evName  
 * @param  {Boolean}    keepOne     Wether to keep one, or delete all
 */
function dltEvMatches(evs, evName, keepOne=false) {
    var strtInd = 0;
    if (keepOne === true) {
        strInd = 1;
    }

    if (evs === null) {
       return;
    }
    for (let i = strtInd; i < evs.length; i++) {
        evs[i].deleteEvent();
    }
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