function myFunction(query_object) {
  var param_value_list = Object.entries(query_object);
  var param_strings = param_value_list.map(([param, value]) => `${param}=${value}`);
  var query_string = param_strings.join('&');
  return '?' + query_string;
}

function main()
{
  var x = {"jikael": 6, "ella": 4};
  console.log(myFunction(x))
}



function deadCode_() {
  return;

  function getDateTime(timeString) {
  var dateTime = new Date();

  var parts = timeString.split(':');
  var hours = parseInt(parts[0], 10);
  var minutes = parseInt(parts[1], 10);

  dateTime.setHours(hours, minutes, 0, 0); // Set the time

  return dateTime;
  }


  function getThresholdTime(startTime) {
    var dateTime = new Date();

    var parts = startTime.split(':');
    var hours = parseInt(parts[0], 10);
    var minutes = parseInt(parts[1], 10);

    dateTime.setHours(hours + 2, minutes, 0, 0); // Set the time
    return dateTime;
  }

  var emailBody = 
    "Here is a copy of your submission: \n\n- HEAD RUN: " + headRun + "\n- DISTANCE: " + distance + "\n\n------ ATTENDEES ------\n" + attendees + "\n\n*I declare all attendees have provided their waiver and paid the one-time member fee*  > " + confirmation + "\n\nComments: " + notes + "\n\nKeep up the amazing work!\n\nBest,\nMcRUN Team"
  ;

  var headRunTime = getHeadRunTime(todayWeekDay);
  if(headRunTime.length < 1) return;  // exit if no head run today

  var dateTime, thresholdTime;

  for(const time of headRunTime) {
    dateTime = getDateTime(time);   // convert to Date object
    thresholdTime = getThresholdTime(time);  // add 2 hours

    Logger.log(today);
    Logger.log(thresholdTime);

    if (today.setHours(today.getHours() + 2) ) {};
  }

  var test = new Date(submissionDates[0]).getDate();

  for (var i = 0; i < data.length; i++) {
    var cellValue = data[i][0];
    if (cellValue instanceof Date) {
      cellValue.setHours(0, 0, 0, 0); // Set the time to midnight for comparison
      if (cellValue.getTime() !== today.getTime()) {
        // If the date in the cell doesn't match today's date
        // Send a notification email
      }
    }
  }

  headRunTime.forEach(
    function(item) { Logger.log(item); }
  );

}
