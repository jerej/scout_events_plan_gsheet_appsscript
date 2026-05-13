/**
 * Lookup solar data based on latitude & longitude.
 *
 *@param {"35.5"} string Latitude
 *@param {"-79.2e"} string Longitude
 *@param {"YYYY-MM-dd"} Datetime format
 *@param {"1|0"} boolean 1=Sunrise 0=Sunset
 *@customfunction
 *
 * https://stackoverflow.com/questions/48040528/import-sunrise-set-based-on-coordinates-into-google-sheet-using-api
 */
function SunRiseSet(lat,long,date,type) {
  var response = UrlFetchApp.fetch("https://api.sunrise-sunset.org/json?lat="+lat+"&lng="+long+"&date="+date);
  var json = response.getContentText();
  var data = JSON.parse(json);
  var sunrise = data.results.sunrise;
  var sunset = data.results.sunset;
  if (type == 1) {
    return sunrise;
  } else {
    return sunset;
  };
}

/**
 * Converts a datetime string to a datetime string in a targe timezone.
 *
 *@param {"October 29, 2016 1:00 PM CDT"} datetimeString Date, time and timezone.
 *@param {"Timezone"} timeZone Target timezone
 *@param {"YYYY-MM-dd hh:mm a z"} Datetime format
 *@customfunction
 */
function formatDate(datetimeString,timeZone,format) {
  var moment = new Date(datetimeString);
  if(moment instanceof Date && !isNaN(moment)){
    return Utilities.formatDate(moment, timeZone, format)
  } else {
    throw 'datetimeString can not be parsed as a JavaScript Date object'
  }
}
