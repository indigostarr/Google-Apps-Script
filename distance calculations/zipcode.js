// read values and create variables
function findClosestCity() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var cities = spreadsheet.getSheets()[0]; //
  var baseLocs = spreadsheet.getSheets()[1];
  var startRow = 2; // First row of data to process
  var cityNumRows = cities.getLastRow() - 1; // Number of rows to process
  var cityLastColumn = cities.getLastColumn(); // Last column
  var cityDataRange = cities.getRange(startRow, 1, cityNumRows, cityLastColumn); // Fetch the data range of the active sheet
  var cityData = cityDataRange.getValues(); // Fetch values for each row in the rangeFind and replace
  var baseNumRows = baseLocs.getLastRow() - 1; // Number of rows to process
  var baseLastColumn = baseLocs.getLastColumn(); // Last column
  var baseDataRange = baseLocs.getRange(
    startRow,
    1,
    baseNumRows,
    baseLastColumn
  ); // Fetch the data range of the active sheet
  var baseData = baseDataRange.getValues(); // Fetch values for each row in the rangeFind and replace

  function availCities() {
    // loop through city in cities
    for (var i = 0; i < cityData.length; i++) {
      var row = cityData[i];
      var state = row[0];

      // get coords
      var lat = row[2];
      var long = row[3];

      // blank cells to input city, distance, and zone
      var closestLoc = row[11];
      // var distanceFromCityMatch = row[cityLastColumn + 2];
      // var zoneMatch = row[cityLastColumn + 3];

      // create comparison / counter variables
      let shortestDistance = 3000;
      let closestbaseLocZone = "";
      let closestbaseLoc = "";

      // for any new cities that haven't been assesed
      if (closestLoc != "") {
        continue;
      }
      {
        // loop through base locations
        for (var j = 0; j < baseData.length; j++) {
          var mRow = baseData[j];
          var matchCity = mRow[1];
          var mState = mRow[2];
          var mZone = mRow[4];

          // get base loc coords
          var mLat = mRow[7];
          var mLong = mRow[8];

          // for any base loc where state is the same as the city
          if (mState == state) {
            // use the city and base loc coordinates to calculate the distance between them
            var distanceFromCity = getDistanceFromLatLonInKm(
              lat,
              long,
              mLat,
              mLong
            );

            // if the distance is closer than the latest comparison city assign the city, distance and zone comparison variables
            if (distanceFromCity < shortestDistance) {
              // round distance and convert to miles
              shortestDistance = Math.round(distanceFromCity) * 0.621371;
              closestbaseLoc = matchCity;
              closestbaseLocZone = mZone;
            }

            // after looping through all avail base loc cities set values
            cities
              .getRange(startRow + i, cityLastColumn + 1)
              .setValue(closestbaseLoc);
            cities
              .getRange(startRow + i, cityLastColumn + 2)
              .setValue(shortestDistance);
            cities
              .getRange(startRow + i, cityLastColumn + 3)
              .setValue(closestbaseLocZone);
          }
        }
      }
    }
  }

  // run avail cities func
  availCities();

  // call back func for calculating distance between cities
  function getDistanceFromLatLonInKm(lat1, lon1, lat2, lon2) {
    var R = 6371; // Radius of the earth in km
    var dLat = deg2rad(lat2 - lat1); // deg2rad below
    var dLon = deg2rad(lon2 - lon1);
    var a =
      Math.sin(dLat / 2) * Math.sin(dLat / 2) +
      Math.cos(deg2rad(lat1)) *
        Math.cos(deg2rad(lat2)) *
        Math.sin(dLon / 2) *
        Math.sin(dLon / 2);
    var c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
    var d = R * c; // Distance in km
    return d;
  }

  // call back
  function deg2rad(deg) {
    return deg * (Math.PI / 180);
  }

  // Converts numeric degrees to radians
  function toRad(Value) {
    return (Value * Math.PI) / 180;
  }
}
