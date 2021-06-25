"use strict";

const locationToCalculateFrom = document.querySelector(".city");
const baseLocationCalc = document.querySelector(".base-loc");
// const selectedbaseLoc = document.querySelector(".base-loc");

console.log(locationToCalculateFrom);
let shortestDistance = 3000;
let cityToCalculateFrom = "";
let closestbaseLoc = "";

let baseLoc = {
  "New York": {
    Albany: [42.63378, -73.76322],
    Bronx: [40.82077, -73.92387],
  },
  Maryland: {
    Baltimore: [39.29479, -76.6222],
    Rockville: [39.09129, -77.18085],
    Frederick: [39.44605, -77.33495],
  },
};
// console.log(baseLoc.Maryland);

let currentCity = {
  "New York": {
    "Fishers Island": [41.27024, -71.98783],
    Sloansville: [42.75887, -74.3672],
  },
  Maryland: {
    Waldorf: [38.62939, -76.97697],
    Abell: [38.25952, -76.73699],
    Accokeek: [38.67227, -77.02018],
  },
};
// console.log(currentCity.Maryland);

const cityMatch = function (currState, matchState) {
  for (let city in currState) {
    let currCity = currState[city];
    cityToCalculateFrom = city;
    // console.log(currCity);

    for (let loc in matchState) {
      let matchCity = matchState[loc];
      let distanceFromCity = getDistanceFromLatLonInKm(
        [...currCity],
        [...matchCity]
      );
      // console.log(matchCity);
      if (distanceFromCity < shortestDistance) {
        shortestDistance = distanceFromCity;
        closestbaseLoc = loc;
      }
    }
    locationToCalculateFrom.innerHTML += `<br>City: ${cityToCalculateFrom} | Distance is ${Math.round(
      shortestDistance
    )}km\b`;
    baseLocationCalc.innerHTML += `<br>Closest base Loc: ${closestbaseLoc}\b`;

    // locationToCalculateFrom.innerHTML = cityToCalculateFrom;
    // selectedbaseLoc.innerHTML = closestbaseLoc;
  }
};

function getDistanceFromLatLonInKm([lat1, lon1], [lat2, lon2]) {
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

function deg2rad(deg) {
  return deg * (Math.PI / 180);
}

// Converts numeric degrees to radians
function toRad(Value) {
  return (Value * Math.PI) / 180;
}

function availCities() {
  for (let state in currentCity) {
    let currState = currentCity[state];
    let matchState = baseLoc[state];
    cityMatch(currState, matchState);
  }
}

availCities();

Excel.run(function (context) {
  var sheets = context.workbook.worksheets;
  sheets.load("items/name");

  return context.sync().then(function () {
    if (sheets.items.length > 1) {
      console.log(
        `There are ${sheets.items.length} worksheets in the workbook:`
      );
    } else {
      console.log(`There is one worksheet in the workbook:`);
    }
    sheets.items.forEach(function (sheet) {
      console.log(sheet.name);
    });
  });
}).catch(errorHandlerFunction);

// document.querySelector(".city").innerHTML = FishersIsland;
// console.log(currentLoc[FishersIsland].value);
// console.log(
//   getDistanceFromLatLonInKm([...currentLoc[FishersIsland].value], [...nyc])
// );
// console.log(shortestDistance);

// database of arrays of values

// for each base location array
// console.log(currentLoc.Sloansville);
// loop through database of base loc
