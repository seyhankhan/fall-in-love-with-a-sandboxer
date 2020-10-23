///////////// Seyhan Van Khan
///////////// The Pair Calculator
///////////// From a spreadsheet of people, find each person a new match every month
///////////// October 2020
///////////// github.com/seyhanvankhan

/* -------------------------------- CONSTANTS ------------------------------- */

// index of each column in Sandboxers array
const COLUMN = {
  "Name":           0,
  "Number":         1,
  "Email":          2,
  "Location":       3,
  "Link":           4,
  "Comments":       5,
  "Opted In":       6,
  "Remove if odd":  7,
  "Pairs":          8
};
const SPREADSHEET_URL = `https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}/values/${SPREADSHEET_RANGE}?key=${API_KEY}`;



/* -------------------------------- FUNCTIONS ------------------------------- */


function shuffle(array) {
  // https://bost.ocks.org/mike/shuffle/
  var m = array.length, t, i;

  // While there remain elements to shuffle…
  while (m) {
    // Pick a remaining element…
    i = Math.floor(Math.random() * m--);
    // And swap it with the current element.
    t = array[m];
    array[m] = array[i];
    array[i] = t;
  }
  return array;
}

function log(output) {
  console.log(output);
  document.getElementById('errorBox').innerHTML += output + '<br>';
}


/* -------------------------- GET SPREADSHEET DATA -------------------------- */


fetch(SPREADSHEET_URL)
.then(resp => resp.json())


/* ------------------------------ FORMAT DATA ------------------------------- */


.then(data => {
  // "Name"	"Number"	"Email"	"Location"	"Link"	"Comments"	"Opted In"	"The person to remove if odd number"	"Pairs ---->"
  // remove the first row (the headers)
  Sandboxers = data['values'].slice(1);

  // iterate through each sandboxer, saving each sandboxer's ID if they have opted in (participating in this month)
  // array of ints
  participatingIDs = [];
  for (id = 0; id < Sandboxers.length; id++) {
    if (Sandboxers[id][COLUMN["Opted In"]] === '1') {
      participatingIDs.push(id)
    }
  }


/* -------------------------- ODD NUMBER OF PEOPLE -------------------------- */


  // If odd number of people, remove a specified person from the list
  if (participatingIDs.length % 2 === 1) {
    log("<b>Odd number of people found.</b>");
    // iterate through sandboxers until you find the row with "yes" under "The person to remove if odd number"
    // keep iterating while (index still inside array AND NOT equal to yes AND NOT equal to 1))
    // (you have to check that the value is not blank before converting to lowercase)
    for (
      id = 0;
      id < Sandboxers.length
        && !(Sandboxers[id][COLUMN['Remove if odd']] && Sandboxers[id][COLUMN['Remove if odd']].toLowerCase() === 'yes')
        && !(Sandboxers[id][COLUMN['Remove if odd']] === '1');
      id++);
    // if there is a sandboxer selected for "The person to remove if odd number"
    // (if the loop broke because it found someone)
    if (id !== Sandboxers.length) {
      // if the selectped person to be removed for odd nums IS OPTED IN
      // if that person's ID is included in the list of valid peoples ID
      if (participatingIDs.includes(id)) {
        outputOfRemoval = '<b>"' + Sandboxers[id][COLUMN['Name']] + '"</b> was removed to make an even number of people';
        log(outputOfRemoval);
        // put massive text at top of screen showing steve was removed
        document.body.innerHTML = `<h2 class="unselectable">` + outputOfRemoval + `</h2>` + document.body.innerHTML;
        // find index of the person's ID in the VALID list of people.
        // Remove him
        participatingIDs.splice(participatingIDs.indexOf(id), 1);
      // if the selected person to be removed for odd nums HAS OPTED OUT
      } else {
        log("ERROR - The person selected for 'Remove if odd number' is not opted in. (<b>" + Sandboxers[id][COLUMN['Name']] + "</b>)");
        log("HELP - On the '<b>Participants</b>' tab, '<b>Remove if odd number</b>' column, put <b>Yes</b> for someone who has opted in");
        return;
      }
    // if no sandboxer was selected for "removal if odd number"
    // (if the loop broke because it reached the end of the array)
    } else {
      log('ERROR - No one is selected to be removed.');
      log('HELP - On the "<b>Participants</b>" table, "<b>Remove if odd number</b>" column, put <b>Yes</b> for 1 person');
      return;
    }
  }


/* ---------------------------- CALCULATE PAIRS ----------------------------- */


  log('Calculating pairs......');

  // keep shuffling pairs until there are no clashes
  for (repeats = 0, numClashes = -1; numClashes !== 0; repeats++) {
    // randomise order of list and pair people next to each other
    // list = [8,1,6,3,2,0]
    // 8-1, 6-3, 2-0 ....
    shuffledIDs = shuffle(participatingIDs);

    // how many pairings have happened before
    numClashes = 0;
    for (i = 0; i < shuffledIDs.length; i += 2) {
      // check bob's list of IDs he has paired with before
      // if that pairing already happened, it would be in the pairs section of bob's row
      // index of location must be in that "Pairs --->" section
      // or -1, which still isnt in the Pairs section
      if (Sandboxers[shuffledIDs[i]].indexOf(shuffledIDs[i+1].toString()) >= COLUMN['Pairs']) {
        numClashes++
      }
    }
  }

  console.log(repeats.toString() + ' attempts to create unique pairing');


/* ------------------------- CONVERT PAIRS TO TABLE ------------------------- */


  var table = [];
  for (i = 0; i < shuffledIDs.length; i += 2) {
    s1 = Sandboxers[shuffledIDs[i]];
    s2 = Sandboxers[shuffledIDs[i+1]];

    // Emails  Name 1  Location 1  Number 1  Scheduling Link 1 Name 2  Location 2  Number 2  Scheduling Link 2
    table.push([
      s1[COLUMN['Email']] + ", " + s2[COLUMN['Email']],
      s1[COLUMN['Name']],
      s1[COLUMN['Location']],
      s1[COLUMN['Number']],
      s1[COLUMN['Link']],
      s2[COLUMN['Name']],
      s2[COLUMN['Location']],
      s2[COLUMN['Number']],
      s2[COLUMN['Link']]
    ]);
  }


/* ------------------------- CONVERT TABLE TO HTML ------------------------- */


  headers = ['Emails','Name 1','Location 1','Number 1','Scheduling Link 1','Name 2','Location 2','Number 2','Scheduling Link 2'];
  // add headers to html table
  mailMergeTableHTML = `
  <thead>
    <tr>
      <td>` + headers.join("</td><td>") + `</td>
    </tr>
  </thead>
  <tbody>`;

  table.forEach(row => {
    mailMergeTableHTML += "<tr><td>" + row.join("</td><td>") + "</td></tr>"
  });
  mailMergeTableHTML += "</tbody>";


/* -------------- THE NEW PAIR IDS COLUMN FOR PARTICIPANTS TAB -------------- */


  // array of ints
  pairIDsTable = new Array(Sandboxers.length);
  pairIDsTable.fill("<br>");
  for (i = 0; i < shuffledIDs.length; i += 2) {
    pairIDsTable[shuffledIDs[i]] = shuffledIDs[i+1].toString()+'<br>';
    pairIDsTable[shuffledIDs[i+1]] = shuffledIDs[i].toString()+'<br>';
  }

  document.getElementById('pairsTable').innerHTML = `
  <thead class="unselectable">
    <tr>
      <td>Pairs<br></td>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td>` + pairIDsTable.join("</td></tr><tr><td>") + `</td>
    </tr>
  </tbody>`;


/* ---------------------------- HIDE/DISPLAY HTML --------------------------- */


  // hide the loading screen
  document.getElementById('loadingImage').style = 'display:none';
  // hide the div showing errors
  document.getElementById('errorBox').style = 'display:none';
  // put the table html onto the website
  document.getElementById('mailMergeTable').innerHTML = mailMergeTableHTML;
});
