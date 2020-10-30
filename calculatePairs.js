///////////// Seyhan Van Khan
///////////// The Pair Calculator
///////////// From a spreadsheet of people, find each person a new match every month
///////////// October 2020
///////////// github.com/seyhanvankhan

/* -------------------------------- CONSTANTS ------------------------------- */

// index of each column in Sandboxers array
const COLUMN = {
  "ID":             0,
  "Name":           1,
  "Email":          2,
  "Number":         3,
  "Location":       4,
  "Link":           5,
  "Comments":       6,
  "Opted Out":      7,
  "Remove if odd":  8,
  "Pairs":          9
};

// the id of the first sandboxer to signup with the form
// the first entry in 'Signup Form Responses'
const SIGNUP_RESPONSES_STARTING_ID = 238;
const PARTICIPANTS_SPREADSHEET_URL = `https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}/values/${PARTICIPANTS_RANGE}?key=${API_KEY}`;
const OPTOUT_SPREADSHEET_URL       = `https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}/values/${OPTOUT_RANGE}?key=${API_KEY}`;


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
  if (document.getElementById("loadingImage").src !== "resources/img/error.jpg")
    document.getElementById("loadingImage").src = "resources/img/error.jpg";

  console.log(output);
  document.getElementById('loadingScreen').innerHTML += output + '<br>';
}

function lowerCase(str) {
  return (str) ? str.toLowerCase() : ""
}


/* -------------------------------- BUTTONS --------------------------------- */


document.getElementById("button-mailmerge").onclick = () => {
  document.getElementById('table-mailmerge').style.display = 'block';
  document.getElementById('table-pairs').style.display = 'none';
  document.getElementById('button-mailmerge').className = 'active'
  document.getElementById('button-pairs').className = '';
};

document.getElementById("button-pairs").onclick = () => {
  document.getElementById('table-mailmerge').style.display = 'none';
  document.getElementById('table-pairs').style.display = 'block';
  document.getElementById('button-mailmerge').className = '';
  document.getElementById('button-pairs').className = 'active';
};


/* ------------------------- GET PARTICIPANTS DATA -------------------------- */


fetch(PARTICIPANTS_SPREADSHEET_URL)
.then(resp => resp.json())
.then(data => {
  // remove the first row (the headers)
  const Sandboxers = data['values'].slice(1);


/* -------------------------- CHECK FOR DUPLICATES -------------------------- */


  const currentData = (Sandboxers.map(row => [
    lowerCase(row[COLUMN['Name']]),
    lowerCase(row[COLUMN['Email']]),
    lowerCase(row[COLUMN['Number']])
  ])).flat();

  // filter the array for only duplicate, non blank values
  // (returns the 2nd+ occurance of every duplicate)
  // create a unique set of duplicate values
  theDuplicates = new Set(
    currentData.filter((item, index) => (
      currentData.indexOf(item) !== index && item !== ""
    ))
  );

  if (theDuplicates.size !== 0) {
    log("<h3>There are duplicates of people.</h3>");
    // create a unique set of all nonblank items in the list that were the 2nd occurance
    // for each duplicate
    theDuplicates.forEach(duplicateValue => {
      // each sandboxer has 3 elements in the personal info list
      // floor division of index // 3 returns the sandboxer's ID
      // + 2 to get actual row in 'Participants'
      originalRow = Math.floor(currentData.indexOf(duplicateValue) / 3) + 2;
      duplicateSandboxerID = Math.floor(currentData.lastIndexOf(duplicateValue) / 3);
      // minus 238 to get how many rows down the person is from black line
      // + 2 to get the actual row in 'Signup Form Responses'
      duplicateRow = duplicateSandboxerID - SIGNUP_RESPONSES_STARTING_ID + 2;
      log(`<b>Delete Row ${duplicateRow} in 'Signup Form Responses'</b>. It's a duplicate of Row ${originalRow} in 'Participants'`);
    });
    return;
  }


/* --------------------------- GET OPTED OUT DATA --------------------------- */


  fetch(OPTOUT_SPREADSHEET_URL)
  .then(resp => resp.json())
  .then(data => {
    // remove the first row (the headers)
    OptedOut = data['values'].slice(1);


/* ------------------------- CLEAN UP THE OPTED OUT ------------------------- */


    // list of IDs that match the values in the cells in OptedOut range
    // if it matches someone:
    //      adds their ID
    // if it doesnt match anyone's data:
    //      adds -1
    // keep 1 of every id in there (unique set)
    const optoutIDs = new Set();
    OptedOut.forEach(row => {
      if (row[1]) optoutIDs.add(Math.floor(currentData.indexOf(row[1].toLowerCase()) / 3));
      if (row[2]) optoutIDs.add(Math.floor(currentData.indexOf(row[2].toLowerCase()) / 3));
    });
    // delete the -1s. These represented the names or emails not found in any sandboxer data
    optoutIDs.delete(-1);

    optoutIDs.forEach(optoutID => {
      // check if they havent already been opted out
      if (Sandboxers[optoutID][COLUMN['Opted Out']] !== "1") {
        log(
          `<h3>
            ${Sandboxers[optoutID][COLUMN['Name']]} opted out.
          </h3>
          <b>Row ${optoutID + 2} in 'Participants':</b> Put 1 in the 'Opted Out' column`
        )
      }
    });
    if (OptedOut.length !== 0) {
      log(`<h4>Clear all rows in 'Optout Form Responses' (except for the headers)</h4>`);
      return;
    }


/* -------------------------- GET LIST OF OPTED IN -------------------------- */


    // iterate through each sandboxer, saving each sandboxer's ID if they have opted in (opted out column is BLANK)
    // array of ints
    const participatingIDs = [];
    for (id = 0; id < Sandboxers.length; id++) {
      if (Sandboxers[id][COLUMN["Opted Out"]] !== "1") {
        participatingIDs.push(id)
      }
    }


/* -------------------------- ODD NUMBER OF PEOPLE -------------------------- */


    // If odd number of people, remove a specified person from the list
    if (participatingIDs.length % 2 === 1) {
      // iterate through sandboxers until you find the row with "1" under "The person to remove if odd number"
      // keep iterating while (index still inside array AND value NOT equal to 1))
      for (
        id = 0;
        id < Sandboxers.length && Sandboxers[id][COLUMN['Remove if odd']] !== '1';
        id++);
      // if there is a sandboxer selected for "The person to remove if odd number"
      // (if the loop broke because it found someone)
      if (id !== Sandboxers.length) {
        // if person to remove if odd IS OPTED IN
        if (participatingIDs.includes(id)) {
          // put massive text at top of screen showing steve was removed
          document.getElementById("oddRemoved").innerHTML = `<b>${Sandboxers[id][COLUMN['Name']]}</b> was removed to have an even number of people`;
          // find index of the person's ID in the VALID list of people.
          // Remove him
          participatingIDs.splice(participatingIDs.indexOf(id), 1);
        // if the selected person to be removed for odd nums HAS OPTED OUT
        } else {
          log(`
            <h3>The person selected for 'Remove if odd number' has opted out. </h3>
            (<b>ID: ${id}</b>)
            <br>Choose someone that isn't opted out.
            <br>In 'Participants', put <b>1</b> under '<b>Remove if odd number</b>' column for only 1 person.`
          );
          return;
        }
      // if no sandboxer was selected for "remove if odd number"
      // (if the loop broke because it reached the end of the array)
      } else {
        log(`
          <h3>No one is selected to be removed if odd number.</h3>
          <br>In 'Participants', put <b>1</b> under '<b>Remove if odd number</b>' column for only 1 person.`
        );
        return;
      }
    }


/* ---------------------------- CALCULATE PAIRS ----------------------------- */


    // keep shuffling pairs until there are no clashes
    // if the for loop broke before the end of list bc it found a clashing pair
    do {
      // randomise order of list and pair people next to each other
      // list = [8,1,6,3,2,0]
      // 8-1, 6-3, 2-0 ....
      shuffledIDs = shuffle([...participatingIDs]);
      // check if a pairing has happened before
      // keep iterating through the pairs until you reach a non-unique pair or end of list
      // if pairing is non-unique, index of person would be in the pairs section of their row
      for (
        i = 0;
        i < shuffledIDs.length && Sandboxers[shuffledIDs[i]].indexOf(shuffledIDs[i+1].toString()) < COLUMN['Pairs'];
        i += 2);
    } while (i < shuffledIDs.length);


/* ---------------------- CONVERT PAIRS TO TABLE HTML ----------------------- */


    // html table with headers
    mailMergeTable = `
      <thead class="unselectable">
        <tr>
          <td>` + [
            'Emails',
            'Name 1',
            'Location 1',
            'Number 1',
            'Scheduling Link 1',
            'Name 2',
            'Location 2',
            'Number 2',
            'Scheduling Link 2'
          ].join("</td><td>") + `</td>
        </tr>
      </thead>
      <tbody>`;

    for (i = 0; i < shuffledIDs.length; i += 2) {
      s1 = Sandboxers[shuffledIDs[i]];
      s2 = Sandboxers[shuffledIDs[i+1]];
      // Emails  Name 1  Location 1  Number 1  Scheduling Link 1 Name 2  Location 2  Number 2  Scheduling Link 2
      mailMergeTable += `
        <tr>
          <td>` + [
            s1[COLUMN['Email']] + ", " + s2[COLUMN['Email']],
            s1[COLUMN['Name']],
            s1[COLUMN['Location']],
            s1[COLUMN['Number']],
            s1[COLUMN['Link']],
            s2[COLUMN['Name']],
            s2[COLUMN['Location']],
            s2[COLUMN['Number']],
            s2[COLUMN['Link']]
          ].join("</td><td>") + `</td>
        </tr>`
    }
    mailMergeTable += "</tbody>";


/* --------------- THE NEW PAIR IDS COLUMN FOR 'PARTICIPANTS' --------------- */


    // array of ints
    pairIDsTable = new Array(Sandboxers.length);
    pairIDsTable.fill("<br>");
    for (i = 0; i < shuffledIDs.length; i += 2) {
      pairIDsTable[shuffledIDs[i]] = `${shuffledIDs[i+1]}<br>`;
      pairIDsTable[shuffledIDs[i+1]] = `${shuffledIDs[i]}<br>`;
    }

    document.getElementById('table-pairs').innerHTML = `
      Pairs
      <tbody>
        <tr>
          <td>` + pairIDsTable.join("</td></tr><tr><td>") + `</td>
        </tr>
      </tbody>`;


/* ---------------------------- HIDE/DISPLAY HTML --------------------------- */


    // hide the loading screen (image & text)
    document.getElementById('loadingScreen').style = 'display:none';
    // put the table html onto the website
    document.getElementById('table-mailmerge').innerHTML = mailMergeTable;
  });
});
