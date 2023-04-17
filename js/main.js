//using jQuery for the file import window and to display the file(s) selected (NICK)
$(document).on("change", ".file-input", function () {
  var filesCount = $(this)[0].files.length;
  var textbox = $(this).prev();

  if (filesCount === 1) {
    var fileName = $(this).val().split("\\").pop();
    textbox.text(fileName);
  } else {
    textbox.text(filesCount + " files selected");
  }
});

// setting global variables to store array data from CSV (NICK)
//add in new constant for workshop number based on csv file
const idData = [];
const nameData = [];
const workshopNum = [];
const preferenceData = [];
const applicationData = [];
const programmingData = [];
const timezoneData = [];
const projectdayData = [];
const timeData = [];
const teamNum = [];
const resultArray = [];

// creating a get element to obtain the uploaded file from user and parse the data into a JS array to be manipluated (NICK)
const uploadConfirm = document.getElementById("uploadConfirm").addEventListener("click", () => {
  //using PapaParse to parse the CSV to an array (NICK)

  Papa.parse(document.getElementById("fileSelect").files[0], {
    download: true,
    header: true,
    skipEmptyLines: true,
    complete: function (results) {
      //looping through data and allocating the result into an array (NICK)
      for (i = 0; i < results.data.length; i++) {
        idData.push(results.data[i].id);
        nameData.push(results.data[i].sName);
        workshopNum.push(results.data[i].sWorkshop);
        preferenceData.push(results.data[i].sPreference);
        applicationData.push(results.data[i].appType);
        programmingData.push(results.data[i].progLang);
        timezoneData.push(results.data[i].timeZone);
        projectdayData.push(results.data[i].projDay);
        timeData.push(results.data[i].dayNight);
      }
      //temporary console.log to see if the data is being stored in the ideal manner (NICK)
      console.log(idData);
      console.log(nameData);
      console.log(workshopNum);
      console.log(preferenceData);
      console.log(applicationData);
      console.log(programmingData);
      console.log(timezoneData);
      console.log(projectdayData);
      console.log(timeData);

      //function for reading the rows in the csv and displaying an error if there is more than 421 rows (SIBEL & NICK)
      if (idData.length >= 422) {
        alert("This file exceeds the row limit (421)");
      } else {
        alert("File has been sucessfully uploaded!");
      }
    },
  });
});

// Uploading CSV though the use of PAPA PARSE (MATT)

// Assigning an event lister to the Submit button (MATT)
let btnUploadCsv = document.getElementById("btnUploadCsv").addEventListener("click", () => {
  //Testing for success (MATT)
  console.log("the button is clicked");
  Papa.parse(document.getElementById("fileSelect").files[0], {
    download: true,
    header: false,
    // parsing the results from CSV (MATT)
    complete: function (results) {
      console.log(results);
      let i = 0;
      results.data.map((data, index) => {
        if (i === 0) {
          let table = document.getElementById("tblData");
          generateTableHead(table, data);
        } else {
          let table = document.getElementById("tblData");
          generateTableRows(table, data);
        }
        i++;
        console.log(data);
      });
    },
  });
});

// Uploading a XLSX file (TANISH)
function UploadProcess() {
  //Reference the FileUpload element. (TANISH)
  var fileUpload = document.getElementById("fileSelect");

  //function for reading the rows in the csv and displaying an error if there is more than 421 rows (SIBEL & NICK)
  if (idData.length >= 422) {
    alert("This file exceeds the row limit (421)");
  } else {
    alert("File has been sucessfully uploaded!");
  }

  //Validate whether File is valid Excel file. (TANISH)
  var regex = /^([a-zA-Z0-9\s_\\.\-:])+(.xls|.xlsx)$/;
  if (regex.test(fileUpload.value.toLowerCase())) {
    if (typeof (FileReader) != "undefined") {
      var reader = new FileReader();
      var is_60 = (fileUpload.value.split('\\')[2] == "Team-allocation-preferences-60.xlsx")

      //For Browsers other than IE (TANISH)
      if (reader.readAsBinaryString) {
        reader.onload = function (e) {
          let table = document.getElementById("tblData");
          GetTableFromExcel(e.target.result, is_60);
        };
        reader.readAsBinaryString(fileUpload.files[0]);
      } else {
        reader.onload = function (e) {
          var data = "";
          var bytes = new Uint8Array(e.target.result);
          for (var i = 0; i < bytes.byteLength; i++) {
            data += String.fromCharCode(bytes[i]);
          }
          GetTableFromExcel(data, is_60);
        };
        reader.readAsArrayBuffer(fileUpload.files[0]);
      }
    } else {
      alert("This browser does not support HTML5.");
    }
  } else {
    alert("Please upload a valid Excel file.");
  }
};

// Creating table heads to the data table div (MATT)
function generateTableHead(table, data) {
  let thead = table.createTHead();
  let row = thead.insertRow();
  for (let key of data) {
    let th = document.createElement("th");
    let text = document.createTextNode(key);
    th.appendChild(text);
    row.appendChild(th);
  }
}

// Creating table heads to the data table div (MATT)
function generateTableRows(table, data) {
  let newRow = table.insertRow(-1);
  data.map((row, index) => {
    let newCell = newRow.insertCell();
    let newText = document.createTextNode(row);
    newCell.appendChild(newText);
  });
}

// Clear table function (MATT)
let clearTable = document.getElementById("clearTable").addEventListener("click", () => {
  console.log("the button is clicked");
  this.tblData.innerHTML = "";
});


// Function to paste data to table and store value in array (TANISH AND NICK)
function GetTableFromExcel(data, is_60) {

  //Read the Excel File data in binary (TANISH)
  var workbook = XLSX.read(data, {
    type: 'binary'
  });

  //get the name of First Sheet (TANISH)
  var Sheet = workbook.SheetNames[0];

  //Read all rows from First Sheet into an JSON array (TANISH)
  var excelRows = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[Sheet]);

  //Create a HTML Table element (TANISH)
  var myTable = document.createElement("table");
  myTable.border = "1";

  //Add the header row (TANISH)
  var row = myTable.insertRow(-1);
  if (is_60) {
    headerCell = document.createElement("TH");
    headerCell.innerHTML = "ID";
    row.appendChild(headerCell);

    headerCell = document.createElement("TH");
    headerCell.innerHTML = "Start time";
    row.appendChild(headerCell);

    headerCell = document.createElement("TH");
    headerCell.innerHTML = "Completion time";
    row.appendChild(headerCell);

    headerCell = document.createElement("TH");
    headerCell.innerHTML = "Email";
    row.appendChild(headerCell);

    headerCell = document.createElement("TH");
    headerCell.innerHTML = "Name";
    row.appendChild(headerCell);

    headerCell = document.createElement("TH");
    headerCell.innerHTML = "Your name  (FirstName LastName)";
    row.appendChild(headerCell);

    headerCell = document.createElement("TH");
    headerCell.innerHTML = "Your student ID";
    row.appendChild(headerCell);

    headerCell = document.createElement("TH");
    headerCell.innerHTML = "Your workshop class";
    row.appendChild(headerCell);

    headerCell = document.createElement("TH");
    headerCell.innerHTML = "If you prefer to be in a team with particular students, list their student IDs separated by commas (max. 6 student IDs):";
    row.appendChild(headerCell);

    headerCell = document.createElement("TH");
    headerCell.innerHTML = "Project Preference";
    row.appendChild(headerCell);

    headerCell = document.createElement("TH");
    headerCell.innerHTML = "TechStack";
    row.appendChild(headerCell);

    headerCell = document.createElement("TH");
    headerCell.innerHTML = "timeZone";
    row.appendChild(headerCell);

    headerCell = document.createElement("TH");
    headerCell.innerHTML = "team_work";
    row.appendChild(headerCell);

    //Add the data rows from Excel file (TANISH)
    for (var i = 0; i < excelRows.length; i++) {
      //Add the data row (TANISH)
      var row = myTable.insertRow(-1);

      //Add the data cells (TANISH)
      var cell = row.insertCell(-1);
      cell.innerHTML = excelRows[i].ID;

      cell = row.insertCell(-1);
      cell.innerHTML = excelRows[i].start_time;

      cell = row.insertCell(-1);
      cell.innerHTML = excelRows[i].completion_time;

      cell = row.insertCell(-1);
      cell.innerHTML = excelRows[i].email;

      cell = row.insertCell(-1);
      cell.innerHTML = excelRows[i].name;

      cell = row.insertCell(-1);
      cell.innerHTML = excelRows[i].first_name_last_name;

      cell = row.insertCell(-1);
      cell.innerHTML = excelRows[i].sWorkshop;

      cell = row.insertCell(-1);
      cell.innerHTML = excelRows[i].sttudent_id;

      cell = row.insertCell(-1);
      cell.innerHTML = excelRows[i].class;

      cell = row.insertCell(-1);
      cell.innerHTML = excelRows[i].Max_6;

      cell = row.insertCell(-1);
      cell.innerHTML = excelRows[i].project_preference;

      cell = row.insertCell(-1);
      cell.innerHTML = excelRows[i].tech;

      cell = row.insertCell(-1);
      cell.innerHTML = excelRows[i].timezone;

      cell = row.insertCell(-1);
      cell.innerHTML = excelRows[i].team_work;

    }

  } else {
    // creating header names (TANISH)
    headerCell = document.createElement("TH");
    headerCell.innerHTML = "id";
    row.appendChild(headerCell);

    headerCell = document.createElement("TH");
    headerCell.innerHTML = "sName";
    row.appendChild(headerCell);

    headerCell = document.createElement("TH");
    headerCell.innerHTML = "sWorkshop";
    row.appendChild(headerCell);

    headerCell = document.createElement("TH");
    headerCell.innerHTML = "sPreference";
    row.appendChild(headerCell);

    headerCell = document.createElement("TH");
    headerCell.innerHTML = "appType";
    row.appendChild(headerCell);

    headerCell = document.createElement("TH");
    headerCell.innerHTML = "progLang";
    row.appendChild(headerCell);

    headerCell = document.createElement("TH");
    headerCell.innerHTML = "timeZone";
    row.appendChild(headerCell);

    headerCell = document.createElement("TH");
    headerCell.innerHTML = "projDay";
    row.appendChild(headerCell);

    headerCell = document.createElement("TH");
    headerCell.innerHTML = "dayNight";
    row.appendChild(headerCell);


    //Add the data rows from Excel file (TANISH)
    for (var i = 0; i < excelRows.length; i++) {
      //Add the data row (TANISH)
      var row = myTable.insertRow(-1);

      //Add the data cells (TANISH)
      var cell = row.insertCell(-1);
      cell.innerHTML = excelRows[i].id;

      cell = row.insertCell(-1);
      cell.innerHTML = excelRows[i].sName;

      cell = row.insertCell(-1);
      cell.innerHTML = excelRows[i].sWorkshop;

      cell = row.insertCell(-1);
      cell.innerHTML = excelRows[i].sPreference;

      cell = row.insertCell(-1);
      cell.innerHTML = excelRows[i].appType;

      cell = row.insertCell(-1);
      cell.innerHTML = excelRows[i].progLang;

      cell = row.insertCell(-1);
      cell.innerHTML = excelRows[i].timeZone;

      cell = row.insertCell(-1);
      cell.innerHTML = excelRows[i].projDay;

      cell = row.insertCell(-1);
      cell.innerHTML = excelRows[i].dayNight;

      // storing all the parsed data into the global array's (NICK)
      idData.push(excelRows[i].id);
      nameData.push(excelRows[i].sName);
      workshopNum.push(excelRows[i].sWorkshop);
      preferenceData.push(excelRows[i].sPreference);
      applicationData.push(excelRows[i].appType);
      programmingData.push(excelRows[i].progLang);
      timezoneData.push(excelRows[i].timeZone);
      projectdayData.push(excelRows[i].projDay);
      timeData.push(excelRows[i].dayNight);
    }
  }
  var ExcelTable = document.getElementById("tblData");
  ExcelTable.innerHTML = "";
  ExcelTable.appendChild(myTable);

  // logging the variables to ensure successful storage of arrays (NICK)
  console.log(idData);
  console.log(nameData);
  console.log(workshopNum);
  console.log(preferenceData);
  console.log(applicationData);
  console.log(programmingData);
  console.log(timezoneData);
  console.log(projectdayData);
  console.log(timeData);
};

// creating global variables for teams (Used for the basicAllocation() function) (SIBEL AND NICK)
const teams1 = [];
const teams2 = [];
const teams3 = [];
const teams4 = [];
const teams5 = [];
const teams6 = [];
const teams7 = [];
const teams8 = [];
const teams9 = [];
const teams10 = [];
const teams11 = [];
const teams12 = [];
const teams13 = [];
const teams14 = [];
const teams15 = [];
const teams16 = [];
const teams17 = [];
const teams18 = [];
const teams19 = [];
const teams20 = [];
const teams21 = [];
const teams22 = [];
const teams23 = [];
const teams24 = [];
const teams25 = [];
const teams26 = [];
const teams27 = [];
const teams28 = [];
const teams29 = [];
const teams30 = [];
const teams31 = [];
const teams32 = [];
const teams33 = [];
const teams34 = [];
const teams35 = [];
const teams36 = [];

// Basic allocation function (SIBEL (changed from Sandras SPRINT 1 CODE)) (SIBEL)

function basicAllocation() {

  alert("Allocation is successful, procced to export");

  //new chunk size allocation code, makes results arrays (SIBEL)
  var size = document.getElementById("size").value;
  const perChunk = size // items per chunk  (SIBEL)  
  //grabbing from constant array idData (SIBEL)

  const result = idData.reduce((resultArray, item, index) => {
    const chunkIndex = Math.floor(index / perChunk)

    if (!resultArray[chunkIndex]) {
      resultArray[chunkIndex] = [] // start a new chunk (SIBEL)
    }

    resultArray[chunkIndex].push(item)

    return resultArray
  }, [])

  console.log(result); // result: [ [# of items], [# of items], [left overs] ]


  //working preference allocation (SIBEL)

  for (let i = 0; i < idData.length; i++) {
    if (workshopNum[i] == 1 && applicationData[i] == 'Web' && programmingData[i] == 'HTML') { // 1 - web - html
      teams1.push(idData[i] + " " + workshopNum[i] + " " + applicationData[i] + " " + programmingData[i]);

    } else if (workshopNum[i] == 1 && applicationData[i] == 'Desktop' && programmingData[i] == 'Java') { // 1 - desktop - java
      teams2.push(idData[i] + " " + workshopNum[i] + " " + applicationData[i] + " " + programmingData[i]);

    } else if (workshopNum[i] == 1 && applicationData[i] == 'Desktop' && programmingData[i] == 'Python') { // 1 - desktop - python
      teams3.push(idData[i] + " " + workshopNum[i] + " " + applicationData[i] + " " + programmingData[i]);

    } else if (workshopNum[i] == 1 && applicationData[i] == 'iOS' && programmingData[i] == 'Javascript') { // 1 - iOS - javascript
      teams4.push(idData[i] + " " + workshopNum[i] + " " + applicationData[i] + " " + programmingData[i]);

    } else if (workshopNum[i] == 1 && applicationData[i] == 'iOS' && programmingData[i] == 'Python') { // 1 - iOS - python
      teams5.push(idData[i] + " " + workshopNum[i] + " " + applicationData[i] + " " + programmingData[i]);

    } else if (workshopNum[i] == 2 && applicationData[i] == 'Web' && programmingData[i] == 'HTML') { // 2 - web - html
      teams6.push(idData[i] + " " + workshopNum[i] + " " + applicationData[i] + " " + programmingData[i]);

    } else if (workshopNum[i] == 2 && applicationData[i] == 'Desktop' && programmingData[i] == 'Java') { // 2 - desktop - java
      teams7.push(idData[i] + " " + workshopNum[i] + " " + applicationData[i] + " " + programmingData[i]);

    } else if (workshopNum[i] == 2 && applicationData[i] == 'Desktop' && programmingData[i] == 'Python') { // 2 - desktop - python
      teams8.push(idData[i] + " " + workshopNum[i] + " " + applicationData[i] + " " + programmingData[i]);

    } else if (workshopNum[i] == 2 && applicationData[i] == 'iOS' && programmingData[i] == 'Javascript') { // 2 - iOS - javascript
      teams9.push(idData[i] + " " + workshopNum[i] + " " + applicationData[i] + " " + programmingData[i]);

    } else if (workshopNum[i] == 2 && applicationData[i] == 'Desktop' && programmingData[i] == 'Python') { // 2 - iOS - python
      teams10.push(idData[i] + " " + workshopNum[i] + " " + applicationData[i] + " " + programmingData[i]);

    } else if (workshopNum[i] == 3 && applicationData[i] == 'Web' && programmingData[i] == 'HTML') { // 3 - web - html
      teams11.push(idData[i] + " " + workshopNum[i] + " " + applicationData[i] + " " + programmingData[i]);

    } else if (workshopNum[i] == 3 && applicationData[i] == 'Desktop' && programmingData[i] == 'Java') { // 3 - desktop - java
      teams12.push(idData[i] + " " + workshopNum[i] + " " + applicationData[i] + " " + programmingData[i]);

    } else if (workshopNum[i] == 3 && applicationData[i] == 'Desktop' && programmingData[i] == 'Python') { // 3 - desktop - python
      teams13.push(idData[i] + " " + workshopNum[i] + " " + applicationData[i] + " " + programmingData[i]);

    } else if (workshopNum[i] == 3 && applicationData[i] == 'iOS' && programmingData[i] == 'Javascript') { // 3 - iOS - javascript
      teams14.push(idData[i] + " " + workshopNum[i] + " " + applicationData[i] + " " + programmingData[i]);

    } else if (workshopNum[i] == 3 && applicationData[i] == 'iOS' && programmingData[i] == 'Python') { // 3 - iOS - python
      teams15.push(idData[i] + " " + workshopNum[i] + " " + applicationData[i] + " " + programmingData[i]);

    } else if (workshopNum[i] == 1 && applicationData[i] == '' && programmingData[i] == 'Python') {
      teams16.push(idData[i] + " " + workshopNum[i] + " " + programmingData[i]); //wrk 1 - python programming

    } else if (workshopNum[i] == 1 && applicationData[i] == '' && programmingData[i] == 'Javascript') {
      teams17.push(idData[i] + " " + workshopNum[i] + " " + programmingData[i]); //wrk 1 - javascript programming

    } else if (workshopNum[i] == 1 && applicationData[i] == '' && programmingData[i] == 'Java') {
      teams18.push(idData[i] + " " + workshopNum[i] + " " + programmingData[i]); //wrk 1 - Java programming

    } else if (workshopNum[i] == 1 && applicationData[i] == '' && programmingData[i] == 'HTML') {
      teams19.push(idData[i] + " " + workshopNum[i] + " " + programmingData[i]); //wrk 1 - HTML programming

    } else if (workshopNum[i] == 2 && applicationData[i] == '' && programmingData[i] == 'Python') {
      teams20.push(idData[i] + " " + workshopNum[i] + " " + programmingData[i]); //wrk 2 - python programming

    } else if (workshopNum[i] == 2 && applicationData[i] == '' && programmingData[i] == 'Javascript') {
      teams21.push(idData[i] + " " + workshopNum[i] + " " + programmingData[i]); //wrk 2 - javascript programming

    } else if (workshopNum[i] == 2 && applicationData[i] == '' && programmingData[i] == 'Java') {
      teams22.push(idData[i] + " " + workshopNum[i] + " " + programmingData[i]); //wrk 2 - java programming

    } else if (workshopNum[i] == 2 && applicationData[i] == '' && programmingData[i] == 'HTML') {
      teams23.push(idData[i] + " " + workshopNum[i] + " " + programmingData[i]); //wrk 2 - HTML programming

    } else if (workshopNum[i] == 3 && applicationData[i] == '' && programmingData[i] == 'Python') {
      teams24.push(idData[i] + " " + workshopNum[i] + " " + programmingData[i]); //wrk 3 - python programming

    } else if (workshopNum[i] == 3 && applicationData[i] == '' && programmingData[i] == 'Javascript') {
      teams25.push(idData[i] + " " + workshopNum[i] + " " + programmingData[i]); //wrk 3 - javascript programming

    } else if (workshopNum[i] == 3 && applicationData[i] == '' && programmingData[i] == 'Java') {
      teams26.push(idData[i] + " " + workshopNum[i] + " " + programmingData[i]); //wrk 3 - java programming

    } else if (workshopNum[i] == 3 && applicationData[i] == '' && programmingData[i] == 'HTML') {
      teams27.push(idData[i] + " " + workshopNum[i] + " " + programmingData[i]); //wrk 3 - HTML programming

    } else if (workshopNum[i] == 1 && applicationData[i] == 'Web' && programmingData[i] == '') {
      teams28.push(idData[i] + " " + workshopNum[i] + " " + applicationData[i]); //wrk 1 - Web app

    } else if (workshopNum[i] == 1 && applicationData[i] == 'Desktop' && programmingData[i] == '') {
      teams29.push(idData[i] + " " + workshopNum[i] + " " + applicationData[i]); //wrk 1 - Desktop app

    } else if (workshopNum[i] == 1 && applicationData[i] == 'iOS' && programmingData[i] == '') {
      teams30.push(idData[i] + " " + workshopNum[i] + " " + applicationData[i]); //wrk 1 - iOS app

    } else if (workshopNum[i] == 2 && applicationData[i] == 'Web' && programmingData[i] == '') {
      teams31.push(idData[i] + " " + workshopNum[i] + " " + applicationData[i]); //wrk 2 - Web app

    } else if (workshopNum[i] == 2 && applicationData[i] == 'Desktop' && programmingData[i] == '') {
      teams32.push(idData[i] + " " + workshopNum[i] + " " + applicationData[i]); //wrk 2 - Desktop app

    } else if (workshopNum[i] == 2 && applicationData[i] == 'iOS' && programmingData[i] == '') {
      teams33.push(idData[i] + " " + workshopNum[i] + " " + applicationData[i]); //wrk 2 - iOS app

    } else if (workshopNum[i] == 3 && applicationData[i] == 'Web' && programmingData[i] == '') {
      teams34.push(idData[i] + " " + workshopNum[i] + " " + applicationData[i]); //wrk 3 - Web app

    } else if (workshopNum[i] == 3 && applicationData[i] == 'Desktop' && programmingData[i] == '') {
      teams35.push(idData[i] + " " + workshopNum[i] + " " + applicationData[i]); //wrk 3 - Desktop app

    } else if (workshopNum[i] == 3 && applicationData[i] == 'iOS' && programmingData[i] == '') {
      teams36.push(idData[i] + " " + workshopNum[i] + " " + applicationData[i]); //wrk 3 - iOS app
    }
  }

  // Logging arrays to ensure success (NICK)
  console.log(teams1, teams2, teams3, teams4, teams5, teams6, teams7, teams8,
    teams9, teams10, teams11, teams12, teams13, teams14, teams15, teams16,
    teams17, teams18, teams19, teams20, teams21, teams22, teams23, teams24,
    teams25, teams26, teams27, teams28, teams29, teams30, teams31, teams32,
    teams33, teams34, teams35, teams36);
}

// CSV export function (NICK)
function export2csv() {
  let data = "";
  const tableData = [];
  // creating data table headers (NICK)
  const rows = [
    ['Student ID', 'Student Name', 'Workshop Time', 'ID Preference', 'Application Preference', 'Programming Preference', 'Time Zone', 'Project Day', 'Day / Night']
  ];
  // using a for loop to push all data for each student to be printed into the exported CSV in the correct format (NICK)
  for (var i = 0; i < idData.length; i++) {
    rows.push([
      [idData[i], nameData[i], workshopNum[i], preferenceData[i], applicationData[i], programmingData[i], timezoneData[i], projectdayData[i], timeData[i]]
    ]);
  }

  // hard-coded push variables due to large number of teams (NICK)

  rows.push('');
  rows.push([
    ['Team 1', teams1]
  ]);
  rows.push([
    ['Team 2', teams2]
  ]);
  rows.push([
    ['Team 3', teams3]
  ]);
  rows.push([
    ['Team 4', teams4]
  ]);
  rows.push([
    ['Team 5', teams5]
  ]);
  rows.push([
    ['Team 6', teams6]
  ]);
  rows.push([
    ['Team 7', teams7]
  ]);
  rows.push([
    ['Team 8', teams8]
  ]);
  rows.push([
    ['Team 9', teams9]
  ]);
  rows.push([
    ['Team 10', teams10]
  ]);
  rows.push([
    ['Team 11', teams11]
  ]);
  rows.push([
    ['Team 12', teams12]
  ]);
  rows.push([
    ['Team 13', teams13]
  ]);
  rows.push([
    ['Team 14', teams14]
  ]);
  rows.push([
    ['Team 15', teams15]
  ]);
  rows.push([
    ['Team 16', teams16]
  ]);
  rows.push([
    ['Team 17', teams17]
  ]);
  rows.push([
    ['Team 18', teams18]
  ]);
  rows.push([
    ['Team 19', teams19]
  ]);
  rows.push([
    ['Team 20', teams20]
  ]);
  rows.push([
    ['Team 21', teams21]
  ]);
  rows.push([
    ['Team 22', teams22]
  ]);
  rows.push([
    ['Team 23', teams23]
  ]);
  rows.push([
    ['Team 24', teams24]
  ]);
  rows.push([
    ['Team 25', teams25]
  ]);
  rows.push([
    ['Team 26', teams26]
  ]);
  rows.push([
    ['Team 27', teams27]
  ]);
  rows.push([
    ['Team 28', teams28]
  ]);
  rows.push([
    ['Team 29', teams29]
  ]);
  rows.push([
    ['Team 30', teams30]
  ]);
  rows.push([
    ['Team 31', teams31]
  ]);
  rows.push([
    ['Team 32', teams32]
  ]);
  rows.push([
    ['Team 33', teams33]
  ]);
  rows.push([
    ['Team 34', teams34]
  ]);
  rows.push([
    ['Team 35', teams35]
  ]);
  rows.push([
    ['Team 36', teams36]
  ]);

  // for loop to print individual data into CSV format (NICK)
  for (const row of rows) {
    const rowData = [];
    for (const column of row) {
      rowData.push(column);
    }
    tableData.push(rowData.join(","));
  }
  data += tableData.join("\n");
  const a = document.createElement("a");
  a.href = URL.createObjectURL(new Blob([data], {
    type: "text/csv"
  }));
  // declaring the name of the file on export (NICK)
  a.setAttribute("download", "TeamAllocation.csv");
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  return rows;
}

// Exporting XLSX function (NICK)
function onXLSX() {
  var wb = XLSX.utils.book_new();
  // data for the XLSX document (NICK)
  wb.Props = {
    Title: "Team Allocation",
    Subject: "SEPM",
    Author: "B1-2",
    CreatedDate: new Date(2017, 12, 19)
  };

  wb.SheetNames.push("Team Allocation");
  // creating data table headers (NICK)
  const ws_data = [
    ['Student ID', 'Student Name', 'Workshop Number', 'ID Preference', 'Application Preference', 'Programming Preference', 'Time Zone', 'Project Day', 'Day / Night']
  ];
  // using a for loop to push all data for each student to be printed into the exported CSV in the correct format (NICK)
  // different method of pushing data as XLSX is different than CSV methodology (NICK)
  for (var i = 0; i < idData.length; i++) {
    ws_data.push([
      idData[i], nameData[i], workshopNum[i], preferenceData[i], applicationData[i], programmingData[i], timezoneData[i], projectdayData[i], timeData[i]
    ]);
  };

  // adding empty push to create gap in array for team allocation (NICK)
  ws_data.push([
    [
      ['']
    ]
  ]);
  ws_data.push([
    [
      ['Team Data']
    ]
  ]);

  // hard-coded push variables due to large number of teams (NICK)
  // different method of pushing data as XLSX is different than CSV methodology (NICK)

  ws_data.push([
    [
      ['Team 1', teams1[0], teams1[1], teams1[3], teams1[4], teams1[5], teams1[6], teams1[7]]
    ]
  ]);
  ws_data.push([
    [
      ['Team 2', teams2[0], teams2[1], teams2[3], teams2[4], teams2[5], teams2[6], teams2[7]]
    ]
  ]);
  ws_data.push([
    [
      ['Team 3', teams3[0], teams3[1], teams3[3], teams3[4], teams3[5], teams3[6], teams3[7]]
    ]
  ]);
  ws_data.push([
    [
      ['Team 4', teams4[0], teams4[1], teams4[3], teams4[4], teams4[5], teams4[6], teams4[7]]
    ]
  ]);
  ws_data.push([
    [
      ['Team 5', teams5[0], teams5[1], teams5[3], teams5[4], teams5[5], teams5[6], teams5[7]]
    ]
  ]);
  ws_data.push([
    [
      ['Team 6', teams6[0], teams6[1], teams6[3], teams6[4], teams6[5], teams6[6], teams6[7]]
    ]
  ]);
  ws_data.push([
    [
      ['Team 7', teams7[0], teams7[1], teams7[3], teams7[4], teams7[5], teams7[6], teams7[7]]
    ]
  ]);
  ws_data.push([
    [
      ['Team 8', teams8[0], teams8[1], teams8[3], teams8[4], teams8[5], teams8[6], teams8[7]]
    ]
  ]);
  ws_data.push([
    [
      ['Team 9', teams9[0], teams9[1], teams9[3], teams9[4], teams9[5], teams9[6], teams9[7]]
    ]
  ]);
  ws_data.push([
    [
      ['Team 10', teams10[0], teams10[1], teams10[3], teams10[4], teams10[5], teams10[6], teams10[7]]
    ]
  ]);
  ws_data.push([
    [
      ['Team 11', teams11[0], teams11[1], teams11[3], teams11[4], teams11[5], teams11[6], teams11[7]]
    ]
  ]);
  ws_data.push([
    [
      ['Team 12', teams12[0], teams12[1], teams12[3], teams12[4], teams12[5], teams12[6], teams12[7]]
    ]
  ]);
  ws_data.push([
    [
      ['Team 13', teams13[0], teams13[1], teams13[3], teams13[4], teams13[5], teams13[6], teams13[7]]
    ]
  ]);
  ws_data.push([
    [
      ['Team 14', teams14[0], teams14[1], teams14[3], teams14[4], teams14[5], teams14[6], teams14[7]]
    ]
  ]);
  ws_data.push([
    [
      ['Team 15', teams15[0], teams15[1], teams15[3], teams15[4], teams15[5], teams15[6], teams15[7]]
    ]
  ]);
  ws_data.push([
    [
      ['Team 16', teams16[0], teams16[1], teams16[3], teams16[4], teams16[5], teams16[6], teams16[7]]
    ]
  ]);
  ws_data.push([
    [
      ['Team 17', teams17[0], teams17[1], teams17[3], teams17[4], teams17[5], teams17[6], teams17[7]]
    ]
  ]);
  ws_data.push([
    [
      ['Team 18', teams18[0], teams18[1], teams18[3], teams18[4], teams18[5], teams18[6], teams18[7]]
    ]
  ]);
  ws_data.push([
    [
      ['Team 19', teams19[0], teams19[1], teams19[3], teams19[4], teams19[5], teams19[6], teams19[7]]
    ]
  ]);
  ws_data.push([
    [
      ['Team 20', teams20[0], teams20[1], teams20[3], teams20[4], teams20[5], teams20[6], teams20[7]]
    ]
  ]);
  ws_data.push([
    [
      ['Team 21', teams21[0], teams21[1], teams21[3], teams21[4], teams21[5], teams21[6], teams21[7]]
    ]
  ]);
  ws_data.push([
    [
      ['Team 22', teams22[0], teams22[1], teams22[3], teams22[4], teams22[5], teams22[6], teams22[7]]
    ]
  ]);
  ws_data.push([
    [
      ['Team 23', teams23[0], teams23[1], teams23[3], teams23[4], teams23[5], teams23[6], teams23[7]]
    ]
  ]);
  ws_data.push([
    [
      ['Team 24', teams24[0], teams24[1], teams24[3], teams24[4], teams24[5], teams24[6], teams24[7]]
    ]
  ]);
  ws_data.push([
    [
      ['Team 25', teams25[0], teams25[1], teams25[3], teams25[4], teams25[5], teams25[6], teams25[7]]
    ]
  ]);
  ws_data.push([
    [
      ['Team 26', teams26[0], teams26[1], teams26[3], teams26[4], teams26[5], teams26[6], teams26[7]]
    ]
  ]);
  ws_data.push([
    [
      ['Team 27', teams27[0], teams27[1], teams27[3], teams27[4], teams27[5], teams27[6], teams27[7]]
    ]
  ]);
  ws_data.push([
    [
      ['Team 28', teams28[0], teams28[1], teams28[3], teams28[4], teams28[5], teams28[6], teams28[7]]
    ]
  ]);
  ws_data.push([
    [
      ['Team 29', teams29[0], teams29[1], teams29[3], teams29[4], teams29[5], teams29[6], teams29[7]]
    ]
  ]);
  ws_data.push([
    [
      ['Team 30', teams30[0], teams30[1], teams30[3], teams30[4], teams30[5], teams30[6], teams30[7]]
    ]
  ]);
  ws_data.push([
    [
      ['Team 31', teams31[0], teams31[1], teams31[3], teams31[4], teams31[5], teams31[6], teams31[7]]
    ]
  ]);
  ws_data.push([
    [
      ['Team 32', teams32[0], teams32[1], teams32[3], teams32[4], teams32[5], teams32[6], teams32[7]]
    ]
  ]);
  ws_data.push([
    [
      ['Team 33', teams33[0], teams33[1], teams33[3], teams33[4], teams33[5], teams33[6], teams33[7]]
    ]
  ]);
  ws_data.push([
    [
      ['Team 34', teams34[0], teams34[1], teams34[3], teams34[4], teams34[5], teams34[6], teams34[7]]
    ]
  ]);
  ws_data.push([
    [
      ['Team 35', teams35[0], teams35[1], teams35[3], teams35[4], teams35[5], teams35[6], teams35[7]]
    ]
  ]);
  ws_data.push([
    [
      ['Team 36', teams36[0], teams36[1], teams36[3], teams36[4], teams36[5], teams36[6], teams36[7]]
    ]
  ]);

  var ws = XLSX.utils.aoa_to_sheet(ws_data);
  wb.Sheets["Team Allocation"] = ws;

  // ensuring that the variable is in 'binary' trype as XLSX is in binary format (NICK)
  var wbout = XLSX.write(wb, {
    bookType: 'xlsx',
    type: 'binary'
  });

  function s2ab(s) {

    var buf = new ArrayBuffer(s.length);
    var view = new Uint8Array(buf);
    for (var i = 0; i < s.length; i++) view[i] = s.charCodeAt(i) & 0xFF;
    return buf;

  }
  // event lister for onclick to download the CSV file (NICK)
  $("#button-a").click(function () {
    console.log(ws_data)
    saveAs(new Blob([s2ab(wbout)], {
      type: "application/octet-stream"
    }), 'Team-Alocation.xlsx');
  });
}