
/**
 * Exempel code copied from https://code-boxx.com/javascript-excel-html-table/ for test and reference.
 * Combined code from https://docs.sheetjs.com/docs/getting-started/installation/standalone
 * https://www.webslesson.info/2021/07/how-to-display-excel-data-in-html-table.html
 * https://web.dev/read-files/#read-content
 * Uses the library provided by https://sheetjs.com/
**/ 
let fileInput = document.getElementById('template-file-selector');
let uploadBtn = document.getElementById('template-upload');
let output = document.getElementById('template-file');

let selectedFile;

fileInput.addEventListener('change', (event) => {
  // set the selectedFile variable to the chosen file
  selectedFile = event.target.files[0];
});

uploadBtn.addEventListener('click', () => {
  if (!selectedFile) {
    alert('Please select a file first!');
    return;
  }

  let reader = new FileReader();
  
  reader.onload = (event) => {
    let data = new Uint8Array(event.target.result);
    let workbook = XLSX.read(data, {type: 'array'});
    let sheet = workbook.Sheets[workbook.SheetNames[0]];
    let json = XLSX.utils.sheet_to_json(sheet);
    let table = document.createElement('table')
    table.setAttribute('id', 'templateTable');
    table.setAttribute('style','border:1px solid black;')

    // create header row
    let header = table.createTHead().insertRow();
    Object.keys(json[0]).forEach(key => {
      let cell = header.insertCell();
      cell.textContent = key;
    });

    // create data rows
    let tbody = document.createElement('tbody');
    for (let i = 0; i < json.length; i++) {
      let rowData = json[i];
      let row = tbody.insertRow();
      Object.values(rowData).forEach(value => {
        let cell = row.insertCell();
        cell.textContent = value;
        cell.setAttribute('style','border:1px solid black;')
      });
    }

    // add table to output div
    table.appendChild(tbody);
    output.innerHTML = '';
    output.appendChild(table);
  };
  
  reader.readAsArrayBuffer(selectedFile);

  // Call function to apply eventlistener with a set time-out of 1sec to make sure the table is complete first
  // Credit to https://stackoverflow.com/questions/5990725/how-to-delay-execution-in-between-the-following-in-my-javascript : Credit to T.J. Crowder
  setTimeout(function() {
    tableHandler();
  }, 1000); 
});

let tableArray =  [];

// Function to set eventlistener on all tabledata and retrieve the clicked cell from ecent object on uploaded excelfile
// Code copied from https://stackoverflow.com/questions/62259233/javascript-get-table-cell-content-on-click : credit to Teemu!
function tableHandler() {
  const tbody = document.querySelector('#templateTable tbody');
  tbody.addEventListener('click', function (e) {
    const cell = e.target.closest('td');
    if (!cell) {return;} // Quit, not clicked on a cell
    const row = cell.parentElement;
    console.log(cell.innerHTML, row.rowIndex, cell.cellIndex);
    let cellID = row.rowIndex.toString()*3 + cell.cellIndex.toString();
    let listID = row.rowIndex.toString()*3 + cell.cellIndex.toString() + 'list';
    let listInputID = row.rowIndex.toString()*3 + cell.cellIndex.toString() + 'input';
    // Changing color on selected or deselected cells and creating array with selected cells as objects.
    if (cell.style.backgroundColor !== "lightgreen") {
      cell.style.border = "2px solid green";
      cell.style.backgroundColor = "lightgreen";
      let tableObject = {cellvalue: cell.innerHTML, rowindex: row.rowIndex, 
          cellindex: cell.cellIndex, cellID: cellID};
      console.log(tableObject.cellID)
      tableArray.push(tableObject);
      console.log(tableArray);
        console.log(listID); 
        console.log(listInputID); 
      //Function to display cell value choises
      displayTableSelections(cell.innerHTML, cellID, listID, listInputID);
      // If cell is deselected make sure to remove object from array and change back color.
    } else if (cell.style.backgroundColor === "lightgreen") {
      cell.style.border = "1px solid black";
      cell.style.backgroundColor = "white";
      // Find the deselected cell and remove it from array. Credit to Peter Mortensen on https://stackoverflow.com/questions/21659888/find-and-remove-objects-in-an-array-based-on-a-key-value-in-javascript
      let cellIDs = cellID;
      tableArray = tableArray.filter(function( obj ) {
        return obj.cellID !== cellIDs;
      })
      // Remove h4 element connected to the cellID
        console.log(cellID);
      let h4Element = document.getElementById(cellID);
      h4Element.remove();
      // Remove datalist element connected to the listID
        console.log(listID);
      let dataList = document.getElementById(listID);
      dataList.remove();
      // Remove input element connected to the listInputID
      console.log(listInputID);
      let dataInput = document.getElementById(listInputID);
      dataInput.remove();
        
    }
  });
}

// Display the selected cells as h4 elements with its cell value
function displayTableSelections(cellValue, cellId, listID, listInputID) {
    let templateChoises = document.getElementById('template-choises');
    let selectedValue = document.createElement('h4');
    selectedValue.setAttribute('id', cellId)
    selectedValue.innerHTML = cellValue;
      console.log(selectedValue);
      console.log(cellValue);
    templateChoises.appendChild(selectedValue);

    let dataListDiv = document.getElementById('datalist-alternatives-div');
    let inputData = document.createElement('Input');
    let dataList = document.createElement('datalist');

    inputData.setAttribute('list', listID);
    inputData.setAttribute('id', listInputID);
    inputData.setAttribute('placeholder', 'Append to column');
    dataList.setAttribute('id', listID);
    
    

    let listValues = `
    <option value="Append to column"></option>
    <option value="Append to row"></option>
    <option value="Append to new column"></option>
    <option value="Append to new row"></option>
    `;

    dataList.innerHTML = listValues;
      console.log(dataList);
      dataListDiv.appendChild(inputData);
      dataListDiv.appendChild(dataList);
    
}


// Set up event listener for file input
let filesInput = document.getElementById('multi-files-selector');
filesInput.addEventListener('change', handleFilesInput);

function handleFilesInput(e) {
  console.log(tableArray);
  if (tableArray.length < 1) {
    alert('Please select a template file first!');
    e = "";
    return;
  }
  const files = e.target.files;
  for (let i = 0; i < files.length; i++) {
    const file = files[i];
    const reader = new FileReader();
    reader.onload = function(event) {
      const workbook = XLSX.read(event.target.result, { type: 'binary' });
      displayExcelData(workbook, file.name);
      console.log(file.name);
    }
  reader.readAsBinaryString(file);
  }
}

function displayExcelData(workbook, filename) {
      // Define the cell locations that contain the data to be read
      console.log(tableArray[0].rowindex);
  let dataLocations = [];
  for (let i = 0; i < tableArray.length; i++) {
    dataLocations.push({ filename: '', row: tableArray[i].rowindex + 1, col: tableArray[i].cellindex + 1 });
    console.log(dataLocations);
  }
  // Create table header
  const table = document.getElementById('merged-table');
  const headerRow = document.createElement('tr');
  headerRow.innerHTML = `<th>${filename}</th>`;
  // dataLocations.forEach(loc => {
  //   headerRow.innerHTML += `<th>(${loc.row},${loc.col})</th>`;
  // });
  table.appendChild(headerRow);

  // Loop through data locations and retrieve data from specified location
  dataLocations.forEach(loc => {
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const cellAddress = XLSX.utils.encode_cell({ r: loc.row - 1, c: loc.col - 1 });
    const cell = sheet[cellAddress];
    const value = cell ? cell.v : null;
      console.log(dataLocations);
    // Create table row for data
    const row = document.createElement('tr');
    row.innerHTML = `<td>${value}</td>`;
    table.appendChild(row);
  });
}


// Download the new excel table from the merged files to a new excel workbook
let downloadBtn = document.getElementById('download-btn');
  downloadBtn.addEventListener('click', Table2XLSX);
/* Create worksheet from HTML DOM TABLE */
function Table2XLSX() {
  // Copied from https://docs.sheetjs.com/docs/ "Export an HTML Table to Excel XLSX"
  const table = document.getElementById("merged-table");
  const wb = XLSX.utils.table_to_book(table);

  /* Export to file (start a download) */
  XLSX.writeFile(wb, "AEFMerger.xlsx");

}