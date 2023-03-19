
// Exempel code copied from https://code-boxx.com/javascript-excel-html-table/ for test and reference.
// Combined code from https://docs.sheetjs.com/docs/getting-started/installation/standalone
//https://www.webslesson.info/2021/07/how-to-display-excel-data-in-html-table.html
//https://web.dev/read-files/#read-content
//Uses the library provided by https://sheetjs.com/
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

// Function to set eventlistener on all tabledata and retrieve the clicked cell from ecent object on uploaded excelfile
// Code copied from https://stackoverflow.com/questions/62259233/javascript-get-table-cell-content-on-click : credit to Teemu!
function tableHandler() {
  let tableArray = []
  const tbody = document.querySelector('#templateTable tbody');
  tbody.addEventListener('click', function (e) {
    const cell = e.target.closest('td');
    if (!cell) {return;} // Quit, not clicked on a cell
    const row = cell.parentElement;
    console.log(cell.innerHTML, row.rowIndex, cell.cellIndex);
    let cellID = row.rowIndex.toString() + cell.cellIndex.toString();
    // Changing color on selected or deselected cells and creating array with selected cells as objects.
    if (cell.style.backgroundColor !== "lightgreen") {
      cell.style.border = "2px solid green";
      cell.style.backgroundColor = "lightgreen";
      let tableObject = {cellvalue: cell.innerHTML, rowindex: row.rowIndex, 
          cellindex: cell.cellIndex, cellID: row.rowIndex.toString() + cell.cellIndex.toString()};
      tableArray.push(tableObject);
      console.log(tableArray);
      //Function to display cell value choises
      displayTableSelections(cell.innerHTML, cellID);
      // If cell is deselected make sure to remove object from array and change back color.
    } else if (cell.style.backgroundColor === "lightgreen") {
      cell.style.border = "1px solid black";
      cell.style.backgroundColor = "white";
      // Find the deselected cell and remove it from array. Credit to Peter Mortensen on https://stackoverflow.com/questions/21659888/find-and-remove-objects-in-an-array-based-on-a-key-value-in-javascript
      let cellID = row.rowIndex.toString() + cell.cellIndex.toString();
      tableArray = tableArray.filter(function( obj ) {
        return obj.cellID !== cellID;
      })
      // Remove h4 element connected to the cellID
      let h4Element = document.getElementById(cellID);
      h4Element.remove();
    }
    
  });
}

// Display the selected cells as h4 elements with its cell value
function displayTableSelections(cellValue, ID) {
    let templateChoises = document.getElementById('template-choises');
    let selectedValue = document.createElement('h4');
    selectedValue.setAttribute('id', ID)
    selectedValue.innerHTML = cellValue;
    console.log(selectedValue);
    console.log(cellValue);
    templateChoises.appendChild(selectedValue);
}

