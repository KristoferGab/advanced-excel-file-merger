
// Exempel code copied from https://code-boxx.com/javascript-excel-html-table/ for test and reference.

const fileInput = document.getElementById('template-file-selector');
const uploadBtn = document.getElementById('template-upload');
const output = document.getElementById('template-file');

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

  const reader = new FileReader();
  
  reader.onload = (event) => {
    const data = new Uint8Array(event.target.result);
    const workbook = XLSX.read(data, {type: 'array'});
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const json = XLSX.utils.sheet_to_json(sheet);
    const table = document.createElement('table')
    table.setAttribute('id', 'templateTable');
    table.setAttribute('style','border:1px solid black;')

    // create header row
    const header = table.createTHead().insertRow();
    Object.keys(json[0]).forEach(key => {
      const cell = header.insertCell();
      cell.textContent = key;
    });

    // create data rows
    const tbody = document.createElement('tbody');
    for (let i = 0; i < json.length; i++) {
      const rowData = json[i];
      const row = tbody.insertRow();
      Object.values(rowData).forEach(value => {
        const cell = row.insertCell();
        cell.textContent = value;
      });
    }

    // add table to output div
    table.appendChild(tbody);
    output.innerHTML = '';
    output.appendChild(table);
  };
  
  reader.readAsArrayBuffer(selectedFile);
  setTimeout(function() {
    tableHandler();
  }, 1000); 
});


function tableHandler() {
    let cells = document.getElementById('templateTable').rows;

    for (var i = 0; i < cells.length; i++) {
        cells[i].addEventListener('click', function (event) {
            console.log(i);
        })
    }
}