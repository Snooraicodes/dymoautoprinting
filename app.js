let excelData = null;
let labelXml = null;
let labelObjects = [];
let columnHeaders = [];
let mappings = {};
let quantityColumn = null;
let selectedPrinter = null;

function processFiles() {
    const excelFile = document.getElementById('excelFile').files[0];
    const labelFile = document.getElementById('labelFile').files[0];

    if (!excelFile || !labelFile) {
        alert('Please upload both the Excel file and the label template.');
        return;
    }

    // Read label file
    const labelReader = new FileReader();
    labelReader.onload = function(e) {
        labelXml = e.target.result;
        parseLabelXml(labelXml);
    };
    labelReader.readAsText(labelFile);

    // Read Excel file
    const excelReader = new FileReader();
    excelReader.onload = function(e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        excelData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
        columnHeaders = excelData[0];
        createMappingInterface();
    };
    excelReader.readAsArrayBuffer(excelFile);
}

function parseLabelXml(xml) {
    const parser = new DOMParser();
    const xmlDoc = parser.parseFromString(xml, 'text/xml');
    labelObjects = [];

    // Extract Text Objects
    const textObjects = xmlDoc.getElementsByTagName('TextObject');
    for (let obj of textObjects) {
        const name = obj.getElementsByTagName('Name')[0].textContent;
        labelObjects.push(name);
    }

    // Extract Barcode Objects
    const barcodeObjects = xmlDoc.getElementsByTagName('BarcodeObject');
    for (let obj of barcodeObjects) {
        const name = obj.getElementsByTagName('Name')[0].textContent;
        labelObjects.push(name);
    }
}

function createMappingInterface() {
    const mappingDiv = document.getElementById('mapping');
    mappingDiv.innerHTML = '<h2>Map Columns to Label Objects</h2>';

    // Quantity column selection
    mappingDiv.innerHTML += '<label>Quantity Column:</label><select id="quantitySelect"><option value="">Select Quantity Column</option></select><br>';
    const quantitySelect = document.getElementById('quantitySelect');
    columnHeaders.forEach(col => {
        quantitySelect.innerHTML += `<option value="${col}">${col}</option>`;
    });

    // Object mappings
    labelObjects.forEach(obj => {
        mappingDiv.innerHTML += `<label>${obj}:</label><select id="map_${obj}"><option value="">Select Column</option></select><br>`;
        const select = document.getElementById(`map_${obj}`);
        columnHeaders.forEach(col => {
            select.innerHTML += `<option value="${col}">${col}</option>`;
        });
    });

    // Populate printer dropdown
    const printers = dymo.label.framework.getPrinters();
    if (printers.length === 0) {
        alert('No DYMO printers found. Please ensure the DYMO software is installed and the printer is connected.');
        return;
    }
    const printerSelect = document.getElementById('printerSelect');
    printerSelect.style.display = 'block';
    printers.forEach(printer => {
        printerSelect.innerHTML += `<option value="${printer.name}">${printer.name}</option>`;
    });

    document.getElementById('printButton').style.display = 'block';
}

function printLabels() {
    quantityColumn = document.getElementById('quantitySelect').value;
    selectedPrinter = document.getElementById('printerSelect').value;

    if (!quantityColumn) {
        alert('Please select the quantity column.');
        return;
    }
    if (!selectedPrinter) {
        alert('Please select a printer.');
        return;
    }

    // Gather mappings
    labelObjects.forEach(obj => {
        const select = document.getElementById(`map_${obj}`);
        mappings[obj] = select.value;
    });

    try {
        const label = dymo.label.framework.openLabelXml(labelXml);

        // Skip header row (first row)
        for (let i = 1; i < excelData.length; i++) {
            const row = excelData[i];
            const rowData = {};
            columnHeaders.forEach((col, index) => {
                rowData[col] = row[index] || '';
            });

            // Set data for each mapped object
            for (let obj in mappings) {
                const column = mappings[obj];
                if (column) {
                    label.setObjectText(obj, rowData[column]);
                }
            }

            // Print the label 'quantity' times
            const quantity = parseInt(rowData[quantityColumn]) || 1;
            for (let j = 0; j < quantity; j++) {
                label.print(selectedPrinter);
            }
        }
        alert('Labels have been sent to the printer.');
    } catch (error) {
        alert('Error printing labels: ' + error.message);
    }
}