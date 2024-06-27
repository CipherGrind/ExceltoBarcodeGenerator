let workbook;
let worksheet;
let sheetName = "Sheet1"; // Change to the name of your sheet

document.getElementById('fileInput').addEventListener('change', handleFile);

function handleFile(event) {
    const file = event.target.files[0];
    const reader = new FileReader();
    reader.onload = function(event) {
        const data = new Uint8Array(event.target.result);
        workbook = XLSX.read(data, { type: 'array' });
        sheetName = workbook.SheetNames[0]; // Dynamically get the first sheet name
        worksheet = workbook.Sheets[sheetName];
        displayForm();
    };
    reader.readAsArrayBuffer(file);
}

function displayForm() {
    const formContainer = document.getElementById('formContainer');
    formContainer.innerHTML = '';

    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
    if (jsonData.length === 0) {
        formContainer.innerHTML = 'No data found in the uploaded file.';
        return;
    }

    let headers = jsonData[0];
    let rows = jsonData.slice(1);

    // Limit the number of columns to 4
    if (headers.length > 4) {
        headers = headers.slice(0, 4);
        rows = rows.map(row => row.slice(0, 4));
    }

    // Set grid template to accommodate row numbers and column headers
    formContainer.style.gridTemplateColumns = `40px repeat(${headers.length}, 160px)`; // Increased for better fit

    // Add column letters
    formContainer.appendChild(createCell('', 'header-cell')); // Top-left empty cell
    headers.forEach((_, index) => {
        formContainer.appendChild(createCell(columnToLetter(index), 'header-cell'));
    });

    // Add Excel column headers
    formContainer.appendChild(createCell('', 'header-cell')); // Top-left empty cell again for the row headers
    headers.forEach(header => {
        formContainer.appendChild(createCell(header, 'header-cell'));
    });

    // Add rows with row numbers and barcode cells
    rows.forEach((row, rowIndex) => {
        formContainer.appendChild(createCell(rowIndex + 1, 'header-cell row-header')); // Row number cell
        row.forEach((cellValue, colIndex) => {
            const barcodeDiv = document.createElement('div');
            barcodeDiv.className = 'cell';
            const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
            barcodeDiv.appendChild(svg);
            if (cellValue) {
                try {
                    JsBarcode(svg, cellValue, { format: "CODE128", displayValue: true, text: cellValue, width: 2, height: 50 });
                } catch (e) {
                    svg.textContent = 'Invalid Data';
                }
            } else {
                svg.textContent = 'No Data';
            }
            formContainer.appendChild(barcodeDiv);
        });
    });
}

function createCell(content, className = 'cell') {
    const cell = document.createElement('div');
    cell.className = className;
    if (typeof content === 'string' || typeof content === 'number') {
        cell.textContent = content;
    } else {
        cell.appendChild(content);
    }
    return cell;
}

function columnToLetter(column) {
    let temp;
    let letter = '';
    while (column >= 0) {
        temp = column % 26;
        letter = String.fromCharCode(temp + 65) + letter;
        column = Math.floor(column / 26) - 1;
    }
    return letter;
}

function clearForm() {
    document.getElementById('fileInput').value = '';
    document.getElementById('formContainer').innerHTML = '';
    workbook = null;
    worksheet = null;
    document.getElementById('fileInput').addEventListener('change', handleFile); // Reattach the event listener
}

function saveAsPDF() {
    const element = document.getElementById('formContainer');
    const opt = {
        margin: [0.5, 0.5, 0.5, 0.5], // [top, left, bottom, right]
        filename: 'barcodes.pdf',
        image: { type: 'jpeg', quality: 0.98 },
        html2canvas: { scale: 2 },
        jsPDF: { unit: 'in', format: 'letter', orientation: 'portrait' }
    };
    html2pdf().from(element).set(opt).save();
}

function printBarcodes() {
    const formContainer = document.getElementById('formContainer');
    const printContent = formContainer.innerHTML;
    const originalContent = document.body.innerHTML;

    document.body.innerHTML = `<div style="margin-top: 80px; display: grid; grid-template-columns: ${formContainer.style.gridTemplateColumns};">${printContent}</div>`;
    document.body.style.cssText = formContainer.style.cssText;

    window.onafterprint = function() {
        document.body.innerHTML = originalContent;
        document.body.style.cssText = '';
    };

    window.print();
}
