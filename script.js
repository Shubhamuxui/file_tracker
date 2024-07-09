const passcode = "123456789";  // Replace 'yourpasscode' with your actual passcode

// Function to load saved data from local Excel file
function loadSavedData() {
    const savedData = localStorage.getItem('excelData');
    if (savedData) {
        const workbook = XLSX.read(savedData, { type: 'binary' });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const files = XLSX.utils.sheet_to_json(sheet, { header: 1 });
        
        files.slice(1).forEach(row => {
            const fileCode = row[0];
            const fileName = row[1];
            const updates = JSON.parse(row[2] || '[]');
            const fileElement = createFileElement(fileCode, fileName, updates);
            document.getElementById('fileList').appendChild(fileElement);
        });
    }
}

// Load saved data when the page loads
document.addEventListener('DOMContentLoaded', function() {
    loadSavedData();
});

// Function to save data to Excel and localStorage
function saveDataToExcel() {
    const files = Array.from(document.querySelectorAll('.file')).map(file => {
        const code = file.querySelector('.file-header h2').textContent.split(': ')[0];
        const name = file.querySelector('.file-header h2').textContent.split(': ')[1];
        const updates = Array.from(file.querySelectorAll('tbody tr')).map(tr => {
            const tds = tr.querySelectorAll('td');
            return {
                date: tds[0].querySelector('input').value,
                presentAt: tds[1].querySelector('input').value,
                number: tds[2].querySelector('input').value,
                headingTo: tds[3].querySelector('input').value,
                timestamp: tds[5].textContent
            };
        });
        return [code, name, JSON.stringify(updates)];
    });

    const ws = XLSX.utils.aoa_to_sheet([['File Code', 'File Name', 'Updates'], ...files]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Files");

    const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'binary' });
    const blob = new Blob([s2ab(wbout)], { type: "application/octet-stream" });
    const url = URL.createObjectURL(blob);
    localStorage.setItem('excelData', btoa(url));
}

// Convert string to ArrayBuffer
function s2ab(s) {
    const buf = new ArrayBuffer(s.length);
    const view = new Uint8Array(buf);
    for (let i = 0; i !== s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
    return buf;
}

// Function to prompt for passcode
function promptForPasscode() {
    const input = prompt('Enter passcode:');
    return input === passcode;
}

// Add file form submission
document.getElementById('fileForm').addEventListener('submit', function(event) {
    event.preventDefault();
    const fileCode = document.getElementById('fileCode').value.trim();
    const fileName = document.getElementById('fileName').value.trim();

    if (fileCode && fileName) {
        const fileElement = createFileElement(fileCode, fileName, []);
        document.getElementById('fileList').appendChild(fileElement);

        saveDataToExcel();

        document.getElementById('fileCode').value = '';
        document.getElementById('fileName').value = '';
    } else {
        alert('Please enter both File Code and File Name.');
    }
});

// Search functionality
document.getElementById('searchInput').addEventListener('input', function() {
    const query = this.value.toLowerCase();
    const files = document.querySelectorAll('.file');

    files.forEach(file => {
        const code = file.querySelector('.file-header h2').textContent.toLowerCase();
        const name = file.querySelector('.file-header h2').textContent.toLowerCase();
        if (code.includes(query) || name.includes(query)) {
            file.style.display = '';
        } else {
            file.style.display = 'none';
        }
    });
});

// Function to create file element
function createFileElement(code, name, updates) {
    const fileDiv = document.createElement('div');
    fileDiv.className = 'file';

    const header = document.createElement('div');
    header.className = 'file-header';

    const fileTitle = document.createElement('h2');
    fileTitle.textContent = `${code}: ${name}`;

    const addButton = document.createElement('button');
    addButton.textContent = 'Add Update';
    addButton.addEventListener('click', function() {
        const updateRow = createUpdateRow('', '', '', '', '', code, name);
        tbody.insertBefore(updateRow, tbody.firstChild);
    });

    const deleteButton = document.createElement('button');
    deleteButton.textContent = 'Delete';
    deleteButton.addEventListener('click', function() {
        if (promptForPasscode()) {
            fileDiv.remove();
            saveDataToExcel();
        } else {
            alert('Incorrect passcode. Deletion canceled.');
        }
    });

    header.appendChild(fileTitle);
    header.appendChild(addButton);
    header.appendChild(deleteButton);

    const table = document.createElement('table');
    const thead = document.createElement('thead');
    const tr = document.createElement('tr');
    ['Date', 'Present at', 'Number', 'Heading to', 'Actions', 'Timestamp'].forEach(text => {
        const th = document.createElement('th');
        th.textContent = text;
        tr.appendChild(th);
    });
    thead.appendChild(tr);
    table.appendChild(thead);

    const tbody = document.createElement('tbody');
    updates.forEach(update => {
        const updateRow = createUpdateRow(update.date, update.presentAt, update.number, update.headingTo, update.timestamp, code, name);
        tbody.appendChild(updateRow);
    });
    table.appendChild(tbody);

    fileDiv.appendChild(header);
    fileDiv.appendChild(table);

    return fileDiv;
}

// Function to create update row
function createUpdateRow(date = '', presentAt = '', number = '', headingTo = '', timestamp = '', fileCode, fileName) {
    const tr = document.createElement('tr');

    const dateInput = document.createElement('input');
    dateInput.type = 'text';
    dateInput.value = date;

    const presentAtInput = document.createElement('input');
    presentAtInput.type = 'text';
    presentAtInput.value = presentAt;

    const numberInput = document.createElement('input');
    numberInput.type = 'text';
    numberInput.value = number;

    const headingToInput = document.createElement('input');
    headingToInput.type = 'text';
    headingToInput.value = headingTo;

    const tdDate = document.createElement('td');
    tdDate.appendChild(dateInput);
    tr.appendChild(tdDate);

    const tdPresentAt = document.createElement('td');
    tdPresentAt.appendChild(presentAtInput);
    tr.appendChild(tdPresentAt);

    const tdNumber = document.createElement('td');
    tdNumber.appendChild(numberInput);
    tr.appendChild(tdNumber);

    const tdHeadingTo = document.createElement('td');
    tdHeadingTo.appendChild(headingToInput);
    tr.appendChild(tdHeadingTo);

    const tdActions = document.createElement('td');
    const saveButton = document.createElement('button');
    saveButton.textContent = 'Save';
    saveButton.addEventListener('click', function() {
        if (saveButton.textContent === 'Save') {
            if (promptForPasscode()) {
                dateInput.disabled = true;
                presentAtInput.disabled = true;
                numberInput.disabled = true;
                headingToInput.disabled = true;
                saveButton.textContent = 'Edit';

                // Update timestamp
                const currentTimestamp = new Date().toLocaleString();
                tdTimestamp.textContent = currentTimestamp;

                saveDataToExcel();
            }
        } else {
            if (promptForPasscode()) {
                dateInput.disabled = false;
                presentAtInput.disabled = false;
                numberInput.disabled = false;
                headingToInput.disabled = false;
                saveButton.textContent = 'Save';
            }
        }
    });
    tdActions.appendChild(saveButton);
    tr.appendChild(tdActions);

    const tdTimestamp = document.createElement('td');
    tdTimestamp.textContent = timestamp;
    tr.appendChild(tdTimestamp);

    return tr;
}
