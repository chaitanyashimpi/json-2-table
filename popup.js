document.getElementById('convertButton').addEventListener('click', convertJsonToTable);
document.getElementById('downloadButton').addEventListener('click', downloadExcel);

function convertJsonToTable() {
    const jsonInput = document.getElementById('jsonInput').value;
    const tableContainer = document.getElementById('tableContainer');
    tableContainer.innerHTML = ''; // Clear previous content

    try {
        const data = JSON.parse(jsonInput);
        if (Array.isArray(data)) {
            const table = document.createElement('table');
            const headerRow = document.createElement('tr');

            Object.keys(data[0]).forEach(key => {
                const th = document.createElement('th');
                th.textContent = key;
                headerRow.appendChild(th);
            });

            table.appendChild(headerRow);

            data.forEach(item => {
                const row = document.createElement('tr');
                Object.values(item).forEach(value => {
                    const cell = document.createElement('td');
                    cell.textContent = value;
                    row.appendChild(cell);
                });
                table.appendChild(row);
            });

            tableContainer.appendChild(table);
        } else {
            tableContainer.textContent = 'Invalid JSON format: Root element should be an array.';
        }
    } catch (error) {
        tableContainer.textContent = 'Invalid JSON format';
    }
}

function downloadExcel() {
    convertJsonToTable();
    
    const table = document.querySelector('table');
    if (!table) return;

    const wb = XLSX.utils.table_to_book(table);
    XLSX.writeFile(wb, 'data.xlsx');
}
