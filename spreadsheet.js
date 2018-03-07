var spreadsheet = function (spreadsheetEl) {

    function initSpreadsheet() {
        // Create a 5x5 spreadsheet as starting point
        var m = 5;
        var n = 5;
        var i, j, row, dataCol;
        for(i = 0; i < m; i++) {
            row = spreadsheetEl.insertRow(i);
            if(i === 0) {
                row.addEventListener('click', function(e){
                    if(!(e.target.getAttribute('id') === 'exportCSVId'))
                        processSortAction(e);
                });
            }
            for(j = 0; j < n; j++) {
                dataCol = row.insertCell(j);
                if(i > 0) {
                    dataCol.innerHTML = "<div class='cell' contenteditable='true'></div>";
                } else {
                    dataCol.innerHTML = "<button id='exportCSVId' onclick='exportCSV()'>Export CSV</button>";
                }
            }
        }
        // Add header and row labels
        updateHeaderLabels();
        updateRowLabels();
    }

    function updateHeaderLabels() {
        var row = spreadsheetEl.rows[0];
        var cell;
        for(var i = 1; i < row.cells.length; i++) {
            cell = row.cells[i];
            cell.setAttribute('class', 'headerLabel');
            cell.innerHTML = "<div class='headerText'>C"+(i)+"</div>";
        }
    }

    function updateRowLabels() {
        var cell;
        for(var i = 1; i < spreadsheetEl.rows.length; i++) {
            cell = spreadsheetEl.rows[i].cells[0];
            cell.setAttribute('class', 'rowLabel');
            cell.innerHTML = "<div>R"+(i)+"</div>";
        }
    }

    function insertRow(rowIndex) {
        var newRow = spreadsheetEl.insertRow(rowIndex ? rowIndex : -1);
        var dataCol;
        for(var i = 0; i < spreadsheetEl.rows[0].cells.length; i++) {
            dataCol = newRow.insertCell(i);
            dataCol.innerHTML = "<div class='cell' contenteditable='true'></div>";
        }
        updateRowLabels();
    }

    function deleteRow(rowIndex) {
        spreadsheetEl.deleteRow(rowIndex ? rowIndex : -1);
        updateRowLabels();
    }

    function insertColumn(colIndex) {
        var rows = spreadsheetEl.rows;
        var row, dataCol;
        for(var i = 0; i < rows.length; i++) {
            row = spreadsheetEl.rows[i];
            dataCol = row.insertCell(colIndex);
            dataCol.innerHTML = "<div class='cell' contenteditable='true'></div>";
        }
        updateHeaderLabels();
    }

    function deleteColumn(colIndex) {
        var rows = spreadsheetEl.rows;
        var row;
        for(var i = 0; i < rows.length; i++) {
            row = rows[i];
            row.deleteCell(colIndex);
        }
        updateHeaderLabels();
    }

    function sortData(sortByIndex, descOrder) {
        var rows = spreadsheetEl.rows;
        var continueSorting = true;
        var needSwapping, i, row1Data, row2Data;
        while(continueSorting) {
            continueSorting = false;
            for(i = 1; i < rows.length - 1; i++) {
                needSwapping = false;
                row1Data = rows[i].cells[sortByIndex];
                row2Data = rows[i+1].cells[sortByIndex];
                if(descOrder ? row1Data.innerText.toLowerCase() < row2Data.innerText.toLowerCase() : row1Data.innerText.toLowerCase() > row2Data.innerText.toLowerCase()) {
                    needSwapping = true;
                    spreadsheetEl.rows[i].parentNode.insertBefore(rows[i+1], rows[i]);
                    continueSorting = true;
                    break;
                }
            }
        }
        updateRowLabels();
        updateSortIcons(sortByIndex, descOrder);
    }
    
    function updateSortIcons(sortByIndex, descOrder) {
        var header = spreadsheetEl.rows[0];
        var colEl;
        for(var i = 1; i < header.cells.length; i++) {
            colEl = header.cells[i];
            colEl.classList.remove('ascSort');
            colEl.classList.remove('descSort');
        }
        var currCol = header.cells[sortByIndex];
        if(descOrder) {
            currCol.classList.add('descSort');
            currCol.classList.remove('ascSort');
        } else {
            currCol.classList.remove('descSort');
            currCol.classList.add('ascSort');
        }
    }

    function getRowData(index) {
        return spreadsheetEl.rows[index].cloneNode(true);
    }

    function pasteRow(index, rowData) {
        var newRow = spreadsheetEl.insertRow(index ? index : -1);
        var dataCol;
        for(var i = 0; i < rowData.cells.length; i++) {
            dataCol = newRow.insertCell(i);
            dataCol.innerHTML = rowData.cells[i].innerHTML;
        }
        updateRowLabels();
    }

    function getFullData() {
        var fullData = [];
        var rows = spreadsheetEl.rows;
        var rowData, cols, i, j;
        for(i = 1; i < rows.length; i++) {
            rowData = [];
            cols = rows[i].cells;
            for(j = 1; j < cols.length; j++) {
                rowData.push(cols[j].innerText.trim());
            }
            fullData.push(rowData.join(","));
        }
        return fullData.join("\n");
    }

    return {
        initSpreadsheet : initSpreadsheet,
        insertRow: insertRow,
        deleteRow: deleteRow,
        insertColumn: insertColumn,
        deleteColumn: deleteColumn,
        sortData: sortData,
        getRowData: getRowData,
        pasteRow: pasteRow,
        getFullData: getFullData
    }
};