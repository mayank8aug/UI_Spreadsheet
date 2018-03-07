var table;
var menu1;
var menu2;
var spreadsheetInstance;
var menuLaunchedCellIndex;
var menuLaunchedRowIndex;
var sortByIndex;
var clipboardRow;

function init() {
    table = document.getElementById('mainGridId');
    menu1 = document.getElementById('menuId');
    menu2 = document.getElementById('copyPasteMenuId');

    //Creating and initializing spreadsheet instance
    spreadsheetInstance = spreadsheet(table);
    spreadsheetInstance && spreadsheetInstance.initSpreadsheet();

    attachListeners();
}

function attachListeners() {
    table.oncontextmenu = positionMenu;
    var menuItems = document.querySelectorAll('.menuItem');
    var menu, i;
    for(i = 0; i < menuItems.length; i++) {
        menu = menuItems[i];
        menu.addEventListener('mousedown', function(e){
            processMenuAction(e);
        });
    }
    var menu2Items = document.querySelectorAll('.menuItem2');
    for(i = 0; i < menu2Items.length; i++) {
        menu = menu2Items[i];
        menu.addEventListener('mousedown', function(e){
            processCopyPasteMenuAction(e);
        });
    }
    document.addEventListener('mousedown', function() {
        hideMenu(menu1);
        hideMenu(menu2);
    });
}

function processCopyPasteMenuAction(e) {
    var clickedElId = e.target.getAttribute('id');
    switch(clickedElId) {
        case 'cutRowId':
            cutRow();
            break;
        case 'copyRowId':
            copyRow();
            break;
        case 'pasteRowAboveId':
            pasteRow(menuLaunchedRowIndex);
            break;
        case 'pasteRowBelowId':
            pasteRow(menuLaunchedRowIndex+1);
            break;
    }
    hideMenu(menu2);
    return false;
}

function cutRow() {
    copyRow();
    deleteRow();
    enablePasteRow();
}

function copyRow() {
    clipboardRow = spreadsheetInstance && spreadsheetInstance.getRowData(menuLaunchedRowIndex);
    enablePasteRow();
}

function enablePasteRow() {
    var items = document.querySelectorAll('.pasteDisabled');
    for(var i = 0; i < items.length; i++)
        items[i].classList.remove('pasteDisabled');
}
function pasteRow(index) {
    clipboardRow && spreadsheetInstance && spreadsheetInstance.pasteRow(index, clipboardRow);
}

function processSortAction(e) {
    var index = e.target.parentNode.cellIndex;
    var descOrder;
    if(sortByIndex === index) {
        descOrder = true;
        sortByIndex = -1;
    } else {
        sortByIndex = index;
    }
    spreadsheetInstance && spreadsheetInstance.sortData(index, descOrder);
}

function processMenuAction(e) {
    var clickedElId = e.target.getAttribute('id');
    switch(clickedElId) {
        case 'insertRowAboveId':
            insertRow(menuLaunchedRowIndex);
            break;
        case 'insertRowBelowId':
            insertRow(menuLaunchedRowIndex+1);
            break;
        case 'insertRowEndId':
            insertRow(-1);
            break;
        case 'deleteRowId':
            deleteRow();
            break;
        case 'insertColLtId':
            insertCol(menuLaunchedCellIndex);
            break;
        case 'insertColRtId':
            insertCol(menuLaunchedCellIndex+1);
            break;
        case 'insertColEndId':
            insertCol(-1);
            break;
        case 'deleteColId':
            deleteCol();
            break;
    }
    hideMenu(menu1);
    return false;
}

function positionMenu(e) {
    var cellDataNode = e.target.parentNode;
    menuLaunchedRowIndex = cellDataNode.parentNode.rowIndex;
    menuLaunchedCellIndex = cellDataNode.cellIndex;
    if(menuLaunchedCellIndex === 0) {
        if(menu2) {
            showMenu(menu2);
            setMenuCoordinates(e, menu2);
        }
    } else if(menuLaunchedRowIndex === 0) {

    } else {
        if(menu1) {
            showMenu(menu1);
            setMenuCoordinates(e, menu1);
        }
    }
    return false;
}

function setMenuCoordinates(e, currMenu) {
    var clickedX = e.clientX;
    var clickedY = e.clientY;
    var docWidth = document.body.offsetWidth;
    var docHeight = document.body.offsetHeight;
    var menuWidth = currMenu.offsetWidth;
    var menuHeight = currMenu.offsetHeight;
    var x,y;
    if(clickedX + menuWidth > docWidth) {
        x = clickedX - menuWidth;
    }
    if(clickedY + menuHeight > docHeight) {
        y = clickedY - menuHeight;
    }
    currMenu.style.left = String(x ? x : clickedX);
    currMenu.style.top = String(y ? y : clickedY);
}

function insertRow(index) {
    spreadsheetInstance && spreadsheetInstance.insertRow(index);
}

function deleteRow() {
    spreadsheetInstance && spreadsheetInstance.deleteRow(menuLaunchedRowIndex);
}

function insertCol(index) {
    spreadsheetInstance && spreadsheetInstance.insertColumn(index);
}

function deleteCol() {
    spreadsheetInstance && spreadsheetInstance.deleteColumn(menuLaunchedCellIndex);
}

function showMenu(thisMenu) {
    if(thisMenu)
        thisMenu.style.display = 'block';
}

function hideMenu(thisMenu) {
    if(thisMenu)
        thisMenu.style.display = 'none';
}

function exportCSV() {
    var csvData = spreadsheetInstance && spreadsheetInstance.getFullData();
    var temp;
    var link = document.getElementsByTagName('a');
    temp = new Blob([csvData], {type: "text/csv"});
    if(!(link && link.length > 0)) {
        link = document.createElement("a");
        link.download = 'spreadsheet.csv';
        link.style.display = "none";
        document.body.appendChild(link);
    } else {
        link = link[0];
    }
    link.href = window.URL.createObjectURL(temp);
    link.click();
}