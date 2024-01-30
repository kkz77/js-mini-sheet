const spreadsheetContainer = document.querySelector('#spread-sheet-container')
const exportBtn = document.querySelector("#export-btn")
const cellStatus = document.querySelector("#cell-status")
const formula = document.querySelector("#input-formula")
const inputSubmit = document.querySelector("#input-submit")
const errorMessage = document.querySelector("#error-message")
const fillColor = document.querySelector("#fill-color")
const textColor = document.querySelector("#text-color")
const bold = document.querySelector("#bold")
const italic = document.querySelector("#italic")
const reset = document.querySelector("#reset")
const row = 6
const column = 6
const spreadsheet = []
const operation = ['SUM', 'AVERAGE', 'MIN', 'MAX'];
class Cell {
    constructor(isHeader, disabled, data, row, rowName, column, columnName, active = false) {
        this.isHeader = isHeader
        this.disabled = disabled
        this.data = data
        this.row = row
        this.rowName = rowName
        this.column = column
        this.columnName = columnName
        this.active = active
    }
}

initSpreadSheet()
function initSpreadSheet() {
    fillColor.value = "#DDDDDD"
    for (let i = 0; i < row; i++) {
        let spreadsheetRow = []
        for (let j = 0; j < column; j++) {
            let cell_data = ""
            let disable_active = false
            let isHeader = false
            if (i == 0) {
                cell_data = String.fromCharCode(j + 64)
                disable_active = true
                isHeader = true
            }
            if (j == 0) {
                cell_data = i
                disable_active = true
                isHeader = true
                columnName = ""
            }
            else {
                columnName = String.fromCharCode(j + 64)
            }

            if (cell_data <= 0) {
                cell_data = ""
            }
            rowName = i

            const cell = new Cell(isHeader, disable_active, cell_data, i, rowName, j, columnName, false)
            spreadsheetRow.push(cell)

        }
        spreadsheet.push(spreadsheetRow)
    }
    drawSheet()
}

exportBtn.onclick = function (e) {
    let csv = ""
    for (let i = 1; i < spreadsheet.length; i++) {
        csv += spreadsheet[i].filter((item) => !item.isHeader)
            .map((item) => item.data).join(",") + "\r\n"
    }

    const csvObj = new Blob([csv]) //creates a Blob (Binary Large Object) with the content of the CSV data (csv variable)
    const csvUrl = URL.createObjectURL(csvObj) //creates a URL for the Blob object to reference the Blob in the browser.
    const a = document.createElement('a')
    a.href = csvUrl
    a.download = 'Exported Spreadsheet.csv'
    a.click()
    URL.revokeObjectURL(csvUrl); //Releases the resources associated with the URL created for the Blob
}

function drawSheet() {
    spreadsheetContainer.innerHTML = ""
    for (let i = 0; i < spreadsheet.length; i++) {
        const rowCoverEl = rowContainerEl()
        spreadsheetContainer.append(rowCoverEl)
        for (let j = 0; j < spreadsheet[i].length; j++) {
            const cell = spreadsheet[i][j]
            rowCoverEl.append(createCellEl(cell))
        }
    }
}

function createCellEl(cell) {
    const cellEl = document.createElement("input")
    cellEl.className = "cell"
    cellEl.id = "cell_" + cell.row + cell.column
    cellEl.value = cell.data
    cellEl.disabled = cell.disabled
    if (cell.isHeader) {
        cellEl.classList.add("header")
    }

    if ((cell.row == 0 && cell.column == 1) || (cell.row == 1 && cell.column == 0)) {
        cellEl.classList.add("active")
    }

    if (cell.row == 1 && cell.column == 1) {
        cellEl.classList.add("cell-active")
        cellEl.focus()
    }
    cellStatus.innerHTML = "A1"
    cellEl.onclick = () => handleClick(cell)
    cellEl.onchange = (e) => handleChange(e.target.value, cell)
    inputSubmit.onclick = function () {
        const cell = changeValueToRowColl(cellStatus.innerHTML)
        const formulaData = formula.value
        if (formulaData == "") {
            errorMessage.style.visibility = 'visible';
            errorMessage.innerHTML = "Please type correct formula to calculate"
            setTimeout(function () {
                errorMessage.style.visibility = "hidden"
            }, 3000);
        }
        else if (!formulaData.startsWith("=")) {
            const cell = changeValueToRowColl(cellStatus.innerHTML)
            cell.value = formula.value
            formula.value = ""
        }
        else if (operation.some(item => formulaData.includes(item))) {
            const value = handleArrayFormula(formulaData)
            if(value === undefined){
                cell.value = ""
            }
            else{
                cell.value = value
            }
            formula.value = ""
        }
        else {
            const value = handleValue(formulaData)
            if(value === undefined){
                cell.value = ""
            }
            else{
                cell.value = value
            }
            formula.value = ""
        }
    }
    return cellEl
}

function rowContainerEl() {
    const rowContainerEl = document.createElement("div")
    rowContainerEl.className = "cell-row"
    return rowContainerEl
}

function handleClick(cell) {
    clearHeaderActive()
    const columnHeader = spreadsheet[0][cell.column]
    const rowHeader = spreadsheet[cell.row][0]
    const columnHeaderEl = getElfromRowCol(columnHeader.row, columnHeader.column)
    const rowHeaderEl = getElfromRowCol(rowHeader.row, rowHeader.column)
    columnHeaderEl.classList.add("active")
    rowHeaderEl.classList.add("active")
    cellStatus.innerHTML = cell.columnName + "" + cell.rowName
    const currentCell = changeValueToRowColl(cellStatus.innerHTML)
    const currentColor = window.getComputedStyle(currentCell)
    const bgColor = currentColor.backgroundColor.match(/\d+/g)
    const color = currentColor.color.match(/\d+/g)
    fillColor.value = rgbToHex(bgColor)
    textColor.value = rgbToHex(color)
}

function handleChange(value, cell) {
    cell.data = value
}

function clearHeaderActive() {
    for (let i = 0; i < spreadsheet.length; i++) {
        for (let j = 0; j < spreadsheet.length; j++) {
            const cell = spreadsheet[i][j]
            let cellEl = getElfromRowCol(cell.row, cell.column)
            if (cell.isHeader) {
                cellEl.classList.remove("active")
            }

            if (cell.row == 1 && cell.column == 1) {
                cellEl.classList.remove("cell-active")
            }
        }
    }
}

function changeValueToRowColl(value) {
    const items = value.split('')
    const column = items[0].charCodeAt(0) - 64
    const row = items[1]
    return getElfromRowCol(row, column)
}

function getElfromRowCol(row, col) {
    return document.querySelector("#cell_" + row + col)
}

function handleValue(value) {
    const tokens = parseStringToArray(value)
    const newValue = replaceValue(tokens)
    try {
        if (newValue[0] == "=") {
            let val = newValue.filter((item) => item !== '=')
            let data = val.join('')
            return eval(data)
        }
    } catch (error) {
        errorMessage.style.visibility = "visible"
        errorMessage.innerHTML = "Invalid formula format!!!"
        setTimeout(function () {
            errorMessage.style.visibility = "hidden"
        }, 3000);
    }

}

function handleArrayFormula(value) {
    const tokens = parseStringToArray2(value)
    const newValue = replaceValue(tokens)
    let total = 0
    let var1 = parseFloat(newValue[2])
    let var2 = parseFloat(newValue[3])
    if (newValue[1] == "SUM") {
        total = var1 + var2
    }
    else if (newValue[1] == "AVERAGE") {
        total = (var1 + var2) / 2
    }
    else if (newValue[1] == "MAX") {
        if (var1 >= var2) {
            total = var1
        }
        if (var1 < var2) {
            total = var2
        }
    }
    else if (newValue[1] == "MIN") {
        if (var1 >= var2) {
            total = var2
        }
        if (var1 < var2) {
            total = var1
        }
    }
    else {
        console.log('Start with "=" to use formula')
    }
    return total
}

function replaceValue(values) {
    const regex = /^[A-Z]\d+$/;
    try {
        const replacedValues = values.map(item => {
            return regex.test(item) ? changeValueToRowColl(item).value : item;
        });
        return replacedValues
    } catch (error) {
        errorMessage.style.visibility = "visible"
        errorMessage.innerHTML = "Invalid input format"
        setTimeout(function () {
            errorMessage.style.visibility = "hidden"
        }, 3000);
    }

}

function parseStringToArray(expression) {
    const tokens = expression.match(/([A-Za-z0-9]+|[+\-*\/=()])/g) || [];
    return tokens
}

function parseStringToArray2(expression) {
    const regex = /^=([A-Z]+)\(([\w\d]+):([\w\d]+)\)$/;
    const match = expression.match(regex);
    try {
        if (match) {
            const components = match.slice(1);
            return ['='].concat(components);
        }
    }
    catch (error) {
        errorMessage.style.visibility = "visible"
        errorMessage.innerHTML = "Invalid input format"
        setTimeout(function () {
            errorMessage.style.visibility = "hidden"
        }, 3000);
    }

}

bold.onclick = function () {
    const cell = changeValueToRowColl(cellStatus.innerHTML)
    return cell.classList.contains('bold-text') ? cell.classList.remove('bold-text') : cell.classList.add('bold-text')
}

italic.onclick = function () {
    const cell = changeValueToRowColl(cellStatus.innerHTML)
    return cell.classList.contains('italic-text') ? cell.classList.remove('italic-text') : cell.classList.add('italic-text')
}

fillColor.onchange = function () {
    const cell = changeValueToRowColl(cellStatus.innerHTML)
    cell.style.backgroundColor = fillColor.value
}

textColor.onchange = function () {
    const cell = changeValueToRowColl(cellStatus.innerHTML)
    cell.style.color = textColor.value
}

reset.onclick = function () {
    const cell = changeValueToRowColl(cellStatus.innerHTML)
    cell.style.color = "#000"
    cell.style.backgroundColor = "#fff"
}

function rgbToHex(match) {
    var r = parseInt(match[0]);
    var g = parseInt(match[1]);
    var b = parseInt(match[2]);
    r = Math.min(255, Math.max(0, r));
    g = Math.min(255, Math.max(0, g));
    b = Math.min(255, Math.max(0, b));
    return `#${(1 << 24 | r << 16 | g << 8 | b).toString(16).slice(1).toUpperCase()}`;
}