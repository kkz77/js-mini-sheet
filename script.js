const spreadsheetContainer = document.querySelector('#spread-sheet-container')
const exportBtn = document.querySelector("#export-btn")
const cellStatus = document.querySelector("#cell-status")
const formula = document.querySelector("#input-formula")
const inputSubmit = document.querySelector("#input-submit")
const row = 6
const column = 6
const spreadsheet = []
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
        const value = handleValue(formula.value)
        cell.value = value
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
    if (newValue[0] == "=") {
        let val = newValue.filter((item) => item !== '=')
        let data = val.join('')
        return eval(data)
    }
}

function replaceValue(values) {
    const regex = /^[A-Z]\d+$/;
    const replacedValues = values.map(item => {
        return regex.test(item) ? changeValueToRowColl(item).value : item;
    });
    return replacedValues
}

function parseStringToArray(expression) {
    const tokens = expression.match(/([A-Za-z0-9]+|[+\-*\/=()])/g) || [];
    return tokens
}
