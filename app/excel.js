
/*--------private section-----------*/

function _generateID(trNum, tdNum) {
    let tdLetter = '';
    do {
        tdNum -= 1;
        tdLetter = String.fromCharCode(65 + (tdNum % 26)) + tdLetter;
        tdNum = (tdNum / 26) >> 0;
    } while(tdNum > 0);

    return {tdLetter, trNum};
}

function _initCell (cell, id) {
    const input = document.createElement('input');
    const span = document.createElement('span');

    cell.addEventListener('click', () => {
        span.classList.add('hide');
        input.classList.remove('hide');
        input.focus();
    })

    cell.setAttribute('id', `${id.tdLetter}${id.trNum}`)
    input.addEventListener('blur', () => {
        span.classList.remove('hide');
        input.classList.add('hide');

        _calculateResult.call(this, cell);
    })

    cell.appendChild(input);
    cell.appendChild(span);
}

function _calculateResult(cell) {
    const span = cell.querySelector('span');
    const input = cell.querySelector('input');
    const value = input && input.value.trim() || '';
    let result = value;

    delete this.expressions[cell.id];
    if (value.substr(0, 1) === "=") {
        let exp = value.substr(1, value.length - 1).toUpperCase();
        const extractedCellsIDs = exp.match(/([A-Z]+[0-9]+)/g);
        if (!_validateCell.call(this, cell.id, extractedCellsIDs)) {
            input.value = ''
            return;
        }

        this.expressions[cell.id] = exp;
        if (extractedCellsIDs) {
            extractedCellsIDs.forEach((expCellID) => {
                const expCell = document.getElementById(expCellID);
                const cellValue = expCell && parseInt(expCell.querySelector(`span`).innerHTML, 10) || 0;

                exp = exp && exp.replace(expCellID, cellValue);
            })
        }

        result = eval(exp);
    }

    span.innerHTML = result;

    // It's needed for making changes in another cells which depends on changed cell
    _recalculateExpressions.call(this, cell);
}

function _validateCell(cellID, extractedCellsIDs) {
    let circularReferenceErr = false;

    const errors = [];
    const cellsForChecking = [];
    const checkCircularReferences = (extractedCellsIDs, prevCellsChain) => {
        extractedCellsIDs.forEach((extractedCellID) => {
            const newCellsChain = prevCellsChain || []
            const exp = this.expressions[extractedCellID]
            const extractedCellsIDs = exp && exp.match(/([A-Z]+[0-9]+)/g) || [];

            newCellsChain.push(extractedCellID)
            circularReferenceErr = exp &&
                                   exp.indexOf(cellID) !== -1 &&
                                   cellsForChecking.push(...newCellsChain)

            if (!circularReferenceErr && extractedCellsIDs.length) {
                checkCircularReferences(extractedCellsIDs, newCellsChain)
            }
        })
    }

    checkCircularReferences(extractedCellsIDs)
    if (circularReferenceErr) {
        errors.push(`Circular reference is not allowed. Please check cells: ${cellsForChecking.join(',')}`);
    }

    errors.forEach((message) => alert(`ERROR: ${message}`));

    return !errors.length;
}

function _recalculateExpressions(cell) {
    Object.keys(this.expressions).forEach((expCellID) => {
        let exp = this.expressions[expCellID].toUpperCase();
        if (expCellID !== cell.id && exp.indexOf(cell.id) !== -1) {
            const expCell = document.getElementById(expCellID);
            const span = expCell.querySelector(`span`);

            const extractedCellsIDs = exp.match(/([A-Z]+[0-9]+)/g);
            if (extractedCellsIDs) {
                extractedCellsIDs && extractedCellsIDs.forEach((expCellID) => {
                    const expCell = document.getElementById(expCellID);
                    const cellValue = expCell && parseInt(expCell.querySelector(`span`).innerHTML, 10) || 0;

                    exp = exp && exp.replace(expCellID, cellValue);
                })
            }

            span.innerHTML = eval(exp);

            _recalculateExpressions.call(this, expCell);
        }
    })
}

/*--------public section-----------*/

export class ExcelTable {
    constructor() {
        this.expressions = {};
    }

    createTable(container, width, height) {
        const table = document.createElement('table');
        let id = null;

        for (let trNum = 0; trNum <= width; trNum++){
            const row = table.insertRow(-1);
            for (let tdNum = 0; tdNum <= height; tdNum++) {
                const cell = row.insertCell(-1);
                id = _generateID(trNum, tdNum);
                if (!tdNum) {
                    cell.innerHTML = id.trNum;
                    cell.classList.add('headerTD');
                }

                if (!trNum) {
                    cell.innerHTML = id.tdLetter;
                    cell.classList.add('headerTD');
                }

                if(trNum && tdNum) {
                    _initCell.call(this, cell, id);
                }
            }
        }

        document.querySelector(container).appendChild(table)
    }

    setCell(tr, td, value) {
        const cell = document.getElementById(td+tr);
        cell.querySelector('input').value = value;

        _calculateResult.call(this, cell);
    }

    getCell(tr, td) {
        return document.getElementById(td+tr);
    }
}