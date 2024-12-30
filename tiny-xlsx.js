// author: Sergey Dilong
//
// Useful links:
// https://habr.com/ru/articles/593397/
// https://stackoverflow.com/questions/18334314/what-do-excel-xml-cell-attribute-values-mean/18346273#18346273

import { ZipReader } from "tiny-unzip";

function parseXmlString(dataView) {
    // splice buffer from start until ">" codepoint or until 100 characters
    let detectedEncoding = null;

    let offset = 6;
    let closingTagCharCode = ">".charCodeAt(0);
    for (let i = 0; i < 100; i++) {
        if (dataView.getUint8(offset) === closingTagCharCode) {
            break;
        }

        offset++;
    }

    // decode the declaration and look for encoding attribute
    let declarationSlice = dataView.buffer.slice(0, offset + 1);
    let xmlDeclaration = new TextDecoder().decode(declarationSlice);
    let encodingMatch = xmlDeclaration.match(/encoding="([^"]+)"/);
    if (encodingMatch) {
        detectedEncoding = encodingMatch[1];
    }

    // console.log("Detected encoding", detectedEncoding);

    // decode the whole string with the detected encoding
    let xmlString = new TextDecoder(detectedEncoding || "utf-8").decode(dataView.buffer);

    const parser = new DOMParser();
    let dom = parser.parseFromString(xmlString, "text/xml");

    return dom;
}

function readAllRows(readRowAt, trimTable, ignoreRogueSpaces) {
    let rows = [];
    let i = 0;
    let emptyStartedAt = null;
    let maxWidth = 0;
    let maxPhysicalWidth = 0;
    const trimCheck = (cell) => cell === undefined ||
        (ignoreRogueSpaces && (cell === "" || cell.trim() === ""));

    while (true) {
        const row = readRowAt(i);
        if (!row) {
            break;
        }

        rows.push(row);

        if (trimTable) {
            let empty = true;

            if (row.length > maxPhysicalWidth) {
                maxPhysicalWidth = row.length;
            }

            for (let j = row.length - 1; j >= 0; j--) {
                if (!trimCheck(row[j])) {
                    if (j > maxWidth) maxWidth = j
                    empty = false;
                    break;
                }
            }

            if (empty) {
                if (emptyStartedAt === null) {
                    emptyStartedAt = i;
                }
            } else {
                emptyStartedAt = null;
            }
        }

        i++;
    }

    if (trimTable) {
        if (emptyStartedAt !== null) {
            // console.log("Trimming length from", rows.length, "to", emptyStartedAt);
            rows = rows.slice(0, emptyStartedAt);
        }

        if (maxWidth < maxPhysicalWidth) {
            // console.log("Trimming width from", maxPhysicalWidth, "to", maxWidth);
            for (let i = 0; i < rows.length; i++) {
                rows[i] = rows[i].slice(0, maxWidth + 1);
            }
        }
    }

    return rows;
}

export class XlsxSheetXmlReader {
    constructor(worksheetXmlDataView, sharedStringsDomTree) {
        this.worksheetXmlDataView = worksheetXmlDataView;

        this.sharedStringsDomTree = sharedStringsDomTree;
        this.worksheetDomTree = parseXmlString(worksheetXmlDataView);
    }

    readSharedString(index) {
        // excel stores strings in a separate file
        return this.sharedStringsDomTree.getElementsByTagName("si")[index].textContent;
    }

    cellLetterToIndex(letter) {
        let index = 0;
        for (let i = 0; i < letter.length; i++) {
            index = index * 26 + letter.charCodeAt(i) - 64;
        }

        return index - 1;
    }

    indexToCellLetter(index) {
        let letter = "";
        while (index >= 0) {
            letter = String.fromCharCode(65 + (index % 26)) + letter;
            index = Math.floor(index / 26) - 1;
        }

        return letter;
    }

    getCellNameComponents(cellName, onlyLetter = false) {
        let cellLetter = "";
        for (let k = 0; k < cellName.length; k++) {
            if (cellName.charCodeAt(k) >= 65 && cellName.charCodeAt(k) <= 90) {
                cellLetter += cellName[k];
            } else {
                break;
            }
        }

        if (onlyLetter) { return cellLetter; }

        let cellNumber = cellName.slice(cellLetter.length);

        return [cellLetter, cellNumber];
    }

    getCell2DIndex(cellName) {
        const components = this.getCellNameComponents(cellName);

        const letter = components[0];
        const number = parseInt(components[1]);

        return [number - 1, this.cellLetterToIndex(letter)];
    }

    readRowAt(i, addType) {
        const row = this.worksheetDomTree.querySelector(`sheetData>row[r="${i + 1}"]`);

        if (!row) {
            return undefined;
        }

        const cells = row.getElementsByTagName("c");

        const rowArray = [];
        for (let j = 0; j < cells.length; j++) {
            const cell = cells[j];

            const value = this.interpretCell(cell, addType);

            const cellName = cell.getAttribute("r");
            const cellLetter = this.getCellNameComponents(cellName, true);
            const cellIndex = this.cellLetterToIndex(cellLetter);

            if (cellIndex !== rowArray.length) {
                for (let k = rowArray.length; k < cellIndex; k++) {
                    rowArray.push(undefined);
                }
            }

            rowArray.push(value);
        }

        return rowArray;
    }


    // The xml might contain empty rows which used to contain data
    // trimTable - if true, will remove all empty rows at the end of the table
    readAllRows(trimTable = true, ignoreRogueSpaces, addType) {
        return readAllRows(i => this.readRowAt(i, addType), trimTable, ignoreRogueSpaces);
    }

    getCellRawValue(cell) {
        let valueQuery = cell.getElementsByTagName("v");
        if (valueQuery.length === 0) {
            return undefined;
        }

        return valueQuery[0].textContent;
    }

    interpretCell(cell, addType = false) {
        const type = cell.getAttribute("t");
        let v;

        if (type === "n" || type === null) {
            // number
            const cellValue = this.getCellRawValue(cell);
            
            if (cellValue === undefined) {
                v = undefined;
            } else {
                v = parseFloat(cellValue);
                if (isNaN(v)) {
                    // unknown number format, return as string
                    v = cellValue;
                }
            }
        } else if (type === "s") {
            // shared string
            v = this.readSharedString(parseInt(this.getCellRawValue(cell)));
        } else if (type === "b") {
            // boolean
            v = this.getCellRawValue(cell) === "1";
        } else {
            // unknown type, return as string
            v = this.getCellRawValue(cell);
        }

        if (addType) {
            v[XlsxSheetXmlReader.cellTypeSymbol] = type;
        }

        return v;
    }

    getCellWithName(name, addType) {
        const cell = this.worksheetDomTree.querySelector(`c[r="${name}"]`);

        if (!cell) {
            return undefined;
        }

        return this.interpretCell(cell, addType);
    }
}

XlsxSheetXmlReader.cellTypeSymbol = Symbol("cellType");

export class XlsxWorkbookReader {
    constructor(file) {
        // console.log(file);
        this.file = file;
        this.zipFile = new ZipReader(file);

        this.sharedStringsDomTree = undefined;
        this.workbookDomTree = undefined;
        this.workbookRelDomTree = undefined;

        this.worksheets = {}
    }

    async parseWorkbook() {
        // console.log("Start extraction", this.zipFile);
        await this.zipFile.read();
        // console.log("Workbook inflated");

        const sharedStrings = await this.zipFile.extractByFileName("xl/sharedStrings.xml");
        // console.log("Shared strings", sharedStrings);
        this.sharedStringsDomTree = parseXmlString(new DataView(await sharedStrings.arrayBuffer()));

        const workbook = await this.zipFile.extractByFileName("xl/workbook.xml");
        this.workbookDomTree = parseXmlString(new DataView(await workbook.arrayBuffer()));

        // contains paths to sheet files
        const workbookRel = await this.zipFile.extractByFileName("xl/_rels/workbook.xml.rels");
        this.workbookRelDomTree = parseXmlString(new DataView(await workbookRel.arrayBuffer()));
    }

    async getSheet(n) {
        // n - order of the sheet in the workbook, does not correspond to sheetId
        if (!this.workbookDomTree) {
            await this.parseWorkbook();
        }

        if (this.worksheets[n]) {
            return this.worksheets[n];
        }

        const sheet = this.workbookDomTree.getElementsByTagName("sheet")[n];
        if (!sheet) {
            return undefined;
        }

        const sheetId = sheet.getAttribute("sheetId");

        const relId = sheet.getAttribute("r:id");
        const rel = this.workbookRelDomTree.querySelector(`Relationship[Id="${relId}"]`);
        const sheetPath = `xl/${rel.getAttribute("Target")}`;
        const sheetEntry = await this.zipFile.extractByFileName(sheetPath);
        const sheetXmlDataView = new DataView(await sheetEntry.arrayBuffer());

        const sheetReader = new XlsxSheetXmlReader(sheetXmlDataView, this.sharedStringsDomTree);

        this.worksheets[sheetId] = sheetReader;

        return sheetReader;
    }

    async getAllSheets() {
        // will return sheets in the 
        // order they are listed in the workbook
        // if you need sheetId addressing, use this.worksheets
        // after calling this method
        let i = 0;
        let sheets = [];
        while (true) {
            const sheet = await this.getSheet(i);
            if (!sheet) {
                break;
            }

            sheets.push(sheet);
            i++;
        }

        return sheets;
    }
}

export class CsvReader {
    constructor(file, separator = ",") {
        this.file = file;
        this.separator = separator;
    }

    read() {
        const lines = this.file.split("\n");
        this.lines = lines;
    }

    readRowAt(i) {
        const row = this.lines[i];

        if (!row) {
            return undefined;
        }

        // Excel escaping: 12345,COCR,100.93,980,"Transfer, to card ""32"" **********"

        const rowArray = [];
        let cell = "";
        let inQuote = false;

        for (let j = 0; j < row.length; j++) {
            const char = row[j];

            if (j === row.length - 1 && char === "\r") {
                continue;
            }

            if (char === '"') {
                inQuote = !inQuote;
                if (inQuote && row[j - 1] === '"') {
                    cell += '"';
                }
            } else if (char === this.separator && !inQuote) {
                rowArray.push(cell);
                cell = "";
            } else {
                cell += char;
            }
        }

        rowArray.push(cell);

        return rowArray;
    }

    readAllRows(trimTable = true, ignoreRogueSpaces) {
        return readAllRows(i => this.readRowAt(i), trimTable, ignoreRogueSpaces);
    }
}