# Tiny xlsx reader

A small library to read xlsx and csv files.

Depends on its sister project, [tiny-unzip](https://github.com/Sominemo/tiny-unzip).

## Usage
### XLSX:
```javascript
const reader = new XlsxWorkbookReader(fileBuffer);
const sheet = await reader.getSheet(0);
// readAllRows(trimTable = true, ignoreRogueSpaces, addType)
const rows = await sheet.readAllRows(true);

return rows;
```

### CSV:
```javascript
const decoder = new TextDecoder("utf-8");
const text = decoder.decode(fileBuffer);
// constructor(file, separator = ",")
const reader = new CsvReader(text);
// splits by new line
reader.read();
const rows = reader.readAllRows(true);

return rows;
```
