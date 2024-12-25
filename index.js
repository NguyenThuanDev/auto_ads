const xlsx = require('xlsx');
const workbook1 = xlsx.readFile('./Bulk File Template.xlsx');
const sheets = workbook1.Sheets['Sponsored Products Campaigns'];
const workbook2 = xlsx.readFile('./data.xlsx');
const sheetData = workbook2.Sheets['Sheet1'];

// Read the template sheet with all columns and empty cells
const rawTemplateData = xlsx.utils.sheet_to_json(sheets, {
    header: 1,  // Reads the first row as headers
    defval: ""  // Sets empty cells to an empty string
});

// Extract headers and convert each row to an object with those headers
const headers = rawTemplateData[0];
const dataTemplate = rawTemplateData.slice(1).map(row => {
    const rowObject = {};
    headers.forEach((header, index) => {
        rowObject[header] = row[index] || "";
    });
    return rowObject;
});

const dataSKU = xlsx.utils.sheet_to_json(sheetData);

const changeData = (dataTemplate, post) => {
    let copyData = JSON.parse(JSON.stringify(dataTemplate));

    for (let i = 0; i < 3; i++) {
        copyData[i]['Campaign ID'] = post['Campaign Name'];
    }
    copyData[0]['Campaign Name'] = post['Campaign Name'];
    copyData[2]['SKU'] = post['SKU'];

    return copyData;
};

let result = [];
for (let i = 0; i < dataSKU.length; i++) {
    let value = changeData(dataTemplate, dataSKU[i]);
    result.push(...value);
    console.log(`Đã push sku = ${dataSKU[i]['SKU']}`);
}

// Convert result array to a worksheet
const newSheet = xlsx.utils.json_to_sheet(result);

// Create a new workbook and add the new sheet
const newWorkbook = xlsx.utils.book_new();
xlsx.utils.book_append_sheet(newWorkbook, newSheet, 'Modified Campaigns');

// Write the new workbook to a file
xlsx.writeFile(newWorkbook, './Modified_Bulk_File_Template.xlsx');

console.log('Excel file has been written successfully.');
