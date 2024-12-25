const xlsx = require('xlsx');

// Đọc file Template.xlsx
const workbook1 = xlsx.readFile('./Template.xlsx');

// Đọc sheet 'Sheet1' và chuyển dữ liệu thành JSON
const sheets = workbook1.Sheets['Sheet1'];
const dataCamp = xlsx.utils.sheet_to_json(workbook1.Sheets['data']);
const dataInterest =xlsx.utils.sheet_to_json(workbook1.Sheets['Data Interest']);
const getTarget =(str)=>{
    return dataInterest.filter(item=>{
        return item["Name"]==str; 
    })[0]["Result"]


}
// Chuyển đổi sheet 'Sheet1' thành JSON với các hàng và cột
const rawTemplateData = xlsx.utils.sheet_to_json(sheets, {
    header: 1,  // Đọc hàng đầu tiên làm header
    defval: ""  // Điền giá trị rỗng cho các ô trống
});

// Lấy headers và chuyển từng hàng thành object
const headers = rawTemplateData[0];
const dataTemplate = rawTemplateData.slice(1).map(row => {
    const rowObject = {};
    headers.forEach((header, index) => {
        rowObject[header] = row[index] || "";
    });
    return rowObject;
});

// Tạo dữ liệu mới trong sheet 'result'
const result = [];
dataCamp.forEach(item => {
    const newData = dataTemplate.map((row,index) => {
            if(index!==3){
                const newRow = { ...row };
                newRow["Campaign Name"] = item["Title"];
                newRow["Story ID"] = `s:${item["Post 1"]}`;
                newRow["Ad Set Name"] = row["Ad Set Name"] + item["Name"];
                newRow["Ad Name"] = row["Ad Name"] + item["Name"] + `_${item["Post 1"]}`;
                newRow["Flexible Inclusions"]= getTarget(item["Target chính"]);
           
                return newRow;
            }
            else{
                const newRow = { ...row };
                newRow["Campaign Name"] = item["Title"];
                newRow["Story ID"] = `s:${item["Post 1"]}`;
                newRow["Ad Set Name"] = row["Ad Set Name"] + item["Name"];
                newRow["Ad Name"] = row["Ad Name"] + item["Name"] + `_${item["Post 1"]}`;
                newRow["Flexible Inclusions"]= getTarget(item["Target phụ"])
                return newRow;
            }
         
     

        
    });
    result.push(...newData);
    if (item["Post 2"]) {
        const newData2 = dataTemplate.map((row,index) => {
            if(index!==3){
                const newRow = { ...row };
                newRow["Campaign Name"] = item["Title"];
                newRow["Story ID"] = `s:${item["Post 2"]}`
                newRow["Ad Set Name"] = row["Ad Set Name"] + item["Name"];
                newRow["Ad Name"] = row["Ad Name"] + item["Name"] + `_${item["Post 2"]}`;
                newRow["Flexible Inclusions"]= getTarget(item["Target chính"])
                return newRow;
            }
            else{
                const newRow = { ...row };
                newRow["Campaign Name"] = item["Title"];
                newRow["Story ID"] = `s:${item["Post 2"]}`
                newRow["Ad Set Name"] = row["Ad Set Name"] + item["Name"];
                newRow["Ad Name"] = row["Ad Name"] + item["Name"] + `_${item["Post 2"]}`;
                newRow["Flexible Inclusions"]= getTarget(item["Target phụ"])
                return newRow;
            }
         
     

        
    });
    result.push(...newData2);
    }
});

// Chuyển đổi `result` thành worksheet
const worksheet = xlsx.utils.json_to_sheet(result);

// Thêm hoặc ghi đè dữ liệu vào sheet 'result'
if (!workbook1.Sheets['result']) {
    workbook1.SheetNames.push('result'); // Thêm sheet mới vào danh sách sheet
}
workbook1.Sheets['result'] = worksheet;

// Ghi lại file Template.xlsx
xlsx.writeFile(workbook1, './Template.xlsx');

console.log("Dữ liệu đã được lưu vào sheet 'result'.");
