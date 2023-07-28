const ExcelJS = require("exceljs");

async function readExcelFile() {
  // Replace 'your_excel_file.xlsx' with the path to your actual Excel file
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile("test.xlsx");

  // Assume you want to read the first sheet in the workbook
  const worksheet = workbook.getWorksheet(1);

  const jsonData = [];

  worksheet.eachRow((row, rowNumber) => {
    const rowData = [];
    row.eachCell((cell, colNumber) => {
      const cellValue = cell.value;
      const cellBackgroundColor =
        cell.fill && (cell.fill.fgColor || cell.fill.bgColor)
          ? cell.fill.fgColor.argb
            ? cell.fill.fgColor.argb
            : cell.fill.bgColor.argb
            ? cell.fill.bgColor.argb
            : null
          : null;
      rowData.push({ value: cellValue, backgroundColor: cellBackgroundColor });
    });
    jsonData.push(rowData);
  });

  console.log(jsonData);
}

readExcelFile();
