const { default: axios } = require("axios");
const XLSX = require("xlsx");

async function readExcelFileFromURL(url) {
  axios
    .get(url)
    .then((res) => console.log(XLSX.read(res.data)))
    .catch((err) => console.log(err));
}

// Replace 'your_excel_file_url.xlsx' with the actual URL of your Excel file
const fileURL =
  "https://docs.google.com/spreadsheets/d/1BDCmWc9j1DOX7lVv9dtDItvM79sjZyiXODvbOtQ7ap8/export?format=xlsx";
readExcelFileFromURL(fileURL);
