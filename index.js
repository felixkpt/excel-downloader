const axios = require('axios');
const fs = require('fs');
const ExcelJS = require('exceljs');
const path = require('path');

const url = '';
const downloadFolder = './downloads';
const combinedFilePath = './combined_data.xlsx';

// Function to download files from January to November 2024
async function downloadMonthlyFiles() {
    if (!fs.existsSync(downloadFolder)) {
        fs.mkdirSync(downloadFolder);
    }

    const months = Array.from({ length: 11 }, (_, i) => i + 1); // Months 1 to 11

    for (const month of months) {
        const from_date = `2024-${String(month).padStart(2, '0')}-01`;
        const to_date = `2024-${String(month).padStart(2, '0')}-30`;
        const fileName = `tickets_2024_${month}.xlsx`;
        const filePath = path.join(downloadFolder, fileName);

        try {
            const response = await axios({
                method: 'get',
                url: url + `?from_date=${from_date}&to_date=${to_date}&status=1`,
                responseType: 'stream',
            });

            const writer = fs.createWriteStream(filePath);
            response.data.pipe(writer);

            await new Promise((resolve, reject) => {
                writer.on('finish', resolve);
                writer.on('error', reject);
            });

            console.log(`Downloaded and saved: ${fileName}`);
        } catch (error) {
            console.error(`Failed to download for month ${month}:`, error.message);
        }
    }
}

// Function to combine all downloaded Excel files into one file
async function combineExcelFiles() {
    const combinedWorkbook = new ExcelJS.Workbook();

    for (const month of Array.from({ length: 11 }, (_, i) => i + 1)) {
        const filePath = path.join(downloadFolder, `tickets_2024_${month}.xlsx`);
        const monthWorkbook = new ExcelJS.Workbook();

        try {
            await monthWorkbook.xlsx.readFile(filePath);
            const worksheet = monthWorkbook.worksheets[0]; // Assuming the first sheet

            const newSheet = combinedWorkbook.addWorksheet(`Month_${month}`);
            worksheet.eachRow((row, rowNumber) => {
                const newRow = newSheet.getRow(rowNumber);
                row.eachCell((cell, colNumber) => {
                    newRow.getCell(colNumber).value = cell.value;
                });
            });

            console.log(`Added data from month ${month}`);
        } catch (error) {
            console.error(`Failed to read file for month ${month}:`, error.message);
        }
    }

    await combinedWorkbook.xlsx.writeFile(combinedFilePath);
    console.log(`Combined Excel saved at: ${combinedFilePath}`);
    return combinedFilePath;
}

// Main function to download, combine, and provide download link
async function processAndProvideLink() {
    await downloadMonthlyFiles();
    const combinedFile = await combineExcelFiles();
    console.log(`Download link: ${combinedFile}`);
}

processAndProvideLink().catch(console.error);
