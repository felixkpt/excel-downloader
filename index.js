const axios = require('axios');
const fs = require('fs');
const ExcelJS = require('exceljs');
const path = require('path');

const url = `http://127.0.0.1/api/admin/tickets/export-list`;
const downloadFolder = './downloads';
const combinedFilePath = './combined_data.xlsx';

// Function to download JSON data from January to November 2024
async function downloadMonthlyData() {
    if (!fs.existsSync(downloadFolder)) {
        fs.mkdirSync(downloadFolder);
    }

    const months = Array.from({ length: 11 }, (_, i) => i + 1); // Months 1 to 11

    for (const month of months) {
        const from_date = `2024-${String(month).padStart(2, '0')}-01`;
        const to_date = `2024-${String(month).padStart(2, '0')}-30`;
        const fileName = `tickets_2024_${month}.json`;
        const filePath = path.join(downloadFolder, fileName);

        try {
            const response = await axios.get(url, {
                params: {
                    from_date: from_date,
                    to_date: to_date,
                    status: 1
                }
            });

            fs.writeFileSync(filePath, JSON.stringify(response.data, null, 2));
            console.log(`Downloaded and saved: ${fileName}`);
        } catch (error) {
            console.error(`Failed to download data for month ${month}:`, error.message);
        }
    }
}

// Function to combine all JSON files into a single, long Excel sheet
async function combineJsonFilesToLongExcel() {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('CombinedData');
    const months = Array.from({ length: 11 }, (_, i) => i + 1);

    let headersAdded = false;

    for (const month of months) {
        const filePath = path.join(downloadFolder, `tickets_2024_${month}.json`);

        try {
            const data = JSON.parse(fs.readFileSync(filePath, 'utf8'));

            if (data.length > 0) {
                // Add headers only once, based on the first file's data structure
                if (!headersAdded) {
                    const headers = Object.keys(data[0]);
                    sheet.addRow(headers); // Add headers to the first row
                    headersAdded = true;
                }

                // Append data rows
                data.forEach(item => {
                    const rowData = Object.values(item);
                    sheet.addRow(rowData);
                });

                console.log(`Added data from month ${month} to Excel`);
            }
        } catch (error) {
            console.error(`Failed to read or process JSON file for month ${month}:`, error.message);
        }
    }

    await workbook.xlsx.writeFile(combinedFilePath);
    console.log(`Combined Excel saved at: ${combinedFilePath}`);
    return combinedFilePath;
}

// Main function to download JSON, combine, and provide download link
async function processAndProvideLink() {
    await downloadMonthlyData();
    const combinedFile = await combineJsonFilesToLongExcel();
    console.log(`Download link: ${combinedFile}`);
}

processAndProvideLink().catch(console.error);
