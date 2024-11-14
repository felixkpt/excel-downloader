const axios = require('axios');
const fs = require('fs');
const ExcelJS = require('exceljs');
const path = require('path');

const url = process.argv[2] || '';

console.log("URL: ", url)

const status = 'all';
const downloadFolder = './downloads';
const combinedFilePath = `./combined_data_${status}.xlsx`;

// Function to download JSON data in 10-day intervals from January to yesterday
async function downloadDataByIntervals() {
    if (!fs.existsSync(downloadFolder)) {
        fs.mkdirSync(downloadFolder);
    }

    const today = new Date();
    const startYear = 2024;

    for (let month = 1; month <= 12; month++) {
        let startDay = 1;
        while (startDay <= 31) {
            const from_date = new Date(startYear, month - 1, startDay);
            let endDay = startDay + 9;
            const to_date = new Date(startYear, month - 1, endDay);

            // Stop if the end date goes beyond today
            if (to_date > today) break;

            const formattedFromDate = from_date.toISOString().split('T')[0];
            const formattedToDate = to_date.toISOString().split('T')[0];
            const fileName = `tickets_${formattedFromDate}_to_${formattedToDate}.json`;
            const filePath = path.join(downloadFolder, fileName);

            try {
                const response = await axios.get(url, {
                    params: {
                        from_date: formattedFromDate,
                        to_date: formattedToDate,
                        status: status
                    }
                });

                fs.writeFileSync(filePath, JSON.stringify(response.data, null, 2));
                console.log(`Downloaded and saved: ${fileName}`);
            } catch (error) {
                console.error(`Failed to download data for ${formattedFromDate} to ${formattedToDate}:`, error.message);
            }

            startDay += 10;
        }
    }
}

// Function to combine all JSON files into a single, long Excel sheet
async function combineJsonFilesToLongExcel() {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('CombinedData');

    let headersAdded = false;
    const files = fs.readdirSync(downloadFolder).filter(file => file.endsWith('.json'));

    for (const fileName of files) {
        const filePath = path.join(downloadFolder, fileName);

        try {
            const data = JSON.parse(fs.readFileSync(filePath, 'utf8'));

            if (data.length > 0) {
                // Add headers only once, based on the first file's data structure
                if (!headersAdded) {
                    const headers = Object.keys(data[0]);
                    sheet.addRow(headers);
                    headersAdded = true;
                }

                // Append data rows
                data.forEach(item => {
                    const rowData = Object.values(item);
                    sheet.addRow(rowData);
                });

                console.log(`Added data from file ${fileName} to Excel`);
            }
        } catch (error) {
            console.error(`Failed to read or process JSON file ${fileName}:`, error.message);
        }
    }

    await workbook.xlsx.writeFile(combinedFilePath);
    console.log(`Combined Excel saved at: ${combinedFilePath}`);
    return combinedFilePath;
}
// Main function to download JSON, combine, and provide download link
async function processAndProvideLink() {
    await downloadDataByIntervals();
    const combinedFile = await combineJsonFilesToLongExcel();
    console.log(`Download link: ${combinedFile}`);
}

processAndProvideLink().catch(console.error);
