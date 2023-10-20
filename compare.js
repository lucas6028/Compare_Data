const fs = require('fs');
const Excel = require('exceljs');

// Define the directory path and Excel file name
const directoryPath = 'C:/Users/patricia7909/Desktop/Test/data/';
const excelFileName = 'C:/Users/patricia7909/Desktop/Test/test1.xlsx';

// Read Excel Data
const workbook = new Excel.Workbook();
workbook.xlsx.readFile(excelFileName)
  .then(() => {
    const worksheet = workbook.getWorksheet(1);
    const excelData = [];

    worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
      excelData.push(row.getCell(1).value); // Assuming excel data is in column A
    });

    // Read the directory contents
    fs.readdir(directoryPath, (err, files) => {
      if (err) {
        console.error('Error reading directory:', err);
        return;
      }

      // Filter out JSON files
      const jsonFiles = files.filter(file => file.endsWith('.json'));

      // Process each JSON file
      jsonFiles.forEach(jsonFile => {
        const jsonFilePath = `${directoryPath}${jsonFile}`;

        fs.readFile(jsonFilePath, 'utf8', (err, jsonData) => {
          if (err) {
            console.error(`Error reading file ${jsonFilePath}:`, err);
            return;
          }

          // Process JSON data
          const parsedJsonData = JSON.parse(jsonData);

          // Extract the value associated with the key "GUID"
          const jsonGUID = parsedJsonData.GUID;

          // Initialize flag for match found
          let matchFound = false;

          for (let i = 0; i < excelData.length; i++) {
            const excelA = excelData[i];
            if (excelA === jsonGUID) {
              // Write JSON file name to Excel B
              worksheet.getCell(`B${i + 1}`).value = jsonFile;
              matchFound = true;
              break; // Match found, no need to continue checking
            }
          }

          // If no match found, consider other comparison logic here
          if (!matchFound) {
            // Implement logic to compare other cells or handle no match scenario
            console.log(`No match found for ${jsonFile}`);
          }

          // Save Excel file
          workbook.xlsx.writeFile('output.xlsx')
            .then(() => {
              console.log('Excel file saved with updated data.');
            })
            .catch(error => {
              console.error(`Error writing Excel file:`, error);
            });
        });
      });
    });
  })
  .catch(error => {
    console.error(`Error reading Excel file:`, error);
  });
