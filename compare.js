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
      excelData.push(row.values);
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

          // Search for String in JSON Data
          const searchString = 'specificString';
          const offset = 3; // Number of characters after the string
          const length = 36; // Length of the data to extract

          const foundIndex = parsedJsonData.indexOf(searchString);

          if (foundIndex !== -1 && foundIndex + offset + length <= parsedJsonData.length) {
            // Extract Data for Comparison
            const extractedData = parsedJsonData.substring(foundIndex + offset, foundIndex + offset + length);

            // Initialize flag for match found
            let matchFound = false;

            for (let i = 0; i < excelData.length; i++) {
              const excelA = excelData[i][1]; // Assuming excel data is in column A
              if (extractedData.includes(excelA)) {
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
          } else {
            console.log(`String not found in ${jsonFile}`);
          }
        });
      });
    });
  })
  .catch(error => {
    console.error(`Error reading Excel file:`, error);
  });
