const fs = require("fs");
const Excel = require("exceljs");

// Define the directory path and Excel file name
const directoryPath = "C:/Users/Hao/OneDrive/Documents/data/";
const excelPath = "C:/Users/Hao/OneDrive/Documents/data/guid.xlsx";

// Read Excel Data
const workbook = new Excel.Workbook();
workbook.xlsx
  .readFile(excelPath)
  .then(() => {
    const worksheet = workbook.getWorksheet(1);
    const excelData = [];

    worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
      excelData.push(row.values);
    });

    // Read the directory contents
    fs.readdir(directoryPath, (err, files) => {
      if (err) {
        console.error("Error reading directory:", err);
        return;
      }

      // Filter out JSON files
      const jsonFiles = files.filter((file) => file.endsWith(".json"));

      // Process each JSON file
      jsonFiles.forEach((jsonFile) => {
        const jsonFilePath = `${directoryPath}${jsonFile}`;

        fs.readFile(jsonFilePath, "utf8", (err, jsonData) => {
          if (err) {
            console.error(`Error reading file ${jsonFilePath}:`, err);
            return;
          }

          // Process JSON data
          const parsedJsonData = JSON.parse(jsonData);

          // Compare Data
          const jsonLine1 = parsedJsonData[0]; // Assuming line1 is an array in the JSON

          for (let i = 0; i < excelData.length; i++) {
            const excelA = excelData[i][1]; // Assuming excel data is in column A
            if (jsonLine1 === excelA) {
              // Write JSON file name to Excel B
              worksheet.getCell(`B${i + 1}`).value = jsonFile;
              break; // No need to continue checking
            }
          }

          // Save Excel file
          workbook.xlsx
            .writeFile("output.xlsx")
            .then(() => {
              console.log("Excel file saved with updated data.");
            })
            .catch((error) => {
              console.error(`Error writing Excel file:`, error);
            });
        });
      });
    });
  })
  .catch((error) => {
    console.error(`Error reading Excel file:`, error);
  });
