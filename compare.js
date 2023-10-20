// ...
try {
  // Parse JSON data
  const parsedJsonData = JSON.parse(jsonData);

  // Extract the GUID
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
