function updatePublicSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const source = ss.getSheetByName("YOUR SOURCE SHEET");
  const publicSheet = ss.getSheetByName("YOUR PUBLIC SHEET");

  if (!source) {
    throw new Error('Your source sheet doesnt exist');
  }

  if (!publicSheet) {
    throw new Error('Your public sheet doesnt exist');
  }

  const sourceRows = source.getLastRow();
  const sourceColumns = source.getLastColumn();

  const data = source
    .getRange(1, 1, sourceRows, sourceColumns)
    .getValues();

  // Translates ALL text cells to Spanish
  const translatedData = data.map((row, rowIndex) => {
    const newRow = [...row];

    if (rowIndex > 0) { // Skip header row (row 0)
      for (let colIndex = 0; colIndex < newRow.length; colIndex++) {
        const cellValue = newRow[colIndex];
        
        if (typeof cellValue === "string" && cellValue.trim() !== "") {
          try {
            // Translate to Spanish using Google's Google Translate
            const translated = LanguageApp.translate(cellValue, '', 'es');
            newRow[colIndex] = translated;

            Utilities.sleep(1000);

          } catch (error) { // there may be errors
            console.log(`Failed to translate: "${cellValue}"`, error);
            newRow[colIndex] = cellValue;
          }
        }
      }
    }

    return newRow;
  });

  // Replace public sheet contents
  publicSheet.clearContents();

  publicSheet
    .getRange(1, 1, translatedData.length, translatedData[0].length)
    .setValues(translatedData);
}
