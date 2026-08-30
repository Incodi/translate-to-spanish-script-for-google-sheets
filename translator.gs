function updatePublicSheet() {
  const spreadsheetID = "YOUR SPREADSHEET ID";
  
  const ss = SpreadsheetApp.openById(spreadsheetID);
  const source = ss.getSheetByName("YOUR SOURCE SHEET");
  const publicSheet = ss.getSheetByName("YOUR PUBLIC SHEET");
  const language = 'es'

  if (!source) {
    throw new Error('Your source sheet doesnt exist');
  }

  if (!publicSheet) {
    throw new Error('Your public sheet doesnt exist');
  }

  const sourceRows = source.getLastRow();
  const sourceColumns = source.getLastColumn();
  const data = source.getRange(1, 1, sourceRows, sourceColumns).getValues();

  const translatedData = [];
  let translationCount = 0;
  const translatePerRun = 300;
  const batchSize = 25;
  const pauseTime = 2000;

  for (let rowIndex = 0; rowIndex < data.length; rowIndex++) {
    const row = data[rowIndex];
    const newRow = [...row];

    if (rowIndex > 0) {
      for (let colIndex = 0; colIndex < newRow.length; colIndex++) {
        const cellValue = newRow[colIndex];
        
        if (typeof cellValue === "string" && cellValue.trim() !== "") {
          if (translationCount >= translatePerRun) {
            console.log(`Reached translation limit (${translatePerRun}). Please run again for remaining cells.`);
            newRow[colIndex] = cellValue + " [NEEDS TRANSLATION]";
            continue;
          }


          const translated = LanguageApp.translate(cellValue, '', language);
          newRow[colIndex] = translated;
          translationCount++;

          if (translationCount % batchSize === 0) {
            console.log(`Translated ${translationCount} items. Pausing for ${pauseTime/1000} seconds...`);
            Utilities.sleep(pauseTime);
          }
        }
      }
    }

    translatedData.push(newRow);
  }

  publicSheet.clearContents();
  publicSheet.getRange(1, 1, translatedData.length, translatedData[0].length).setValues(translatedData);

  console.log(`Translation complete. ${translationCount} items translated.`);
  
  if (translationCount < countTotalTranslations(data)) {
    console.log('Not all items were translated. Run the script again to translate remaining items.');
  }
}

function countTotalTranslations(data) {
  let count = 0;
  for (let rowIndex = 1; rowIndex < data.length; rowIndex++) {
    const row = data[rowIndex];
    for (let colIndex = 0; colIndex < row.length; colIndex++) {
      const cellValue = row[colIndex];
      if (typeof cellValue === "string" && cellValue.trim() !== "") {
        count++;
      }
    }
  }
  return count;
}
