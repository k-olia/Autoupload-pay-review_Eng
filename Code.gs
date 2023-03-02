var RESULT_TAB = "Master tab";
var INPUT_FOLDER_ID = "";
var OUTPUT_FILE_ID = "";

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Update data")
    .addItem("Copy", "copy")
    .addToUi();
}

function copy() {
  const folder = DriveApp.getFolderById(INPUT_FOLDER_ID);
  const files = folder.getFiles();
  const outputFile = DriveApp.getFileById(OUTPUT_FILE_ID);
  const outputSS = SpreadsheetApp.open(outputFile);

  while (files.hasNext()) {
    const file = files.next();
    processFile(file, outputSS);
  }
  const notations = [
    {
      letter: "O",
      startIndex: 17,
      endIndex: 731,
    },
    {
      letter: "P",
      startIndex: 17,
      endIndex: 731,
    },
    {
      letter: "Q",
      startIndex: 17,
      endIndex: 731,
    },
    {
      letter: "AN",
      startIndex: 17,
      endIndex: 731,
    },
  ];

  var firstSheetInOutputSS = outputSS.getSheetByName(RESULT_TAB); //.getSheets()[0];
  reevaluateFormulas(firstSheetInOutputSS, notations);
}

function processAdditionalCells() {
  const outputFile = DriveApp.getFileById(OUTPUT_FILE_ID);
  const outputSS = SpreadsheetApp.open(outputFile);
  var firstSheetInOutputSS = outputSS.getSheetByName(RESULT_TAB); //.getSheets()[0];
  const notations = [
    {
      letter: "AO",
      startIndex: 17,
      endIndex: 731,
    },
    {
      letter: "BD",
      startIndex: 17,
      endIndex: 731,
    },
    {
      letter: "BE",
      startIndex: 17,
      endIndex: 731,
    },
    {
      letter: "BF",
      startIndex: 17,
      endIndex: 731,
    },
  ];
  reevaluateFormulas(firstSheetInOutputSS, notations);
}
function processFile(file, outputSS) {
  console.log("Processing file " + file.getName());

  const ss = SpreadsheetApp.open(file);
  const sheet = ss.getSheets()[0];
  let sheetName = file.getName();
  sheet.setName(sheetName);
  sheetName = "Copy of " + sheet.getName();

  var sheetInOutputSS = outputSS.getSheetByName(sheetName);

  console.log(`Deleting sheet '${sheetName}' from the '${RESULT_TAB}'`);
  if (sheetInOutputSS) {
    outputSS.deleteSheet(sheetInOutputSS);
  }
  console.log("Deleted");

  console.log(`Inserting sheet into the '${RESULT_TAB}'`);
  sheet.copyTo(outputSS);
  console.log("Inserted");
}

function reevaluateFormulas(sheet, notations) {
  // const notations = [
  //   {
  //     letter: 'O',
  //     startIndex: 17,
  //     endIndex: 731,
  //   },
  //   {
  //     letter: 'P',
  //     startIndex: 17,
  //     endIndex: 731,
  //   },
  //   {
  //     letter: 'Q',
  //     startIndex: 17,
  //     endIndex: 731,
  //   },
  //   {
  //     letter: 'AN',
  //     startIndex: 17,
  //     endIndex: 731,
  //   },
  //   {
  //     letter: 'AO',
  //     startIndex: 17,
  //     endIndex: 731,
  //   },
  //   {
  //     letter: 'BD',
  //     startIndex: 17,
  //     endIndex: 731,
  //   },
  //  {
  //     letter: 'BE',
  //     startIndex: 17,
  //     endIndex: 731,
  //   },
  //  {
  //     letter: 'BF',
  //     startIndex: 17,
  //     endIndex: 731,
  //   },
  // ];

  for (let j = 0; j < notations.length; j++) {
    const notation = notations[j];
    console.log("Processing range " + getNotation(notation));
    const range = sheet.getRange(getNotation(notation));
    for (let i = 1; i <= notation.endIndex - notation.startIndex + 1; i++) {
      console.log(`\tProcessing cell ${i} in this range`);
      const cell = range.getCell(i, 1);
      const formula = cell.getFormula();
      cell.setFormula(formula);
    }
  }
}

function getNotation(object) {
  return `${object.letter}${object.startIndex}:${object.letter}${object.endIndex}`;
}
