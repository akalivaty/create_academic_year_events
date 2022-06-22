function debug() {
  const ss = SpreadsheetApp.getActiveSheet();
  const data = ss.getRange(2, 2, ss.getLastRow() - 1, 2).getValues();
  data.forEach((value, idx) => {
    console.log(`row: ${idx + 1}, value: ${value[0]}, type: ${typeof (value[0])}, length: ${value[0].split('-').length}`);
  })
}
