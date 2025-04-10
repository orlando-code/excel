function main(workbook: ExcelScript.Workbook) {
  // Get the currently selected range
  const selectedRange = workbook.getActiveCell().getSurroundingRegion();
  const values = selectedRange.getValues();

  const startRow = selectedRange.getRowIndex();
  const startCol = selectedRange.getColumnIndex();

  // Loop through each row of the selected range
  values.forEach((row, rowIndex) => {
    const cellValue = row[0];
    if (typeof cellValue === 'string') {
      const splitValues = cellValue.split(',').map(v => v.trim());
      const targetRange = selectedRange
        .getWorksheet()
        .getRangeByIndexes(startRow + rowIndex, startCol + 1, 1, splitValues.length);
      targetRange.setValues([splitValues]);
    }
  });
}
