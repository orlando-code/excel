function main(workbook: ExcelScript.Workbook) {
  // Get the currently selected range.
  const selectedRange = workbook.getSelectedRange();

  // Get the number of rows and columns in the selection.
  const rowCount = selectedRange.getRowCount();
  const columnCount = selectedRange.getColumnCount();

  // Loop through the selected range, row by row.
  for (let row = 0; row < rowCount; row++) {
    for (let col = columnCount - 1; col >= 0; col--) {
      const currentCell = selectedRange.getCell(row, col);
      const currentValue = currentCell.getValue();

      // If the cell contains data, insert a blank cell to the right without shifting the current cell.
      if (currentValue !== null && currentValue !== "") {
        const targetCell = currentCell.getOffsetRange(0, 1);

        // Insert a blank cell at the target location, shifting cells to the right of it.
        targetCell.insert(ExcelScript.InsertShiftDirection.right);
      }
    }
  }

  console.log("Blank cells inserted to the right of cells containing data without moving the leftmost cell.");
}
