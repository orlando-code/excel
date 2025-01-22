function main(workbook: ExcelScript.Workbook) {
  // Get the currently selected cells.
  const selectedRange = workbook.getSelectedRange();
  const selectedValues = selectedRange.getValues();

  // Loop through the selected cells.
  selectedValues.forEach((row, rowIndex) => {
    row.forEach((cell, colIndex) => {
      if (typeof cell === "string") {
        let match: RegExpMatchArray | null; // Declare the type for 'match'.

        // Match the "a ± b" format.
        match = cell.match(/(-?\d+\.?\d*)\s*±\s*(-?\d+\.?\d*)/);
        if (match) {
          const value = parseFloat(match[1]);
          const uncertainty = parseFloat(match[2]);

          // Write the value into the current cell.
          const targetCell = selectedRange.getCell(rowIndex, colIndex);
          targetCell.setValue(value); // Overwrite the original cell with the value.

          // Write the uncertainty into the adjacent cell.
          targetCell.getOffsetRange(0, 1).setValue(uncertainty);
          return; // Exit current iteration as we've handled this case.
        }

        // Match the "a (b)" format.
        match = cell.match(/(-?\d+\.?\d*)\s*\(([^)]+)\)/);
        if (match) {
          const value = parseFloat(match[1]);
          const parentheticalValue = parseFloat(match[2]);

          // Write the value into the current cell.
          const targetCell = selectedRange.getCell(rowIndex, colIndex);
          targetCell.setValue(value); // Overwrite the original cell with the value.

          // Write the value in parentheses into the adjacent cell.
          targetCell.getOffsetRange(0, 1).setValue(parentheticalValue);
        }
      }
    });
  });
}
