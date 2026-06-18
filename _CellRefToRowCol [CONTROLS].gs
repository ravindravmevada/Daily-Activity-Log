function cellRefToRowCol(cellRef) {
  const match = cellRef.match(/^([A-Z]+)(\d+)$/);
  const colLetters = match[1];
  const rowNumber = parseInt(match[2], 10);
  return { row: rowNumber, col: columnLetterToIndex(colLetters) };
}
