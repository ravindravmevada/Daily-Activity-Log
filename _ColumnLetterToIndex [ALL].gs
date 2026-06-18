function columnLetterToIndex(columnLetter) {
  let columnIndex = 0;
  for (let i = 0; i < columnLetter.length; i++) {
    const ascii = columnLetter.charCodeAt(i);
    const letterValue = ascii - 64;
    columnIndex = columnIndex * 26 + letterValue;
  }
  return columnIndex;
}
