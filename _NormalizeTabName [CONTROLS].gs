function normalizeTabName(value) {
  if (value instanceof Date) {
    const day = String(value.getDate()).padStart(2, "0");
    const month = String(value.getMonth() + 1).padStart(2, "0");
    const year = value.getFullYear();
    return day + "-" + month + "-" + year;
  }
  return value.toString().trim();
}
