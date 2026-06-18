function getFormattedDateTime() {
  const now = new Date();
  const dayName = now.toLocaleString("en-US", { weekday: "long" });
  const day = now.getDate();
  const month = now.toLocaleString("en-US", { month: "long" });
  const year = now.getFullYear();
  const time = now.toLocaleTimeString("en-US", {
    hour: "2-digit", minute: "2-digit", second: "2-digit", hour12: true
  });
  return `${dayName}, ${day} ${month} ${year} - ${time}`;
}
