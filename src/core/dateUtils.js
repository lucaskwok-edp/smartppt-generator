function pad2(value) {
  return String(value).padStart(2, "0");
}

function isValidDate(value) {
  return value instanceof Date && !Number.isNaN(value.getTime());
}

function formatDateToBR(value) {
  if (!value) {
    return "";
  }

  if (isValidDate(value)) {
    const day = pad2(value.getDate());
    const month = pad2(value.getMonth() + 1);
    const year = value.getFullYear();

    return `${day}/${month}/${year}`;
  }

  return String(value);
}

function getTimestampForFileName(date = new Date()) {
  const year = date.getFullYear();
  const month = pad2(date.getMonth() + 1);
  const day = pad2(date.getDate());
  const hour = pad2(date.getHours());
  const minute = pad2(date.getMinutes());
  const second = pad2(date.getSeconds());

  return `${year}${month}${day}_${hour}${minute}${second}`;
}

module.exports = {
  formatDateToBR,
  getTimestampForFileName,
};
