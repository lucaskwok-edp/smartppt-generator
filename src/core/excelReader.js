const XLSX = require("xlsx");

function readExcelRows(excelPath) {
  if (!excelPath) {
    throw new Error("O caminho da planilha Excel não foi informado.");
  }

  const workbook = XLSX.readFile(excelPath, {
    cellDates: true,
  });

  const firstSheetName = workbook.SheetNames[0];

  if (!firstSheetName) {
    throw new Error("A planilha Excel não possui nenhuma aba.");
  }

  const worksheet = workbook.Sheets[firstSheetName];

  const rows = XLSX.utils.sheet_to_json(worksheet, {
    defval: "",
    raw: true,
  });

  if (rows.length === 0) {
    throw new Error("A planilha Excel não possui linhas de dados.");
  }

  const columns = Object.keys(rows[0] || {});

  if (columns.length === 0) {
    throw new Error("A planilha Excel não possui colunas.");
  }

  return {
    sheetName: firstSheetName,
    columns,
    rows,
  };
}

module.exports = {
  readExcelRows,
};
