const { PLACEHOLDER } = require("../shared/constants");
const { formatDateToBR } = require("./dateUtils");

function normalizeKey(value) {
  return String(value || "")
    .trim()
    .toLowerCase();
}

function createPlaceholderText(columnName) {
  return `{${PLACEHOLDER.PREFIX}${columnName}}`;
}

function extractPlaceholdersFromText(text) {
  if (!text) {
    return [];
  }

  const placeholders = [];
  const regex = new RegExp(PLACEHOLDER.REGEX);

  let match;

  while ((match = regex.exec(text)) !== null) {
    const placeholderName = String(match[1] || "").trim();

    if (placeholderName) {
      placeholders.push(placeholderName);
    }
  }

  return placeholders;
}

function extractUniquePlaceholdersFromTexts(texts) {
  const unique = new Map();

  for (const text of texts) {
    const placeholders = extractPlaceholdersFromText(text);

    for (const placeholder of placeholders) {
      const normalized = normalizeKey(placeholder);

      if (!unique.has(normalized)) {
        unique.set(normalized, placeholder);
      }
    }
  }

  return Array.from(unique.values());
}

function createRowValueResolver(row) {
  const normalizedRow = {};

  for (const [key, value] of Object.entries(row)) {
    normalizedRow[normalizeKey(key)] = value;
  }

  return function resolveValue(placeholderName) {
    const normalizedPlaceholder = normalizeKey(placeholderName);
    const value = normalizedRow[normalizedPlaceholder];

    if (value === null || value === undefined) {
      return "";
    }

    if (value instanceof Date) {
      return formatDateToBR(value);
    }

    return String(value);
  };
}

function replacePlaceholdersInText(text, row) {
  if (!text) {
    return "";
  }

  const resolveValue = createRowValueResolver(row);

  return String(text).replace(PLACEHOLDER.REGEX, (_match, placeholderName) => {
    return resolveValue(placeholderName);
  });
}

function validateRequiredColumns(placeholders, columns) {
  const normalizedColumns = new Set(
    columns.map((column) => normalizeKey(column)),
  );

  const missingColumns = placeholders.filter((placeholder) => {
    return !normalizedColumns.has(normalizeKey(placeholder));
  });

  if (missingColumns.length > 0) {
    throw new Error(
      `A planilha não possui as seguintes colunas exigidas pelo template: ${missingColumns.join(
        ", ",
      )}`,
    );
  }
}

module.exports = {
  normalizeKey,
  createPlaceholderText,
  extractPlaceholdersFromText,
  extractUniquePlaceholdersFromTexts,
  createRowValueResolver,
  replacePlaceholdersInText,
  validateRequiredColumns,
};
