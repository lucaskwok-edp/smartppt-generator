const APP_NAME = "SmartPPT Generator";

const FILE_EXTENSIONS = {
  POWERPOINT: ".pptx",
  EXCEL: ".xlsx",
};

const OUTPUT_FILE = {
  PREFIX: "relatorio",
  EXTENSION: ".pptx",
};

const PLACEHOLDER = {
  REGEX: /\{([^{}]+)\}/g,
};

const DATE_FORMAT = {
  BR: "dd/mm/yyyy",
};

module.exports = {
  APP_NAME,
  FILE_EXTENSIONS,
  OUTPUT_FILE,
  PLACEHOLDER,
  DATE_FORMAT,
};
