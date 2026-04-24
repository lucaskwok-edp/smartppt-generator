const path = require("path");
const { OUTPUT_FILE } = require("../shared/constants");
const { getTimestampForFileName } = require("./dateUtils");

function createOutputFileName() {
  const timestamp = getTimestampForFileName();

  return `${OUTPUT_FILE.PREFIX}_${timestamp}${OUTPUT_FILE.EXTENSION}`;
}

function createOutputPath(outputDir) {
  if (!outputDir) {
    throw new Error("A pasta de saída não foi informada.");
  }

  const fileName = createOutputFileName();

  return path.join(outputDir, fileName);
}

module.exports = {
  createOutputFileName,
  createOutputPath,
};
