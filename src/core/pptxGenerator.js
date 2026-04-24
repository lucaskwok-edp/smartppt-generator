const path = require("path");
const fs = require("fs");
const { Automizer, modify } = require("pptx-automizer");

const { readExcelRows } = require("./excelReader");
const { detectTemplateSlide } = require("./templateDetector");
const {
  validateRequiredColumns,
  createRowValueResolver,
  createPlaceholderText,
} = require("./placeholderUtils");
const { createOutputPath } = require("./fileNameUtils");

const TEMPLATE_ALIAS = "template";
const ROOT_TEMPLATE_FILE = "root.pptx";

function getFileName(filePath) {
  return path.basename(filePath);
}

function getDirName(filePath) {
  return path.dirname(filePath);
}

function getProjectRootDir() {
  return path.join(__dirname, "..", "..");
}

function getRootTemplateDir() {
  return path.join(getProjectRootDir(), "templates");
}

function ensureRootTemplateExists() {
  const rootTemplatePath = path.join(getRootTemplateDir(), ROOT_TEMPLATE_FILE);

  if (!fs.existsSync(rootTemplatePath)) {
    throw new Error(
      `Arquivo root.pptx não encontrado. Crie um PowerPoint em branco e salve em: ${rootTemplatePath}`,
    );
  }

  return rootTemplatePath;
}

function createTextReplacements(placeholders, row) {
  const resolveValue = createRowValueResolver(row);

  return placeholders.map((placeholder) => {
    return {
      replace: createPlaceholderText(placeholder),
      by: {
        text: resolveValue(placeholder),
      },
    };
  });
}

async function replaceTextInAllSlideElements(slide, textReplacements) {
  const textElementIds = await slide.getAllTextElementIds();

  for (const elementId of textElementIds) {
    slide.modifyElement(elementId, modify.replaceText(textReplacements));
  }
}

async function addStaticSlidesBeforeTemplate(
  presentation,
  templateSlideNumber,
) {
  for (let slideNumber = 1; slideNumber < templateSlideNumber; slideNumber++) {
    presentation.addSlide(TEMPLATE_ALIAS, slideNumber);
  }
}

async function addGeneratedSlides(
  presentation,
  templateSlideNumber,
  placeholders,
  rows,
) {
  for (const row of rows) {
    const textReplacements = createTextReplacements(placeholders, row);

    presentation.addSlide(
      TEMPLATE_ALIAS,
      templateSlideNumber,
      async (slide) => {
        await replaceTextInAllSlideElements(slide, textReplacements);
      },
    );
  }
}

async function addStaticSlidesAfterTemplate(
  presentation,
  templateSlideNumber,
  totalSlides,
) {
  for (
    let slideNumber = templateSlideNumber + 1;
    slideNumber <= totalSlides;
    slideNumber++
  ) {
    presentation.addSlide(TEMPLATE_ALIAS, slideNumber);
  }
}

async function generatePresentation({ templatePath, excelPath, outputDir }) {
  if (!templatePath) {
    throw new Error("Selecione o template PowerPoint.");
  }

  if (!excelPath) {
    throw new Error("Selecione a planilha Excel.");
  }

  if (!outputDir) {
    throw new Error("Selecione a pasta de saída.");
  }

  ensureRootTemplateExists();

  const excelData = readExcelRows(excelPath);

  const detection = detectTemplateSlide(templatePath);
  const templateSlide = detection.templateSlide;

  validateRequiredColumns(templateSlide.placeholders, excelData.columns);

  const rootTemplateDir = getRootTemplateDir();

  const templateDir = getDirName(templatePath);
  const templateFileName = getFileName(templatePath);

  const outputPath = createOutputPath(outputDir);
  const outputFileName = getFileName(outputPath);

  const automizer = new Automizer({
    templateDir: rootTemplateDir,
    outputDir,
    removeExistingSlides: true,
  });

  const presentation = automizer.loadRoot(ROOT_TEMPLATE_FILE);

  presentation.load(templateDir, templateFileName, TEMPLATE_ALIAS);

  await addStaticSlidesBeforeTemplate(presentation, templateSlide.slideNumber);

  await addGeneratedSlides(
    presentation,
    templateSlide.slideNumber,
    templateSlide.placeholders,
    excelData.rows,
  );

  await addStaticSlidesAfterTemplate(
    presentation,
    templateSlide.slideNumber,
    detection.totalSlides,
  );

  await presentation.write(outputFileName);

  return {
    outputPath,
    rowsCount: excelData.rows.length,
    templateSlideNumber: templateSlide.slideNumber,
    placeholders: templateSlide.placeholders,
  };
}

module.exports = {
  generatePresentation,
};
