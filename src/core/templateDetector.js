const AdmZip = require("adm-zip");
const { extractUniquePlaceholdersFromTexts } = require("./placeholderUtils");

function decodeXmlText(value) {
  return String(value || "")
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'");
}

function getEntryText(zip, entryName) {
  const entry = zip.getEntry(entryName);

  if (!entry) {
    return "";
  }

  return entry.getData().toString("utf8");
}

function extractSlideRelationIds(presentationXml) {
  const ids = [];
  const regex = /<p:sldId\b[^>]*r:id="([^"]+)"/g;

  let match;

  while ((match = regex.exec(presentationXml)) !== null) {
    ids.push(match[1]);
  }

  return ids;
}

function extractPresentationRelationships(relsXml) {
  const relationships = new Map();

  const regex = /<Relationship\b[^>]*Id="([^"]+)"[^>]*Target="([^"]+)"[^>]*>/g;

  let match;

  while ((match = regex.exec(relsXml)) !== null) {
    const id = match[1];
    const target = match[2];

    relationships.set(id, target);
  }

  return relationships;
}

function normalizeSlideTargetToPath(target) {
  let normalizedTarget = String(target || "").replace(/^\/+/, "");

  if (normalizedTarget.startsWith("ppt/")) {
    return normalizedTarget;
  }

  return `ppt/${normalizedTarget}`;
}

function getOrderedSlidePaths(zip) {
  const presentationXml = getEntryText(zip, "ppt/presentation.xml");
  const relsXml = getEntryText(zip, "ppt/_rels/presentation.xml.rels");

  if (!presentationXml) {
    throw new Error(
      "Não foi possível ler o arquivo ppt/presentation.xml do template.",
    );
  }

  if (!relsXml) {
    throw new Error(
      "Não foi possível ler o arquivo ppt/_rels/presentation.xml.rels do template.",
    );
  }

  const slideRelationIds = extractSlideRelationIds(presentationXml);
  const relationships = extractPresentationRelationships(relsXml);

  const slidePaths = slideRelationIds
    .map((relationId) => relationships.get(relationId))
    .filter(Boolean)
    .map(normalizeSlideTargetToPath);

  if (slidePaths.length === 0) {
    throw new Error("Nenhum slide foi encontrado no template PowerPoint.");
  }

  return slidePaths;
}

function extractTextRunsFromXml(xml) {
  const texts = [];
  const regex = /<a:t[^>]*>([\s\S]*?)<\/a:t>/g;

  let match;

  while ((match = regex.exec(xml)) !== null) {
    texts.push(decodeXmlText(match[1]));
  }

  return texts;
}

function extractParagraphTextsFromSlideXml(slideXml) {
  const paragraphTexts = [];
  const paragraphRegex = /<a:p\b[^>]*>([\s\S]*?)<\/a:p>/g;

  let paragraphMatch;

  while ((paragraphMatch = paragraphRegex.exec(slideXml)) !== null) {
    const paragraphXml = paragraphMatch[1];
    const runs = extractTextRunsFromXml(paragraphXml);
    const paragraphText = runs.join("");

    if (paragraphText.trim()) {
      paragraphTexts.push(paragraphText);
    }
  }

  return paragraphTexts;
}

function extractTextsFromSlideXml(slideXml) {
  const individualRuns = extractTextRunsFromXml(slideXml);
  const paragraphTexts = extractParagraphTextsFromSlideXml(slideXml);
  const fullSlideText = individualRuns.join("");

  return [...individualRuns, ...paragraphTexts, fullSlideText];
}

function inspectSlides(templatePath) {
  if (!templatePath) {
    throw new Error("O caminho do template PowerPoint não foi informado.");
  }

  const zip = new AdmZip(templatePath);
  const slidePaths = getOrderedSlidePaths(zip);

  return slidePaths.map((slidePath, index) => {
    const slideXml = getEntryText(zip, slidePath);

    if (!slideXml) {
      throw new Error(`Não foi possível ler o slide: ${slidePath}`);
    }

    const texts = extractTextsFromSlideXml(slideXml);
    const placeholders = extractUniquePlaceholdersFromTexts(texts);

    return {
      index,
      slideNumber: index + 1,
      slidePath,
      texts,
      placeholders,
    };
  });
}

function detectTemplateSlide(templatePath) {
  const slides = inspectSlides(templatePath);

  const slidesWithPlaceholders = slides.filter((slide) => {
    return slide.placeholders.length > 0;
  });

  if (slidesWithPlaceholders.length === 0) {
    throw new Error(
      "Nenhum slide com placeholders foi encontrado. Use placeholders no formato {PPT_NomeDaColuna}.",
    );
  }

  if (slidesWithPlaceholders.length > 1) {
    const slideNumbers = slidesWithPlaceholders
      .map((slide) => slide.slideNumber)
      .join(", ");

    throw new Error(
      `Mais de um slide possui placeholders do sistema. Mantenha placeholders {PPT_...} em apenas um slide-modelo. Slides encontrados: ${slideNumbers}.`,
    );
  }

  return {
    templateSlide: slidesWithPlaceholders[0],
    slides,
    totalSlides: slides.length,
  };
}

module.exports = {
  inspectSlides,
  detectTemplateSlide,
};
