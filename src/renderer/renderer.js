const templateInput = document.getElementById("templatePath");
const excelInput = document.getElementById("excelPath");
const outputInput = document.getElementById("outputDir");

const btnSelectTemplate = document.getElementById("btnSelectTemplate");
const btnSelectExcel = document.getElementById("btnSelectExcel");
const btnSelectOutput = document.getElementById("btnSelectOutput");
const btnGenerate = document.getElementById("btnGenerate");

const statusBox = document.getElementById("statusBox");

function setStatus(message, type = "muted") {
  statusBox.textContent = message;

  statusBox.classList.remove(
    "status-muted",
    "status-info",
    "status-success",
    "status-error",
  );

  statusBox.classList.add(`status-${type}`);
}

function setLoading(isLoading) {
  btnGenerate.disabled = isLoading;
  btnSelectTemplate.disabled = isLoading;
  btnSelectExcel.disabled = isLoading;
  btnSelectOutput.disabled = isLoading;

  btnGenerate.textContent = isLoading ? "Gerando..." : "Gerar PowerPoint";
}

function validateForm() {
  if (!templateInput.value) {
    setStatus("Selecione um template PowerPoint (.pptx).", "error");
    return false;
  }

  if (!excelInput.value) {
    setStatus("Selecione uma planilha Excel (.xlsx).", "error");
    return false;
  }

  if (!outputInput.value) {
    setStatus("Selecione uma pasta de saída.", "error");
    return false;
  }

  return true;
}

async function loadDefaultOutputDir() {
  try {
    const defaultOutputDir = await window.smartPpt.getDefaultOutputDir();

    if (defaultOutputDir) {
      outputInput.value = defaultOutputDir;
    }
  } catch (error) {
    setStatus(
      "Não foi possível carregar a pasta padrão de saída. Selecione manualmente.",
      "error",
    );
  }
}

btnSelectTemplate.addEventListener("click", async () => {
  const selectedPath = await window.smartPpt.selectPptx();

  if (selectedPath) {
    templateInput.value = selectedPath;
    setStatus("Template PowerPoint selecionado.", "info");
  }
});

btnSelectExcel.addEventListener("click", async () => {
  const selectedPath = await window.smartPpt.selectXlsx();

  if (selectedPath) {
    excelInput.value = selectedPath;
    setStatus("Planilha Excel selecionada.", "info");
  }
});

btnSelectOutput.addEventListener("click", async () => {
  const selectedPath = await window.smartPpt.selectOutputDir();

  if (selectedPath) {
    outputInput.value = selectedPath;
    setStatus("Pasta de saída selecionada.", "info");
  }
});

btnGenerate.addEventListener("click", async () => {
  if (!validateForm()) {
    return;
  }

  setLoading(true);
  setStatus("Gerando apresentação. Aguarde...", "info");

  try {
    const result = await window.smartPpt.generatePresentation({
      templatePath: templateInput.value,
      excelPath: excelInput.value,
      outputDir: outputInput.value,
    });

    if (!result.success) {
      setStatus(result.message || "Erro ao gerar o PowerPoint.", "error");
      return;
    }

    setStatus(`PowerPoint gerado com sucesso: ${result.outputPath}`, "success");
  } catch (error) {
    setStatus(
      error.message || "Erro inesperado ao gerar o PowerPoint.",
      "error",
    );
  } finally {
    setLoading(false);
  }
});

loadDefaultOutputDir();
