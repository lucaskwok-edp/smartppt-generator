const { contextBridge, ipcRenderer } = require("electron");

contextBridge.exposeInMainWorld("smartPpt", {
  selectPptx: () => ipcRenderer.invoke("dialog:select-pptx"),

  selectXlsx: () => ipcRenderer.invoke("dialog:select-xlsx"),

  selectOutputDir: () => ipcRenderer.invoke("dialog:select-output-dir"),

  getDefaultOutputDir: () => ipcRenderer.invoke("app:get-default-output-dir"),

  generatePresentation: ({ templatePath, excelPath, outputDir }) =>
    ipcRenderer.invoke("pptx:generate", {
      templatePath,
      excelPath,
      outputDir,
    }),
});
