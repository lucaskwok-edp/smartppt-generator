const { app, BrowserWindow, dialog, ipcMain } = require("electron");
const path = require("path");
const os = require("os");

function createWindow() {
  const mainWindow = new BrowserWindow({
    width: 760,
    height: 520,
    minWidth: 680,
    minHeight: 460,
    resizable: true,
    title: "SmartPPT Generator",
    autoHideMenuBar: true,
    webPreferences: {
      preload: path.join(__dirname, "preload.js"),
      contextIsolation: true,
      nodeIntegration: false,
    },
  });

  mainWindow.loadFile(path.join(__dirname, "../renderer/index.html"));

  if (process.env.NODE_ENV === "development") {
    mainWindow.webContents.openDevTools();
  }
}

function getDesktopPath() {
  return path.join(os.homedir(), "Desktop");
}

ipcMain.handle("dialog:select-pptx", async () => {
  const result = await dialog.showOpenDialog({
    title: "Selecionar template PowerPoint",
    properties: ["openFile"],
    filters: [
      {
        name: "PowerPoint",
        extensions: ["pptx"],
      },
    ],
  });

  if (result.canceled || result.filePaths.length === 0) {
    return null;
  }

  return result.filePaths[0];
});

ipcMain.handle("dialog:select-xlsx", async () => {
  const result = await dialog.showOpenDialog({
    title: "Selecionar planilha Excel",
    properties: ["openFile"],
    filters: [
      {
        name: "Excel",
        extensions: ["xlsx"],
      },
    ],
  });

  if (result.canceled || result.filePaths.length === 0) {
    return null;
  }

  return result.filePaths[0];
});

ipcMain.handle("dialog:select-output-dir", async () => {
  const result = await dialog.showOpenDialog({
    title: "Selecionar pasta de saída",
    defaultPath: getDesktopPath(),
    properties: ["openDirectory"],
  });

  if (result.canceled || result.filePaths.length === 0) {
    return null;
  }

  return result.filePaths[0];
});

ipcMain.handle("app:get-default-output-dir", async () => {
  return getDesktopPath();
});

ipcMain.handle("pptx:generate", async (_event, payload) => {
  try {
    const { generatePresentation } = require("../core/pptxGenerator");

    const result = await generatePresentation({
      templatePath: payload.templatePath,
      excelPath: payload.excelPath,
      outputDir: payload.outputDir,
    });

    return {
      success: true,
      outputPath: result.outputPath,
    };
  } catch (error) {
    return {
      success: false,
      message: error.message || "Erro inesperado ao gerar o PowerPoint.",
    };
  }
});

app.whenReady().then(() => {
  createWindow();

  app.on("activate", () => {
    if (BrowserWindow.getAllWindows().length === 0) {
      createWindow();
    }
  });
});

app.on("window-all-closed", () => {
  if (process.platform !== "darwin") {
    app.quit();
  }
});
