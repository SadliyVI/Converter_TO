import { app, BrowserWindow, ipcMain, dialog, shell } from "electron";
import path from "path";
import { fileURLToPath } from "url";
import { convertPdfToXlsx } from "./converter.js";
import fs from "fs";
import os from "os";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

let win;

function createWindow() {
    
    const preloadPath = path.join(app.getAppPath(), "preload.cjs");
    
    win = new BrowserWindow({
        width: 820,
        height: 740,
        title: "Конвертер TO",
        webPreferences: {
            // preload: path.join(__dirname, "preload.cjs"),
            preload: preloadPath,
            contextIsolation: true,
            nodeIntegration: false
        }
    });



    win.loadFile(path.join(__dirname, "renderer", "index.html"));
    // win.webContents.openDevTools({ mode: "detach" });
    logToFile("preloadPath=" + preloadPath);
    logToFile("exists(preload)=" + fs.existsSync(preloadPath));
}

app.whenReady().then(() => {
    createWindow();
    app.on("activate", () => {
        if (BrowserWindow.getAllWindows().length === 0) createWindow();
    });
});

app.on("window-all-closed", () => {
    if (process.platform !== "darwin") app.quit();
});

function logToFile(msg) {
    try {
        const p = path.join(os.homedir(), "to-converter.log");
        fs.appendFileSync(p, `[${new Date().toISOString()}] ${msg}\n`, "utf8");
    } catch { }
}

/* ---------- IPC ---------- */

ipcMain.handle("select-pdf", async () => {
    const res = await dialog.showOpenDialog(win, {
        title: "Выберите исходный PDF",
        properties: ["openFile"],
        filters: [{ name: "PDF", extensions: ["pdf"] }]
    });
    if (res.canceled || !res.filePaths?.[0]) return null;
    return res.filePaths[0];
});

ipcMain.handle("select-outdir", async () => {
    const res = await dialog.showOpenDialog(win, {
        title: "Выберите папку для сохранения",
        properties: ["openDirectory"]
    });
    if (res.canceled || !res.filePaths?.[0]) return null;
    return res.filePaths[0];
});

ipcMain.handle("start-convert", async (_evt, payload) => {
    const { pdfPath, outDir, outName } = payload || {};
    if (!pdfPath) throw new Error("Не выбран исходный PDF");
    if (!outDir) throw new Error("Не выбрана папка для сохранения");
    if (!outName || !String(outName).trim()) throw new Error("Не задано имя выходного файла");

    const cleanName = String(outName).trim().replace(/[\\/:*?"<>|]/g, "_");
    const outPath = path.join(outDir, `${cleanName}.xlsx`);

    const resultPath = await convertPdfToXlsx(pdfPath, outPath, (progress) => {
        win?.webContents.send("progress", progress);
    });

    return resultPath;
});

ipcMain.handle("open-folder", async (_evt, folderPath) => {
    if (!folderPath) return false;
    await shell.openPath(folderPath);
    return true;
});

ipcMain.handle("show-item-in-folder", async (_evt, filePath) => {
    if (!filePath) return false;
    shell.showItemInFolder(filePath);
    return true;
});

ipcMain.handle("open-file", async (_evt, filePath) => {
    if (!filePath) return false;
    await shell.openPath(filePath);
    return true;
});

process.on("uncaughtException", (e) => logToFile("uncaughtException: " + (e?.stack || e)));
process.on("unhandledRejection", (e) => logToFile("unhandledRejection: " + (e?.stack || e)));

app.on("ready", () => logToFile("App ready. appPath=" + app.getAppPath()));
