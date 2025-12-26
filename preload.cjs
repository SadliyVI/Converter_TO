const { contextBridge, ipcRenderer } = require("electron");

contextBridge.exposeInMainWorld("api", {
    selectPdf: () => ipcRenderer.invoke("select-pdf"),
    selectOutDir: () => ipcRenderer.invoke("select-outdir"),
    startConvert: (payload) => ipcRenderer.invoke("start-convert", payload),

    openFolder: (folderPath) => ipcRenderer.invoke("open-folder", folderPath),
    showItemInFolder: (filePath) => ipcRenderer.invoke("show-item-in-folder", filePath),
    openFile: (filePath) => ipcRenderer.invoke("open-file", filePath),

    onProgress: (cb) => ipcRenderer.on("progress", (_e, data) => cb(data))
});