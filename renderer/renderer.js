// console.log("window.api =", window.api);

let state = {
    pdfPath: null,
    outDir: null,
    confirmedName: null,
    busy: false,
    outPath: null,
    verbose: false
};

const $ = (id) => document.getElementById(id);

function canStart() {
    return !!(state.pdfPath && state.outDir && state.confirmedName);
}

function logLine(s) {
    const el = $("log");
    const line = `[${new Date().toLocaleTimeString()}] ${s}`;
    el.textContent = (el.textContent ? el.textContent + "\n" : "") + line;
    el.scrollTop = el.scrollHeight;
}

function setBusy(v) {
    state.busy = v;
    $("btnPdf").disabled = v;
    $("btnDir").disabled = v;
    $("btnConfirm").disabled = v;
    $("btnStart").disabled = v || !canStart();
    $("outName").disabled = v;

    // кнопки открытия доступны только если есть outPath и не busy
    $("btnShowFile").disabled = v || !state.outPath;
    $("btnOpenFile").disabled = v || !state.outPath;
    $("btnOpenFolder").disabled = v || !state.outDir;
}

function updateUI() {
    $("pdfHint").textContent = state.pdfPath ? state.pdfPath : "Файл не выбран";
    $("dirHint").textContent = state.outDir ? state.outDir : "Папка не выбрана";
    $("nameHint").textContent = state.confirmedName
        ? `Файл будет сохранён как: ${state.confirmedName}.xlsx`
        : "Имя не задано";

    $("btnStart").disabled = state.busy || !canStart();
    $("btnShowFile").disabled = state.busy || !state.outPath;
    $("btnOpenFile").disabled = state.busy || !state.outPath;
    $("btnOpenFolder").disabled = state.busy || !state.outDir;
}

function setProgress(percent, text) {
    $("progressBar").value = percent;
    $("progressText").textContent = text || "…";
}

$("btnPdf").addEventListener("click", async () => {
    const p = await window.api.selectPdf();
    if (p) {
        state.pdfPath = p;
        logLine(`Выбран PDF: ${p}`);

        // автоподстановка имени (без расширения)
        const base = p.split(/[\\/]/).pop().replace(/\.pdf$/i, "");
        if (!state.confirmedName) {        // не перетираем уже подтвержденное имя
            $("outName").value = base;
            // state.confirmedName = base;     // можно автоподтвердить
            logLine(`Имя файла подставлено: ${base}.xlsx`);
        }
    }
    updateUI();
});

$("btnDir").addEventListener("click", async () => {
    const d = await window.api.selectOutDir();
    if (d) {
        state.outDir = d;
        logLine(`Выбрана папка: ${d}`);
    }
    updateUI();
});

$("btnConfirm").addEventListener("click", () => {
    const name = $("outName").value.trim();
    state.confirmedName = name ? name : null;
    logLine(state.confirmedName ? `Имя файла подтверждено: ${state.confirmedName}.xlsx` : "Имя файла не задано");
    updateUI();
});

window.api.onProgress((p) => {
    const { stage, current, total, message } = p || {};
    let percent = 0;
    if (total && current != null) percent = Math.max(0, Math.min(100, Math.round((current / total) * 100)));

    setProgress(percent, message || `${stage ?? "work"}: ${current ?? ""}/${total ?? ""}`);

    if (!message) return;

    // логируем всегда только ключевые стадии
    if (stage === "init" || stage === "excel" || stage === "done") {
        logLine(message);
        return;
    }

    // страницы — только если включен подробный режим
    if (stage === "pages" && state.verbose) {
        logLine(message);
    }
});

$("btnStart").addEventListener("click", async () => {
    if (!canStart()) return;

    setBusy(true);
    state.outPath = null;
    setProgress(0, "Запуск…");
    logLine("Старт конвертации…");

    try {
        const outPath = await window.api.startConvert({
            pdfPath: state.pdfPath,
            outDir: state.outDir,
            outName: state.confirmedName
        });

        state.outPath = outPath;
        setProgress(100, `Готово: ${outPath}`);
        logLine(`Готово: ${outPath}`);
    } catch (e) {
        const msg = e?.message || String(e);
        setProgress(0, `Ошибка: ${msg}`);
        logLine(`Ошибка: ${msg}`);
    } finally {
        setBusy(false);
        updateUI();
    }
});

$("btnOpenFolder").addEventListener("click", async () => {
    if (!state.outDir) return;
    await window.api.openFolder(state.outDir);
});

$("btnShowFile").addEventListener("click", async () => {
    if (!state.outPath) return;
    await window.api.showItemInFolder(state.outPath);
});

$("btnOpenFile").addEventListener("click", async () => {
    if (!state.outPath) return;
    await window.api.openFile(state.outPath);
});

$("chkVerbose").addEventListener("change", () => {
    state.verbose = $("chkVerbose").checked;
    logLine(state.verbose ? "Подробный журнал: ВКЛ" : "Подробный журнал: ВЫКЛ");
});

updateUI();
setProgress(0, "Ожидание…");
logLine("Готово к работе.");