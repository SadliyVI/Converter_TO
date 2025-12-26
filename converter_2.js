import fs from "fs";
import ExcelJS from "exceljs";
import * as pdfjsLib from "pdfjs-dist/legacy/build/pdf.mjs";

/* ===================== helpers ===================== */

function clampCellText(s, max = 32000) {
    if (s == null) return null;
    s = String(s);
    return s.length > max ? s.slice(0, max) : s;
}

const normalizeSpaces = (s) => (s ?? "").toString().replace(/\s+/g, " ").trim();

function cleanupDashes(s) {
    return normalizeSpaces(
        (s ?? "")
            .toString()
            .replace(/[-‐‑‒–—]{3,}/g, " ")
            .replace(/[|_]{3,}/g, " ")
    );
}
const categoryKey = (s) => cleanupDashes(s).replace(/\s+/g, "").toLowerCase();

function normalizeBrokenWords(s) {
    return (s ?? "")
        .toString()
        .replace(/Итог\s+о/gi, "Итого")
        .replace(/кварт\s+алу/gi, "кварталу")
        .replace(/составляющим\s+пород\s+ам/gi, "составляющим породам")
        .replace(/\s+/g, " ")
        .trim();
}

function toNum(s) {
    if (s == null) return null;
    const t = String(s).trim().replace(",", ".");
    if (!t) return null;
    const n = Number(t);
    return Number.isFinite(n) ? n : null;
}
function toIntStrict(s) {
    if (s == null) return null;
    const t = String(s).trim();
    if (!/^\d+$/.test(t)) return null;
    const n = Number(t);
    return Number.isFinite(n) ? n : null;
}
function toFloatLoose(s) {
    if (s == null) return null;
    const m = String(s).replace(",", ".").match(/(\d+(?:\.\d+)?)/);
    if (!m) return null;
    const n = Number(m[1]);
    return Number.isFinite(n) ? n : null;
}
function isIntInRange(v, min, max) {
    return Number.isInteger(v) && v >= min && v <= max;
}
function isPolnotaStr(s) {
    const t = String(s ?? "").trim().replace(",", ".");
    return /^0\.\d$|^1\.0$|^1$/.test(t) && Number(t) <= 1.0;
}
function normalizeBonitet(s) {
    const t = String(s ?? "").trim();
    if (!t) return null;
    if (/^5[аА]$/.test(t)) return "5А";
    if (/^[1-5]$/.test(t)) return t;
    return null;
}

/* ---------------- totals by txt (не зависит от grid) ---------------- */

function isTotalsTitleByTxt(txt) {
    const t = normalizeBrokenWords(String(txt ?? "")).trim();
    if (/^Итого по категории/i.test(t)) return "TOTAL_CATEGORY";
    if (/^Итого по кварталу/i.test(t)) return "TOTAL_QUARTER";
    if (/^По составляющим породам/i.test(t)) return "TOTAL_SPECIES_HEADER";
    return null;
}

/* ---------------- vydel block by rule ---------------- */

function extractLeadingVydelFromText(txt) {
    const t = String(txt ?? "").trim();
    const m = t.match(/^(\d{1,4})(?=\s|$)/);
    if (!m) return null;
    const n = Number(m[1]);
    return Number.isInteger(n) ? n : null;
}

/**
 * "6 22.8 Культуры лесные" -> {vydel, area, tail}
 */
function extractVydelAreaTail(txt) {
    const t = normalizeBrokenWords(String(txt ?? "")).trim();
    const m = t.match(/^(\d{1,4})\s+(\d+(?:[.,]\d+)?)\s+(.+)$/);
    if (!m) return null;
    const vydel = Number(m[1]);
    const area = Number(m[2].replace(",", "."));
    const tail = m[3].trim();
    if (!Number.isInteger(vydel) || !Number.isFinite(area) || !tail) return null;
    return { vydel, area, tail };
}

/* ---------------- composition helpers ---------------- */

/**
 * Состав: допускаем пробелы между сегментами, потом склеиваем.
 * Примеры:
 *  - "5Е5Б" -> "5Е5Б"
 *  - "5Е 5Б" -> "5Е5Б"
 *  - "7Е2Б1ОС+Е" -> "7Е2Б1ОС+Е"
 */
function extractCompositionToken(s) {
    const t = normalizeSpaces(s).toUpperCase();
    const m = t.match(
        /\b\d{1,2}[А-ЯЁA-Z]{1,3}(?:\s*\d{1,2}[А-ЯЁA-Z]{1,3})*(?:\s*\+\s*[А-ЯЁA-Z]{1,3})*\b/
    );
    return m ? m[0].replace(/\s+/g, "") : null;
}

/**
 * FIX: корректно определяет "чистую строку состава", даже если состав в тексте с пробелами ("5Е 5Б").
 */
function isPureCompositionLine(txt) {
    const t = normalizeSpaces(txt).toUpperCase();

    // находим "сырой" матч (с пробелами), чтобы корректно вырезать
    const re = /\b\d{1,2}[А-ЯЁA-Z]{1,3}(?:\s*\d{1,2}[А-ЯЁA-Z]{1,3})*(?:\s*\+\s*[А-ЯЁA-Z]{1,3})*\b/;
    const m = re.exec(t);
    if (!m) return null;

    const rawComp = m[0]; // может быть "5Е 5Б"
    const comp = rawComp.replace(/\s+/g, ""); // "5Е5Б"

    const rest = normalizeSpaces(
        t.replace(rawComp, " ").replace(/[.,;:()]/g, " ")
    );

    return rest ? null : comp;
}

/**
 * "лесные 5Е 5Б" -> { word:"лесные", composition:"5Е5Б" }
 */
function splitLeadingWordAndComposition(desc) {
    const t = normalizeSpaces(desc);
    const m = t.match(
        /^([А-ЯЁA-Zа-яё]{3,})\s+(\d{1,2}[А-ЯЁA-ZЁ]{1,3}(?:\s*\d{1,2}[А-ЯЁA-ZЁ]{1,3})*(?:\s*\+\s*[А-ЯЁA-ZЁ]{1,3})*)\b/
    );
    if (!m) return null;
    return { word: m[1], composition: m[2].toUpperCase().replace(/\s+/g, "") };
}

function includesWordCI(text, word) {
    const t = String(text ?? "").toLowerCase();
    const w = String(word ?? "").toLowerCase();
    if (!t || !w) return false;
    return t.split(/\s+/g).includes(w);
}

/* ---------------- duplicates clean ---------------- */

function normalizeForDup(s) {
    return normalizeSpaces(
        String(s ?? "")
            .toLowerCase()
            .replace(/[.,;:()]/g, " ")
            .replace(/\bс\b/g, " ")
    );
}
function clearDuplicate24(cells) {
    const d = normalizeForDup(cells[3]);
    const h = normalizeForDup(cells[24]);
    if (!d || !h) return cells;
    if (d.includes(h) || h.includes(d)) cells[24] = "";
    return cells;
}

/* ---------------- OBJECT: split description / note ---------------- */

function makeLoosePhraseRegex(phrase) {
    // "Расчистка просек" будет матчить "Расчис тка про сек" и подобное
    const p = String(phrase)
        .trim()
        .split(/\s+/g)
        .map((word) => word.split("").map((ch) => ch.replace(/[.*+?^${}()|[\]\\]/g, "\\$&") + "\\s*").join(""))
        .join("\\s+");
    return new RegExp(p, "gi");
}

function splitObjectDescriptionAndNote(desc, extra24) {
    const d0 = normalizeSpaces(desc);
    const h0 = normalizeSpaces(extra24);

    const PHRASES = [{ canon: "Расчистка просек", re: makeLoosePhraseRegex("Расчистка просек") }];

    let noteParts = [];
    let d = d0;
    let h = h0;

    for (const { canon, re } of PHRASES) {
        if (re.test(d)) {
            noteParts.push(canon);
            d = d.replace(re, " ");
        }
        re.lastIndex = 0;

        if (re.test(h)) {
            noteParts.push(canon);
            h = h.replace(re, " ");
        }
        re.lastIndex = 0;
    }

    d = normalizeSpaces(d);
    h = normalizeSpaces(h);

    noteParts = [...new Set(noteParts)];

    const note = noteParts.length ? noteParts.join(", ") : null;
    return { description: d || null, note, cleaned24: h || null };
}

/* ---------------- pdf extraction ---------------- */
async function getPageItems(pdfDoc, pageNum) {
    const page = await pdfDoc.getPage(pageNum);
    const content = await page.getTextContent();
    return content.items
        .map((it) => ({
            str: (it.str ?? "").toString(),
            x: it.transform[4],
            y: it.transform[5],
            w: it.width ?? 0,
        }))
        .filter((it) => it.str.trim().length > 0);
}

function groupItemsToLines(items) {
    const sorted = [...items].sort((a, b) => b.y - a.y || a.x - b.x);
    const Y_TOL = 2.0;
    const lines = [];
    for (const it of sorted) {
        let line = lines.find((l) => Math.abs(l.y - it.y) <= Y_TOL);
        if (!line) {
            line = { y: it.y, items: [] };
            lines.push(line);
        }
        line.items.push(it);
    }
    for (const l of lines) l.items.sort((a, b) => a.x - b.x);
    return lines.sort((a, b) => b.y - a.y);
}

function lineToText(line) {
    let out = "";
    let prev = null;
    for (const it of line.items) {
        if (prev) {
            const gap = it.x - prev.x;
            if (gap > 8) out += " ";
        }
        out += it.str;
        prev = it;
    }
    return out.replace(/\s+/g, " ").trim();
}

/* ---------------- header parsing ---------------- */
function extractQuarter(linesText) {
    for (const s of linesText) {
        const m = s.match(/Квартал\s+(\d+)/i);
        if (m) return Number(m[1]);
    }
    return null;
}

function extractProtectionCategory(linesText) {
    for (let i = 0; i < linesText.length; i++) {
        const s1 = cleanupDashes(linesText[i]);
        const s2 = cleanupDashes(linesText[i + 1] || "");
        const combined = cleanupDashes(`${s1} ${s2}`);

        if (!/Категория\s+защитности/i.test(combined)) continue;

        const parts = combined.split(/Категория\s+защитности\s*:?\s*/i);
        if (parts.length < 2) continue;

        const after = parts[1];

        const mQ = after.match(/^(.*?)(?:Кв\s*а?\s*р?\s*т?\s*а?\s*л\.?\b)/i);
        const beforeQuarter = mQ ? mQ[1] : after;

        const beforeQuarter2 = beforeQuarter.split(/Кв/i)[0];
        const cat = cleanupDashes(beforeQuarter2);

        if (!cat) continue;
        if (/квартал/i.test(cat)) continue;
        if (cat.length < 3) continue;

        return cat;
    }
    return null;
}

/* ---------------- grid detection ---------------- */
function detectColumnGrid(lines) {
    let best = null;

    // A) ruler by numbers 1..24
    for (const l of lines) {
        const txt = lineToText(l);
        if (!/\b24\b/.test(txt)) continue;

        const anchors = new Map();
        for (const it of l.items) {
            const s = it.str;
            const re = /\b(\d{1,2})\b/g;
            let m;
            while ((m = re.exec(s)) !== null) {
                const n = Number(m[1]);
                if (n < 1 || n > 24) continue;

                const frac = s.length ? m.index / s.length : 0;
                const xApprox = it.x + (it.w || 0) * frac;
                if (!anchors.has(n)) anchors.set(n, xApprox);
            }
        }

        const score = anchors.size;
        if (score >= 18) {
            if (!best || score > best.score) best = { anchors, score, source: "ruler" };
        }
    }
    if (best) return { anchors: best.anchors, source: best.source };

    // B) fallback: first MAIN-like line
    for (const l of lines) {
        const txt = lineToText(l);
        if (!txt.match(/^\d+\s+\d+(\.\d+)?\s+/)) continue;
        if (!txt.match(/\b0\.\d\b/)) continue;

        const xs = l.items
            .filter((it) => it.str.trim() && it.str.trim() !== ":")
            .map((it) => ({ x: it.x, s: it.str.trim() }))
            .sort((a, b) => a.x - b.x);

        const uniq = [];
        const X_TOL = 1.5;
        for (const it of xs) {
            if (!uniq.length || Math.abs(uniq[uniq.length - 1].x - it.x) > X_TOL) uniq.push(it);
        }
        if (uniq.length < 10) continue;

        const colsOrder = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24];
        const anchors = new Map();
        for (let i = 0; i < Math.min(colsOrder.length, uniq.length); i++) {
            anchors.set(colsOrder[i], uniq[i].x);
        }
        return { anchors, source: "main-fallback" };
    }

    return null;
}

/**
 * Мы НЕ кладём overflow сразу в col24. Складываем в cells[0], потом разрулим эвристикой.
 */
function splitLineIntoCells(line, grid) {
    const cells = Array.from({ length: 25 }, () => "");
    cells[0] = ""; // overflow bucket

    if (!grid) {
        cells[24] = lineToText(line);
        return cells;
    }

    const anchorsArr = Array.from(grid.anchors.entries())
        .map(([col, x]) => ({ col, x }))
        .sort((a, b) => a.col - b.col);

    const MAX_ANCHOR_DIST = 18; // подберите под PDF при необходимости

    for (const it of line.items) {
        const piece = it.str.trim();
        if (!piece || piece === ":") continue;

        const cx = it.x + (it.w || 0) / 2;

        let best = null;
        for (const a of anchorsArr) {
            const d = Math.abs(cx - a.x);
            if (!best || d < best.d) best = { col: a.col, d };
        }
        if (!best) continue;

        if (best.d > MAX_ANCHOR_DIST) {
            cells[0] = (cells[0] ? cells[0] + " " : "") + piece;
            continue;
        }

        const c = best.col;
        cells[c] = (cells[c] ? cells[c] + " " : "") + piece;
    }

    for (let i = 1; i <= 24; i++) cells[i] = cells[i].replace(/\s+/g, " ").trim();
    cells[0] = cells[0].replace(/\s+/g, " ").trim();
    return cells;
}

function overflowLooksLikeComposition(s) {
    const t = normalizeSpaces(s).toUpperCase();
    if (!t) return false;
    return /\b\d{1,2}[А-ЯЁA-Z]{1,3}(?:\s*\d{1,2}[А-ЯЁA-Z]{1,3})*(?:\s*\+\s*[А-ЯЁA-Z]{1,3})*\b/.test(t);
}
function overflowLooksLikeNoteText(s) {
    const t = normalizeSpaces(s).toLowerCase();
    if (!t) return false;
    return /(подрост|подлесок|болот|дорог|земли|линейного\s+протяжения|просек|реки|ручь|км|тыс\.шт\/га)/i.test(t);
}

function applyOverflowHeuristics(cells) {
    const ov = (cells[0] || "").trim();
    if (!ov) return cells;

    if (overflowLooksLikeComposition(ov) && !overflowLooksLikeNoteText(ov)) {
        cells[3] = normalizeSpaces([cells[3], ov].filter(Boolean).join(" "));
        cells[0] = "";
        return cells;
    }

    cells[24] = normalizeSpaces([cells[24], ov].filter(Boolean).join(" "));
    cells[0] = "";
    return cells;
}

/* ---------------- per-line normalizers ---------------- */
function normalizeStartColumns(cells) {
    if (!/^\d+$/.test(cells[1] || "")) return cells;

    const c2 = (cells[2] || "").trim();
    if (!c2) return cells;

    const m = c2.match(/^(\d+(?:[.,]\d+)?)(?:\s+(.+))?$/);
    if (!m) return cells;

    const areaPart = m[1];
    const tail = (m[2] || "").trim();
    if (!tail) return cells;

    cells[2] = areaPart.replace(",", ".");
    cells[3] = cells[3] ? (tail + " " + cells[3]).replace(/\s+/g, " ").trim() : tail;

    return cells;
}

function normalizeObjectDescriptionSplit(cells) {
    const c3 = (cells[3] || "").trim();
    const c5 = (cells[5] || "").trim();

    if (c3 && c5 && !/^\d+(\.\d+)?$/.test(c5)) {
        cells[3] = normalizeSpaces(`${c3} ${c5}`);
        cells[5] = "";
    }
    return cells;
}

function normalizeSingleTreesCells(cells) {
    const c8 = (cells[8] || "").trim();
    if (!c8) return cells;
    const m = c8.match(/^(\d+)\s+(\d+)$/);
    if (m && !cells[7]) {
        cells[7] = m[1];
        cells[8] = m[2];
    }
    return cells;
}

/* ---------------- noise filter ---------------- */
function isNoiseText(txt) {
    if (!txt) return true;
    const t = String(txt).trim();

    // totals НЕ шум
    if (/^Итого\b/i.test(t) || /^По\s+составляющим\s+породам\b/i.test(t)) return false;

    if (/^\d{1,4}$/.test(t)) return true;
    if (t.includes("----------------------------------------------------------------")) return true;
    if (/^[-]{5,}$/.test(t)) return true;
    if (t.trim().startsWith(":")) return true;

    // много двоеточий = заголовок таблицы/линейка
    const colonCount = (t.match(/:/g) || []).length;
    if (colonCount >= 10) return true;

    const compact = t.replace(/\s+/g, "").toLowerCase();

    // шапка страницы (в т.ч. разорванные слова)
    if (compact.includes("категориязащитности") && compact.includes("квартал")) return true;
    if (compact.includes("таксационноеописание")) return true;
    if (compact.includes("уч.лес-во") || compact.includes("уч.лес-во".replace(/\s+/g, ""))) return true;

    return false;
}

/* ---------------- helpers ---------------- */
function isEmptyRange(cells, a, b) {
    for (let i = a; i <= b; i++) if (cells[i]) return false;
    return true;
}
function hasAnyRange(cells, a, b) {
    for (let i = a; i <= b; i++) if (cells[i]) return true;
    return false;
}
function looksLikePolnota(cells) {
    return /^\d\.\d$/.test((cells[14] || "").replace(",", "."));
}
function looksLikeStandDescription(desc) {
    const s = (desc ?? "").toString().toLowerCase();
    return s.includes("насажд") || s.includes("культ") || s.includes("полог");
}
function looksLikeMainNoId(cells) {
    if ((cells[1] || "").trim()) return false;
    if (!looksLikePolnota(cells)) return false;

    const y = toIntStrict(cells[4]);
    if (cells[4] && !(y != null && y >= 1 && y <= 5)) return false;

    const el = (cells[6] || "").trim();
    const hasElement = !!el && /^[А-ЯЁA-Z]{1,6}$/.test(el.replace(/\s+/g, ""));

    const age = toIntStrict(cells[7]);
    const h = toIntStrict(cells[8]);
    const d = toIntStrict(cells[9]);

    const hasDims = (age != null && age <= 300) || (h != null && h <= 99) || (d != null && d <= 150);

    return hasElement || hasDims;
}
function looksLikeRealObjectName(desc) {
    const s = (desc ?? "").toString().trim().toLowerCase();
    if (!s) return false;
    return (
        s.startsWith("болота") ||
        s.startsWith("дороги") ||
        s.startsWith("ручьи") ||
        s.startsWith("реки") ||
        s.startsWith("просеки") ||
        s.startsWith("вырубки") ||
        s.startsWith("поляны")
    );
}
function looksLikeDescriptionContinuationByC3(cells) {
    const c3 = (cells[3] || "").trim();
    return !!c3 && /^[+,]/.test(c3);
}

/* ---------------- totals parsing ---------------- */
function extractNumbersFromText(txt) {
    const s = String(txt ?? "").replace(/,/g, ".");
    return (s.match(/\d+(?:\.\d+)?\.?/g) || [])
        .map((x) => x.replace(/\.$/, ""))
        .map((x) => Number(x))
        .filter((n) => Number.isFinite(n));
}

function parseTotalsNumbersFromTwoLines(titleTxt, nextTxt) {
    const nums = extractNumbersFromText(`${titleTxt} ${nextTxt}`);
    const area = nums.length ? nums[0] : null;

    const rest = nums.slice(1);
    const ZapasObshiy = rest.length ? Math.max(...rest) : nums[1] ?? null;

    let edinichnye = null;
    if (rest.length >= 2) {
        const minRest = Math.min(...rest);
        if (minRest !== ZapasObshiy) edinichnye = minRest;
    }

    return { area, ZapasObshiy, edinichnye };
}

function parseTotalsSpeciesRow(txt) {
    const s = normalizeBrokenWords(String(txt ?? "").trim()).replace(/\s+/g, " ");
    const m = s.match(/^([А-ЯЁA-Z]{1,6})(?:\s+(\d+(?:[.,]\d+)?))?$/i);
    if (!m) return null;

    const element = m[1].toUpperCase();
    const zapas = m[2] != null ? Number(m[2].replace(",", ".")) : null;
    if (m[2] != null && !Number.isFinite(zapas)) return null;

    return { element, zapasPoSostavu: zapas };
}

/* ---------------- normalize zapas / klass ---------------- */
function normalizeZapasAndKlassTov(cells) {
    const trySplit = (idx) => {
        const s = (cells[idx] || "").trim();
        if (!s) return false;

        const m = s.match(/^(\d+(?:[.,]\d+)?)\s+(\d+)$/);
        if (!m) return false;

        const first = m[1].replace(",", ".");
        const second = m[2];

        if (idx === 17) {
            cells[17] = first;
            if (!cells[18]) cells[18] = second;
            return true;
        }

        if (idx === 18) {
            if (!cells[17]) cells[17] = first;
            cells[18] = second;
            return true;
        }
        return false;
    };

    if (!trySplit(17)) trySplit(18);
    return cells;
}

/* ---------------- normalize MAIN ---------------- */
function normalizeMainCells(cells) {
    {
        const s = (cells[9] || "").trim();
        const m = s.match(/^(\d+(?:[.,]\d+)?)\s+(\d+)$/);
        if (m) {
            cells[9] = m[1].replace(",", ".");
            if (!cells[10]) cells[10] = m[2];
        }
    }

    {
        const s17 = (cells[17] || "").trim();
        const nums = s17.match(/\d+(?:[.,]\d+)?\.?/g);
        if (nums && nums.length >= 2) {
            const n1 = nums[0].replace(",", ".");
            const n2 = nums[1].replace(",", ".");
            const n3 = nums[2] ? nums[2].replace(",", ".") : null;

            if (!cells[16]) cells[16] = n1;
            cells[17] = n2;
            if (n3 && !cells[18]) cells[18] = n3;
        }
    }

    {
        const s16 = (cells[16] || "").trim();
        const m = s16.match(/^(\d+(?:[.,]\d+)?)\.?\s+(\d+(?:[.,]\d+)?)\.?$/);
        if (m) {
            cells[16] = m[1].replace(",", ".");
            if (!cells[17]) cells[17] = m[2].replace(",", ".");
        }
    }

    return cells;
}

function normalizeMainByRanges(cells) {
    const getInt = (idx) => toIntStrict(cells[idx]);
    const getFloat = (idx) => toFloatLoose(cells[idx]);

    if (!cells[4]) {
        const v = getInt(4);
        if (isIntInRange(v, 1, 5)) cells[4] = String(v);
    } else {
        const v = getInt(4);
        if (!isIntInRange(v, 1, 5)) cells[4] = "";
    }

    if (cells[5]) {
        const v = getInt(5);
        if (!isIntInRange(v, 1, 99)) cells[5] = "";
    }

    if (cells[7]) {
        const v = getInt(7);
        if (!isIntInRange(v, 1, 300)) cells[7] = "";
    }
    if (cells[8]) {
        const v = getInt(8);
        if (!isIntInRange(v, 1, 99)) cells[8] = "";
    }

    {
        const d9 = getInt(9);
        if (!isIntInRange(d9, 1, 150)) {
            const candidates = [9, 10, 11].map((i) => ({ i, v: getInt(i) })).filter((x) => isIntInRange(x.v, 1, 150));
            if (candidates.length) cells[9] = String(candidates[0].v);
        }
    }

    if (cells[10]) {
        const v = getInt(10);
        if (!isIntInRange(v, 1, 12)) cells[10] = "";
    }
    if (cells[11]) {
        const v = getInt(11);
        if (!isIntInRange(v, 1, 10)) cells[11] = "";
    }

    if (cells[12]) {
        const b = normalizeBonitet(cells[12]);
        cells[12] = b ? b : "";
    }

    if (cells[14] && !isPolnotaStr(cells[14])) cells[14] = "";

    if (!cells[16] && cells[17]) {
        const nums = String(cells[17]).replace(",", ".").match(/\d+(?:\.\d+)?\.?/g);
        if (nums && nums.length >= 2) {
            cells[16] = nums[0].replace(",", ".");
            cells[17] = nums[1].replace(",", ".");
            if (nums[2] && !cells[18]) cells[18] = nums[2];
        }
    }

    {
        const z16 = getFloat(16);
        const z17 = getFloat(17);
        const z18 = getFloat(18);
        const looksZapas = (x) => x != null && x > 9;

        if (z16 != null && z17 != null && Math.abs(z16 - z17) < 1e-9 && looksZapas(z18)) {
            cells[17] = String(z18).replace(",", ".");
            const k18 = getInt(18);
            if (!(k18 != null && k18 >= 1 && k18 <= 9)) cells[18] = "";
        } else {
            const k18 = getInt(18);
            if (cells[18] && !(k18 != null && k18 >= 1 && k18 <= 9)) {
                if (looksZapas(z18)) cells[18] = "";
            }
        }
    }

    return cells;
}

function normalizeMainZapasFromText(cells, txt) {
    const s = String(txt ?? "").replace(/,/g, ".");
    const allNums = (s.match(/\d+(?:\.\d+)?\.?/g) || [])
        .map((x) => x.replace(/\.$/, ""))
        .map((x) => Number(x))
        .filter((n) => Number.isFinite(n));

    if (!allNums.length) return cells;

    const isKlass = (n) => Number.isInteger(n) && n >= 1 && n <= 9;
    const isZapas = (n) => n > 9;

    const zap = allNums.filter(isZapas);

    if (zap.length >= 2) {
        const zObsh = zap[zap.length - 2];
        const zSost = zap[zap.length - 1];
        cells[16] = String(zObsh);
        cells[17] = String(zSost);
    } else if (zap.length === 1) {
        cells[16] = String(zap[0]);
    }

    if (!cells[18]) {
        const last3 = allNums.slice(-3);
        if (last3.length === 3) {
            const [a, b, c] = last3;
            if (isZapas(a) && isZapas(b) && isKlass(c)) cells[18] = String(c);
        }
    }

    return cells;
}

/* ---------------- SPECIES normalization ---------------- */
function normalizeSpeciesCells(cells) {
    if (cells[6]) cells[6] = String(cells[6]).replace(/\s+/g, "").toUpperCase();
    if (cells[7]) cells[7] = String(cells[7]).trim();

    {
        const el = (cells[6] || "").trim();
        const m = el.match(/^([А-ЯЁA-Z]{1,6})(\d{1,3})$/i);
        if (m) {
            cells[6] = m[1].toUpperCase();
            if (!cells[7]) cells[7] = m[2];
            return cells;
        }
    }

    {
        const el = (cells[6] || "").trim();
        const m = el.match(/^([А-ЯЁA-Z]{1,6})\s+(\d{1,3})$/i);
        if (m) {
            cells[6] = m[1].toUpperCase();
            if (!cells[7]) cells[7] = m[2];
            return cells;
        }
    }

    {
        const el = (cells[6] || "").trim().toUpperCase();
        const a7 = (cells[7] || "").trim();
        const m = a7.match(/^А\s*(\d{1,3})$/i);
        if (el === "ОЛС" && m) {
            cells[6] = "ОЛСА";
            cells[7] = m[1];
            return cells;
        }
    }

    {
        const el = (cells[6] || "").trim();
        if (/\d/.test(el)) {
            const m = el.match(/^([А-ЯЁA-Z]{1,6})/i);
            if (m) cells[6] = m[1].toUpperCase();
        }
    }

    return cells;
}

/* ---------------- record builder ---------------- */
const EMPTY_FIELDS = {
    area: null,
    description: null,
    yarus: null,
    yarusHigth: null,
    element: null,
    age: null,
    higth: null,
    diam: null,
    ageKlass: null,
    ageGroup: null,
    bonitet: null,
    forestType: null,
    polnota: null,
    zapasNa1ga: null,
    ZapasObshiy: null,
    zapasPoSostavu: null,
    klassTovarnosti: null,
    suhostoy: null,
    redin: null,
    edinichnye: null,
    zakhl: null,
    zakhlLikvid: null,
    hozMeropriyatiya: null,
    note: null,
    raw: null,
};

function buildRecord({ ctx, rowNo, kind, page, vydel, category, quarter }, overrides = {}) {
    return {
        quarter: quarter ?? ctx.quarter,
        rowNo,
        category: category ?? ctx.category,
        vydel: vydel ?? null,
        kind,
        page,
        ...EMPTY_FIELDS,
        ...overrides,
    };
}
function pushRecord(records, rec) {
    rec.note = clampCellText(rec.note);
    rec.raw = clampCellText(rec.raw);
    records.push(rec);
}
function cellsRaw(cells) {
    return clampCellText(
        cells
            .map((x, i) => (i === 0 ? (x ? `0:${x}` : "") : x ? `${i}:${x}` : ""))
            .filter(Boolean)
            .join(" | ")
    );
}

function mapCellsToNamed(cells) {
    return {
        vydel: toNum(cells[1]),
        area: toNum(cells[2]),
        description: cells[3] || null,

        yarus: toNum(cells[4]),
        yarusHigth: toNum(cells[5]),
        element: cells[6] || null,
        age: toNum(cells[7]),
        higth: toNum(cells[8]),
        diam: toNum(cells[9]),
        ageKlass: toNum(cells[10]),
        ageGroup: toNum(cells[11]),
        bonitet: toNum(cells[12]),
        forestType: cells[13] || null,
        polnota: toNum(cells[14]),
        zapasNa1ga: toNum(cells[15]),
        ZapasObshiy: toNum(cells[16]),
        zapasPoSostavu: toNum(cells[17]),
        klassTovarnosti: cells[18] || null,
        suhostoy: cells[19] || null,
        redin: cells[20] || null,
        edinichnye: cells[21] || null,
        zakhl: cells[22] || null,
        zakhlLikvid: cells[23] || null,
        hozMeropriyatiya: cells[24] || null,
    };
}

/* ---------------- classification ---------------- */
function classifyRow(cells, ctx) {
    const c1 = cells[1],
        c2 = cells[2],
        c3 = cells[3];

    const leftText = normalizeBrokenWords(cells.slice(1, 7).filter(Boolean).join(" "));

    // (тоталы теперь обрабатываем по txt раньше)
    if (/^Итого по категории/i.test(leftText)) return "TOTAL_CATEGORY";
    if (/^Итого по кварталу/i.test(leftText)) return "TOTAL_QUARTER";
    if (/^По составляющим породам/i.test(leftText)) return "TOTAL_SPECIES_HEADER";

    if (leftText.match(/^(подрост|подлесок|Болота:|Земли линейного протяжения:|Расчистка просек)/i)) {
        ctx.inSingleTrees = false;
        return "NOTE";
    }

    if (/^Единичные деревья/i.test(leftText)) {
        ctx.inSingleTrees = true;
        return "NOTE";
    }

    if (ctx.inSingleTrees) {
        const whole = cells.slice(1).filter(Boolean).join(" ").toLowerCase();
        if (whole.includes("подрост") || whole.includes("подлесок")) {
            ctx.inSingleTrees = false;
            return "NOTE";
        }
        const coeff = (cells[2] || "").trim();
        const hasNumbersRight = /\d/.test(cells.slice(5).filter(Boolean).join(" "));
        if (/^\d{1,2}[А-ЯЁA-Z]{1,3}$/i.test(coeff) && hasNumbersRight) {
            ctx.inSingleTrees = false;
            return "SINGLE_TREES";
        }
        ctx.inSingleTrees = false;
    }

    if (/^\d+$/.test(c1) && /^\d+(\.\d+)?$/.test(c2)) {
        if (looksLikePolnota(cells)) return "MAIN";
        if (c3) return "OBJECT";
    }

    if (isEmptyRange(cells, 1, 5) && hasAnyRange(cells, 6, 24)) return "SPECIES";
    return "TEXT";
}

/* ---------------- parsePdf with progress ---------------- */
async function parsePdf(pdfPath, onProgress) {
    const data = new Uint8Array(fs.readFileSync(pdfPath));
    const pdfDoc = await pdfjsLib.getDocument({ data }).promise;

    const records = [];
    let lastQuarter = null;
    let lastCategory = null;

    const ctx = {
        quarter: null,
        category: null,
        vydel: null,
        inSingleTrees: false,

        totalsMode: null,
        totalsExpectSpecies: false,
        pendingTotal: null,

        lastMainIndex: null,
        pendingMain: null,
    };

    const totalPages = pdfDoc.numPages;

    for (let p = 1; p <= totalPages; p++) {
        onProgress?.({ stage: "pages", current: p, total: totalPages, message: `Обработка страниц: ${p}/${totalPages}` });

        const items = await getPageItems(pdfDoc, p);
        const lines = groupItemsToLines(items);
        const linesText = lines.map(lineToText);

        const q = extractQuarter(linesText) ?? lastQuarter;
        if (q != null) lastQuarter = q;
        if (q != null) ctx.quarter = q;

        const extractedCat = extractProtectionCategory(linesText);
        let cat = extractedCat ? cleanupDashes(extractedCat) : null;

        if (cat && /^Квартал\b/i.test(cat)) cat = null;
        if (cat && /кв\s*а?\s*р?\s*т?\s*а?\s*л/i.test(cat)) cat = null;
        if (!cat && lastCategory) cat = cleanupDashes(lastCategory);

        if (cat) {
            if (lastCategory && categoryKey(cat) === categoryKey(lastCategory)) cat = cleanupDashes(lastCategory);
            else lastCategory = cat;
            ctx.category = cat;
        } else {
            ctx.category = null;
        }

        const grid = detectColumnGrid(lines);
        let rowNo = 0;

        for (let li = 0; li < lines.length; li++) {
            const line = lines[li];
            const txtRaw = lineToText(line);
            if (isNoiseText(txtRaw)) continue;

            rowNo++;
            const txt = normalizeBrokenWords(txtRaw);

            // 1) totals заголовки — строго по txt
            {
                const totalsKind = isTotalsTitleByTxt(txt);
                if (totalsKind === "TOTAL_CATEGORY" || totalsKind === "TOTAL_QUARTER") {
                    ctx.totalsMode = totalsKind === "TOTAL_CATEGORY" ? "CATEGORY" : "QUARTER";
                    ctx.totalsExpectSpecies = false;
                    ctx.pendingTotal = { kind: totalsKind, titleTxt: txt, page: p, rowNo };
                    ctx.inSingleTrees = false;
                    continue;
                }
                if (totalsKind === "TOTAL_SPECIES_HEADER") {
                    ctx.totalsExpectSpecies = true;
                    ctx.inSingleTrees = false;
                    pushRecord(
                        records,
                        buildRecord({ ctx, rowNo, kind: "TOTAL_SPECIES_HEADER", page: p }, {
                            note: "По составляющим породам",
                            raw: txt,
                        })
                    );
                    continue;
                }
            }

            // 2) numbers line for totals — независимо от kind/grid
            if (ctx.pendingTotal) {
                const t = txt.trim();
                if (/^\d+(?:[.,]\d+)?(?:\s+\d+(?:[.,]\d+)?)+$/.test(t)) {
                    const parsed = parseTotalsNumbersFromTwoLines(ctx.pendingTotal.titleTxt, t);
                    pushRecord(
                        records,
                        buildRecord({ ctx, rowNo, kind: ctx.pendingTotal.kind, page: p }, {
                            area: parsed.area,
                            ZapasObshiy: parsed.ZapasObshiy,
                            edinichnye: parsed.edinichnye,
                            note: ctx.pendingTotal.kind === "TOTAL_CATEGORY" ? "Итого по категории" : "Итого по кварталу",
                            raw: `title:${ctx.pendingTotal.titleTxt} | nums:${t}`,
                        })
                    );
                    ctx.pendingTotal = null;
                    continue;
                } else {
                    ctx.pendingTotal = null;
                }
            }

            // 3) totals species rows — по txt; прекращаем, когда начинается новый выдел
            if (ctx.totalsExpectSpecies) {
                const v0 = extractLeadingVydelFromText(txt);
                if (v0 != null) {
                    ctx.totalsExpectSpecies = false;
                } else {
                    const parsed = parseTotalsSpeciesRow(txt);
                    if (parsed) {
                        pushRecord(
                            records,
                            buildRecord({ ctx, rowNo, kind: "TOTAL_SPECIES_ROW", page: p }, {
                                element: parsed.element,
                                zapasPoSostavu: parsed.zapasPoSostavu,
                                raw: txt,
                            })
                        );
                        continue;
                    }
                }
            }

            // 4) контекст выдела по правилу: старт блока выдела начинается с целого числа
            {
                const v0 = extractLeadingVydelFromText(txt);
                if (v0 != null && ctx.vydel !== v0) {
                    ctx.vydel = v0;
                    ctx.lastMainIndex = null;
                    ctx.pendingMain = null;
                    ctx.inSingleTrees = false;
                }
            }

            // Если после VYDEL_HEADER у культур отдельной строкой идет состав (например "5Е5Б" или "5Е 5Б"),
            // то запоминаем его и не пишем как TEXT.
            if (ctx.pendingMain && looksLikeStandDescription(ctx.pendingMain.description)) {
                const compLine = isPureCompositionLine(txt);
                if (compLine) {
                    ctx.pendingMain.composition = compLine;
                    continue;
                }
            }

            // --- split to cells ---
            let cells = splitLineIntoCells(line, grid);
            cells = applyOverflowHeuristics(cells);
            cells = normalizeStartColumns(cells);
            cells = normalizeObjectDescriptionSplit(cells);
            cells = normalizeZapasAndKlassTov(cells);

            // восстановление "шапки выдела" из txt, чтобы VYDEL_HEADER был "Культуры лесные"
            {
                const hdr = extractVydelAreaTail(txt);
                if (hdr) {
                    const hasMainMetrics = hasAnyRange(cells, 4, 18);
                    if (!hasMainMetrics) {
                        if (!cells[1]) cells[1] = String(hdr.vydel);
                        if (!cells[2]) cells[2] = String(hdr.area);
                        cells[3] = hdr.tail;

                        const c24 = (cells[24] || "").trim();
                        if (c24 && hdr.tail.toLowerCase().includes(c24.toLowerCase())) cells[24] = "";
                    }
                }
            }

            let kind = classifyRow(cells, ctx);

            // pendingMain reset (защита)
            if (ctx.pendingMain && /^\d+$/.test(cells[1] || "") && toNum(cells[1]) !== ctx.pendingMain.vydel) ctx.pendingMain = null;

            // OBJECT -> pendingMain + VYDEL_HEADER
            if (kind === "OBJECT") {
                const desc = (cells[3] || "").trim();
                const hasMainMetrics = hasAnyRange(cells, 4, 18);

                if (desc && !looksLikeRealObjectName(desc) && !hasMainMetrics) {
                    const vyd = toNum(cells[1]) ?? ctx.vydel;
                    const ar = toNum(cells[2]);

                    const headerIndex = records.length;

                    cells = clearDuplicate24(cells);

                    pushRecord(
                        records,
                        buildRecord({ ctx, rowNo, kind: "VYDEL_HEADER", page: p, vydel: vyd }, {
                            area: ar,
                            description: desc,
                            raw: cellsRaw(cells),
                        })
                    );

                    ctx.pendingMain = { vydel: vyd, area: ar, description: desc, page: p, rowNo, headerIndex, composition: null };
                    ctx.vydel = vyd;
                    continue;
                }
            }

            // pendingMain -> MAIN
            if (ctx.pendingMain && (kind === "TEXT" || kind === "OBJECT")) {
                const hasMainMetrics = hasAnyRange(cells, 4, 18);
                if (hasMainMetrics && looksLikePolnota(cells)) {
                    cells[1] = String(ctx.pendingMain.vydel ?? "");
                    cells[2] = "";

                    const d0 = (ctx.pendingMain.description || "").trim();
                    const d1 = (cells[3] || "").trim();

                    if (looksLikeStandDescription(d0)) {
                        // FIX: приоритет — состав, пойманный отдельной строкой
                        if (ctx.pendingMain.composition) {
                            // если d1 содержит "лесные 5Е", можно все равно дописать слово в заголовок
                            const moved = splitLeadingWordAndComposition(d1);
                            if (moved) {
                                const hi = ctx.pendingMain.headerIndex;
                                if (hi != null && records[hi]) {
                                    const curHdr = records[hi].description;
                                    if (!includesWordCI(curHdr, moved.word)) {
                                        records[hi].description = normalizeSpaces([curHdr, moved.word].filter(Boolean).join(" "));
                                        records[hi].raw = clampCellText(normalizeSpaces([records[hi].raw, `| hdr+${moved.word}`].join(" ")));
                                    }
                                }
                            }

                            cells[3] = ctx.pendingMain.composition; // MAIN = только состав
                        } else {
                            const moved = splitLeadingWordAndComposition(d1);
                            if (moved) {
                                const hi = ctx.pendingMain.headerIndex;
                                if (hi != null && records[hi]) {
                                    const curHdr = records[hi].description;
                                    if (!includesWordCI(curHdr, moved.word)) {
                                        records[hi].description = normalizeSpaces([curHdr, moved.word].filter(Boolean).join(" "));
                                        records[hi].raw = clampCellText(normalizeSpaces([records[hi].raw, `| hdr+${moved.word}`].join(" ")));
                                    }
                                }

                                cells[3] = moved.composition;
                            } else {
                                const comp = extractCompositionToken(d1);
                                cells[3] = comp ? comp : d1;
                            }
                        }
                    } else {
                        const comp = extractCompositionToken(d1);
                        cells[3] = comp ? comp : d1;
                    }

                    ctx.pendingMain = null;
                    kind = "MAIN";
                }
            }

            // no-id MAIN (yarus 2..)
            if (kind === "TEXT" && !ctx.pendingMain && ctx.vydel != null && looksLikeMainNoId(cells)) {
                const coeff = (cells[2] || "").trim();
                cells[1] = String(ctx.vydel);
                cells[2] = "";
                if (!cells[3]) {
                    if (coeff) cells[3] = coeff;
                    if (!cells[3]) {
                        const m = txt.trim().match(/^(\d{1,2}[А-ЯЁA-Z]{1,3})\b/);
                        if (m) cells[3] = m[1];
                    }
                }
                kind = "MAIN";
            }

            if (kind === "MAIN") {
                const comp = extractCompositionToken(cells[3]);
                if (comp) cells[3] = comp;

                cells = normalizeMainCells(cells);
                cells = normalizeZapasAndKlassTov(cells);
                cells = normalizeMainByRanges(cells);
                cells = normalizeMainZapasFromText(cells, txt);
            }

            if (kind === "SINGLE_TREES") cells = normalizeSingleTreesCells(cells);

            cells = clearDuplicate24(cells);

            // MAIN/OBJECT write
            if (kind === "MAIN" || kind === "OBJECT") {
                ctx.vydel = toNum(cells[1]) ?? ctx.vydel;
                ctx.inSingleTrees = false;

                let objNote = null;
                if (kind === "OBJECT") {
                    const split = splitObjectDescriptionAndNote(cells[3] || "", cells[24] || "");
                    cells[3] = split.description ?? "";
                    cells[24] = split.cleaned24 ?? "";
                    objNote = split.note;
                    cells = clearDuplicate24(cells);
                }

                const named = mapCellsToNamed(cells);
                pushRecord(
                    records,
                    buildRecord({ ctx, rowNo, kind, page: p, vydel: named.vydel }, {
                        area: named.area,
                        description: named.description,
                        yarus: named.yarus,
                        yarusHigth: named.yarusHigth,
                        element: named.element,
                        age: named.age,
                        higth: named.higth,
                        diam: named.diam,
                        ageKlass: named.ageKlass,
                        ageGroup: named.ageGroup,
                        bonitet: named.bonitet,
                        forestType: named.forestType,
                        polnota: named.polnota,
                        zapasNa1ga: named.zapasNa1ga,
                        ZapasObshiy: named.ZapasObshiy,
                        zapasPoSostavu: named.zapasPoSostavu,
                        klassTovarnosti: named.klassTovarnosti,
                        suhostoy: named.suhostoy,
                        redin: named.redin,
                        edinichnye: named.edinichnye,
                        zakhl: named.zakhl,
                        zakhlLikvid: named.zakhlLikvid,
                        hozMeropriyatiya: named.hozMeropriyatiya,
                        note: objNote,
                        raw: cellsRaw(cells),
                    })
                );

                ctx.pendingMain = null;
                ctx.lastMainIndex = records.length - 1;
                continue;
            }

            if (ctx.vydel == null) continue;

            // NOTE/TEXT: note всегда из txt (сохранение порядка слов)
            if (kind === "NOTE" || kind === "TEXT") {
                if (kind === "TEXT" && ctx.lastMainIndex != null && looksLikeDescriptionContinuationByC3(cells)) {
                    const extra = (cells[3] || "").trim();
                    const prev = records[ctx.lastMainIndex];
                    if (prev && prev.kind === "MAIN" && prev.vydel === ctx.vydel && extra) {
                        prev.description = normalizeSpaces([prev.description, extra].filter(Boolean).join(" "));
                        prev.raw = clampCellText(normalizeSpaces([prev.raw, `| cont:${extra}`].join(" ")));
                        continue;
                    }
                }

                pushRecord(
                    records,
                    buildRecord({ ctx, rowNo, kind, page: p, vydel: ctx.vydel }, {
                        note: txt,
                        raw: cellsRaw(cells),
                    })
                );
                continue;
            }

            // SPECIES / SINGLE_TREES
            if (kind === "SPECIES" || kind === "SINGLE_TREES") {
                if (kind === "SPECIES") cells = normalizeSpeciesCells(cells);
                const named = mapCellsToNamed(cells);

                let description = null;
                let note = null;
                if (kind === "SINGLE_TREES") {
                    const coeff = (cells[2] || "").trim();
                    description = coeff || null;
                    note = "Единичные деревья";
                }

                pushRecord(
                    records,
                    buildRecord({ ctx, rowNo, kind, page: p, vydel: ctx.vydel }, {
                        description,
                        note,
                        yarus: named.yarus,
                        yarusHigth: named.yarusHigth,
                        element: named.element,
                        age: named.age,
                        higth: named.higth,
                        diam: named.diam,
                        ageKlass: named.ageKlass,
                        ageGroup: named.ageGroup,
                        bonitet: named.bonitet,
                        forestType: named.forestType,
                        polnota: named.polnota,
                        zapasNa1ga: named.zapasNa1ga,
                        ZapasObshiy: named.ZapasObshiy,
                        zapasPoSostavu: named.zapasPoSostavu,
                        klassTovarnosti: named.klassTovarnosti,
                        suhostoy: named.suhostoy,
                        redin: named.redin,
                        edinichnye: named.edinichnye,
                        zakhl: named.zakhl,
                        zakhlLikvid: named.zakhlLikvid,
                        hozMeropriyatiya: named.hozMeropriyatiya,
                        raw: cellsRaw(cells),
                    })
                );
            }
        }
    }

    return records;
}

/* ---------------- excel writer ---------------- */
async function writeExcelOneSheet(records, outXlsxPath) {
    const wb = new ExcelJS.Workbook();
    wb.creator = "tax-parser";
    const ws = wb.addWorksheet("Data");

    ws.columns = [
        { header: "quarter", key: "quarter", width: 8 },
        { header: "rowNo", key: "rowNo", width: 6 },
        { header: "category", key: "category", width: 30 },
        { header: "vydel", key: "vydel", width: 8 },
        { header: "kind", key: "kind", width: 22 },
        { header: "page", key: "page", width: 6 },

        { header: "area", key: "area", width: 10 },
        { header: "description", key: "description", width: 30 },

        { header: "yarus", key: "yarus", width: 8 },
        { header: "yarusHigth", key: "yarusHigth", width: 12 },
        { header: "element", key: "element", width: 10 },
        { header: "age", key: "age", width: 8 },
        { header: "higth", key: "higth", width: 8 },
        { header: "diam", key: "diam", width: 8 },
        { header: "ageKlass", key: "ageKlass", width: 10 },
        { header: "ageGroup", key: "ageGroup", width: 10 },
        { header: "bonitet", key: "bonitet", width: 10 },
        { header: "forestType", key: "forestType", width: 12 },
        { header: "polnota", key: "polnota", width: 10 },
        { header: "zapasNa1ga", key: "zapasNa1ga", width: 11 },
        { header: "ZapasObshiy", key: "ZapasObshiy", width: 12 },
        { header: "zapasPoSostavu", key: "zapasPoSostavu", width: 14 },

        { header: "klassTovarnosti", key: "klassTovarnosti", width: 14 },
        { header: "suhostoy", key: "suhostoy", width: 10 },
        { header: "redin", key: "redin", width: 10 },
        { header: "edinichnye", key: "edinichnye", width: 12 },
        { header: "zakhl", key: "zakhl", width: 10 },
        { header: "zakhlLikvid", key: "zakhlLikvid", width: 12 },
        { header: "hozMeropriyatiya", key: "hozMeropriyatiya", width: 20 },

        { header: "note", key: "note", width: 30 },
        { header: "raw", key: "raw", width: 80 },
    ];

    for (const r of records) ws.addRow(r);

    ws.views = [{ state: "frozen", ySplit: 1 }];
    ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: ws.columns.length } };

    await wb.xlsx.writeFile(outXlsxPath);
}

/* ===================== export for Electron main ===================== */
export async function convertPdfToXlsx(pdfPath, outXlsxPath, onProgress) {
    onProgress?.({ stage: "init", current: 0, total: 1, message: "Чтение PDF…" });
    const records = await parsePdf(pdfPath, onProgress);
    onProgress?.({ stage: "excel", current: 0, total: 1, message: "Запись Excel…" });
    await writeExcelOneSheet(records, outXlsxPath);
    onProgress?.({ stage: "done", current: 1, total: 1, message: "Готово" });
    return outXlsxPath;
}