import fs from "fs";
import ExcelJS from "exceljs";
import * as pdfjsLib from "pdfjs-dist/legacy/build/pdf.mjs";
import { DICTS, canon, _internal } from "./dicts.js";

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

function extractCompositionFromTxtAnywhere(txt) {
    const t = normalizeBrokenWords(String(txt ?? "")).toUpperCase();

    const m = t.match(
        /\d{1,2}\s*[А-ЯЁA-Z]{1,6}(?:\s*\d{1,2}\s*[А-ЯЁA-Z]{1,6})*(?:\s*\+\s*[А-ЯЁA-Z]{1,6}(?:\s*,\s*[А-ЯЁA-Z]{1,6})*)*/
    );
    return m ? m[0].replace(/\s+/g, "") : null;
}

/* ---------------- totals by txt (не зависит от grid) ---------------- */
// function isTotalsTitleByTxt(txt) {
//     const t = normalizeBrokenWords(String(txt ?? "")).trim();
//     if (/^Итого по категории/i.test(t)) return "TOTAL_CATEGORY";
//     if (/^Итого по кварталу/i.test(t)) return "TOTAL_QUARTER";
//     if (/^По составляющим породам/i.test(t)) return "TOTAL_SPECIES_HEADER";
//     return null;
// }

// function extractMainDescriptionFromTxt(txt) {
//     // "5 25.4 7Б1ОС2Е+Е 1 16 Б ..."
//     // Берём всё между "vydel area" и " 1..5  <число> "
//     const t = normalizeBrokenWords(String(txt ?? "")).trim();
//     const m = t.match(/^\d{1,4}\s+\d+(?:[.,]\d+)?\s+(.+?)\s+[1-5]\s+\d{1,2}\b/);
//     if (!m) return null;
//     return normalizeSpaces(m[1]);
// }

// function extractMainDescriptionFromTxt(txt) {
//     const t = normalizeBrokenWords(String(txt ?? "")).trim();

//     // Убираем "vydel area"
//     const m = t.match(/^\d{1,4}\s+\d+(?:[.,]\d+)?\s+(.+)$/);
//     if (!m) return null;
//     const rest = m[1];

//     // Находим первую позицию "yarus (1..5) + №выдела/участка (1..2 цифры)"
//     // Важно: не требуем пробела перед yarus (в PDF часто склейка)
//     const re = /[1-5]\s+\d{1,2}\b/;
//     const mm = re.exec(rest);
//     if (!mm) return null;

//     const desc = normalizeSpaces(rest.slice(0, mm.index));
//     return desc || null;
// }

function isTotalsTitleByTxt(txt) {
    const s = String(txt ?? "");

    // 1) обычная нормализация как раньше
    const t = normalizeBrokenWords(s).trim();
    if (/^Итого по категории/i.test(t)) return "TOTAL_CATEGORY";
    if (/^Итого по кварталу/i.test(t)) return "TOTAL_QUARTER";
    if (/^По составляющим породам/i.test(t)) return "TOTAL_SPECIES_HEADER";

    // 2) ✅ устойчиво к разрывам внутри слов: "Ито го", "кат егории"
    const compact = s
        .toLowerCase()
        .replace(/ё/g, "е")
        .replace(/[^a-zа-я0-9]+/g, ""); // убрали пробелы/точки/дефисы

    if (compact.startsWith("итогопокатегории")) return "TOTAL_CATEGORY";
    if (compact.startsWith("итогопокварталу")) return "TOTAL_QUARTER";
    if (compact.startsWith("посоставляющимпородам")) return "TOTAL_SPECIES_HEADER";

    return null;
}

function extractLeadingCompositionFromTxt(txt) {
    const t = normalizeBrokenWords(String(txt ?? "")).trim();

    // берём всё до маркера "ярус высота" (например " 1 11 ")
    const m = t.match(/^(.+?)\s+[1-5]\s+\d{1,2}\b/);
    if (!m) return null;

    // в начале строки может быть только формула, иногда с пробелами — уберём пробелы
    const head = normalizeSpaces(m[1]);
    if (!head) return null;

    // вытащим именно токен состава (ваш extractCompositionToken уже должен уметь + и ,)
    const comp = extractCompositionToken(head);
    return comp ? comp : head.replace(/\s+/g, "");
}

// function extractMainDescriptionFromTxt(txt, cells) {
//     const t = normalizeBrokenWords(String(txt ?? "")).trim();

//     // rest = всё после "vydel area"
//     const m = t.match(/^\d{1,4}\s+\d+(?:[.,]\d+)?\s+(.+)$/);
//     if (!m) return null;
//     const rest = m[1];

//     const yarus = toIntStrict(cells?.[4]);
//     const yarusH = toIntStrict(cells?.[5]);

//     // если есть ярус и высота яруса — режем по ним
//     if (yarus != null && yarusH != null) {
//         const re = new RegExp(`\\b${yarus}\\s+${yarusH}\\b`);
//         const mm = re.exec(rest);
//         if (mm) {
//             const desc = normalizeSpaces(rest.slice(0, mm.index));
//             return desc || null;
//         }
//     }

//     // fallback: старый общий маркер "1..5 + 1..2 цифры"
//     const mm2 = /[1-5]\s+\d{1,2}\b/.exec(rest);
//     if (!mm2) return null;

//     const desc = normalizeSpaces(rest.slice(0, mm2.index));
//     return desc || null;
// }

function extractMainDescriptionFromTxt(txt, cells) {
    const t = normalizeBrokenWords(String(txt ?? "")).trim();

    let rest = null;

    // Вариант 1: нормальная площадь "38.7"
    {
        const m = t.match(/^\d{1,4}\s+\d+(?:[.,]\d+)?\s+(.+)$/);
        if (m) rest = m[1];
    }

    // Вариант 2: разорванная площадь "38 .7" (или "38 ,7")
    if (!rest) {
        const m = t.match(/^\d{1,4}\s+\d+\s+[.,]\d+\s+(.+)$/);
        if (m) rest = m[1];
    }

    if (!rest) return null;

    // На всякий случай: если в начало всё же попала ".7 " — уберём
    rest = rest.replace(/^[.,]\d+\s+/, "");

    // Дальше режем по "ярус + высота", если они известны из cells
    const yarus = toIntStrict(cells?.[4]);
    const yarusH = toIntStrict(cells?.[5]);

    if (yarus != null && yarusH != null) {
        const re = new RegExp(`\\b${yarus}\\s+${yarusH}\\b`);
        const mm = re.exec(rest);
        if (mm) {
            const desc = normalizeSpaces(rest.slice(0, mm.index));
            return desc || null;
        }
    }

    // fallback: старый маркер "1..5 + число"
    const mm2 = /[1-5]\s+\d{1,2}\b/.exec(rest);
    if (!mm2) return normalizeSpaces(rest) || null;

    const desc = normalizeSpaces(rest.slice(0, mm2.index));
    return desc || null;
}

/* ---------------- vydel header by txt ---------------- */
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

/* ---------------- composition helpers (for cultures) ---------------- */
// function extractCompositionToken(s) {
//     const t = normalizeSpaces(s).toUpperCase();
//     const m = t.match(
//         /\b\d{1,2}[А-ЯЁA-Z]{1,3}(?:\s*\d{1,2}[А-ЯЁA-Z]{1,3})*(?:\s*\+\s*[А-ЯЁA-Z]{1,3})*\b/
//     );
//     return m ? m[0].replace(/\s+/g, "") : null;
// }

function extractCompositionToken(s) {
    const t = normalizeSpaces(s).toUpperCase();

    // допускаем:
    // 8Е2Б+ОС,Е
    // 7Б1ОС2Е+Е
    // 5Е 3Б 2ОС (если между сегментами пробелы)
    // 6Е2Б2ИВ+ОЛСА,ОС
    const m = t.match(
        /\b\d{1,2}\s*[А-ЯЁA-Z]{1,6}(?:\s*\d{1,2}\s*[А-ЯЁA-Z]{1,6})*(?:\s*\+\s*[А-ЯЁA-Z]{1,6}(?:\s*,\s*[А-ЯЁA-Z]{1,6})*)*\b/
    );

    return m ? m[0].replace(/\s+/g, "") : null;
}

// function isPureCompositionLine(txt) {
//     const t = normalizeSpaces(txt).toUpperCase();
//     const re = /\b\d{1,2}[А-ЯЁA-Z]{1,3}(?:\s*\d{1,2}[А-ЯЁA-Z]{1,3})*(?:\s*\+\s*[А-ЯЁA-Z]{1,3})*\b/;
//     const m = re.exec(t);
//     if (!m) return null;
//     const rawComp = m[0];
//     const comp = rawComp.replace(/\s+/g, "");
//     const rest = normalizeSpaces(t.replace(rawComp, " ").replace(/[.,;:()]/g, " "));
//     return rest ? null : comp;
// }

function isPureCompositionLine(txt) {
    const t = normalizeSpaces(txt).toUpperCase();
    const re = /\b\d{1,2}\s*[А-ЯЁA-Z]{1,6}(?:\s*\d{1,2}\s*[А-ЯЁA-Z]{1,6})*(?:\s*\+\s*[А-ЯЁA-Z]{1,6}(?:\s*,\s*[А-ЯЁA-Z]{1,6})*)*\b/;
    const m = re.exec(t);
    if (!m) return null;

    const rawComp = m[0];
    const comp = rawComp.replace(/\s+/g, "");
    const rest = normalizeSpaces(t.replace(rawComp, " ").replace(/[.,;:()]/g, " "));
    return rest ? null : comp;
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
// function extractQuarter(linesText) {
//     for (const s of linesText) {
//         const m = s.match(/Квартал\s+(\d+)/i);
//         if (m) return Number(m[1]);
//     }
//     return null;
// }

function extractQuarter(linesText) {
    const toCompact = (s) =>
        String(s ?? "")
            .toLowerCase()
            .replace(/ё/g, "е")
            .replace(/[^a-zа-я0-9]+/g, ""); // всё склеиваем

    for (let i = 0; i < linesText.length; i++) {
        const s1 = String(linesText[i] ?? "");
        const s2 = String(linesText[i + 1] ?? "");
        const s3 = String(linesText[i + 2] ?? "");

        // проверяем 1 строку, 2 строки, 3 строки (на случай когда "Кв а р т а л" отдельно)
        const candidates = [
            s1,
            `${s1} ${s2}`,
            `${s1} ${s2} ${s3}`,
        ];

        for (const c of candidates) {
            const compact = toCompact(c);

            // ищем "квартал" + 1..3 цифры
            const m = compact.match(/квартал(\d{1,3})/);
            if (m) {
                const q = Number(m[1]);
                if (Number.isFinite(q)) return q;
            }
        }
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

function splitLineIntoCells(line, grid) {
    const cells = Array.from({ length: 25 }, () => "");

    if (!grid) {
        cells[24] = lineToText(line);
        return cells;
    }

    const anchorsArr = Array.from(grid.anchors.entries())
        .map(([col, x]) => ({ col, x }))
        .sort((a, b) => a.col - b.col);

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

        const c = best.col;
        cells[c] = (cells[c] ? cells[c] + " " : "") + piece;
    }

    for (let i = 1; i <= 24; i++) cells[i] = cells[i].replace(/\s+/g, " ").trim();
    return cells;
}

function normalizeSpacedDecimals(cells) {
    const fix = (s) =>
        String(s ?? "")
            .replace(/(\d)\s*[.,]\s*(\d)/g, "$1.$2") // "0. 5" -> "0.5", "0 ,5" -> "0.5"
            .replace(/\s+/g, " ")
            .trim();

    // трогаем только числовые колонки (чтобы не ломать description)
    const idxs = [2, 14, 15, 16, 17, 19, 20, 21, 22, 23];
    for (const i of idxs) if (cells[i]) cells[i] = fix(cells[i]);

    return cells;
}

function normalizeVydelAreaSplitAcrossC1C2(cells) {
    const c1 = (cells[1] || "").trim();
    const c2 = (cells[2] || "").trim();
    const c3 = (cells[3] || "").trim();

    const setDesc = (tail) => {
        const t = (tail || "").trim();
        if (!t) return;

        const cur3 = (cells[3] || "").trim();
        // не дублируем "Ручьи Ручьи"
        if (_internal.normKey(t) === _internal.normKey(cur3)) {
            cells[3] = cur3 || t;
        } else {
            cells[3] = normalizeSpaces(`${t} ${cur3}`.trim());
        }
    };

    // ---------- CASE 1: c1="52 0", c2=".3 Ручьи" ----------
    {
        const m1 = c1.match(/^(\d{1,4})\s+(\d+)$/);
        const m2 = c2.match(/^[.,](\d+)(?:\s+(.*))?$/);
        if (m1 && m2) {
            const vydel = m1[1];
            const intPart = m1[2];
            const frac = m2[1];
            const tail = (m2[2] || "").trim();

            cells[1] = vydel;
            cells[2] = `${intPart}.${frac}`;
            if (tail) setDesc(tail);
            return cells;
        }
    }

    // дальше имеет смысл только если c1 = целый vydel
    if (!/^\d{1,4}$/.test(c1)) return cells;

    // ---------- CASE 2: c2="0", c3=".3 Ручьи" ----------
    {
        const m2 = c2.match(/^(\d+)$/); // целая часть площади
        const m3 = c3.match(/^[.,](\d+)(?:\s+(.*))?$/); // дробь и хвост
        if (m2 && m3) {
            const intPart = m2[1];
            const frac = m3[1];
            const tail = (m3[2] || "").trim();

            cells[2] = `${intPart}.${frac}`;
            cells[3] = "";       // хвост переедет в description через setDesc
            if (tail) setDesc(tail);
            return cells;
        }
    }

    // ---------- CASE 3: c2="0." (или "0,"), c3="3 Ручьи" ----------
    {
        const m2 = c2.match(/^(\d+)[.,]$/);
        const m3 = c3.match(/^(\d+)(?:\s+(.*))?$/);
        if (m2 && m3) {
            const intPart = m2[1];
            const frac = m3[1];
            const tail = (m3[2] || "").trim();

            cells[2] = `${intPart}.${frac}`;
            cells[3] = "";
            if (tail) setDesc(tail);
            return cells;
        }
    }

    // ---------- CASE 4: c2=".3", c3="Ручьи"  -> area="0.3" ----------
    {
        const m2 = c2.match(/^[.,](\d+)$/);
        if (m2) {
            const frac = m2[1];
            cells[2] = `0.${frac}`;
            // c3 уже содержит хвост (например "Ручьи"), оставляем
            return cells;
        }
    }

    return cells;
}

function normalizePolnotaAndZapasNa1gaShift(cells) {
    // кейс: 15="0.5 21" -> 14="0.5", 15="21"
    if (!(cells[14] || "").trim()) {
        const s15 = (cells[15] || "").trim().replace(",", ".");
        const m = s15.match(/^(0\.\d|1(?:\.0)?)\s+(\d+(?:\.\d+)?)$/);
        if (m) {
            cells[14] = m[1];
            cells[15] = m[2];
        }
    }
    return cells;
}

function normalizeAgeHeightSplitFromC7(cells) {
    // кейс: 7="120 26" -> 7="120", 8="26"
    if ((cells[7] || "").trim() && !(cells[8] || "").trim()) {
        const m = String(cells[7]).trim().match(/^(\d{1,3})\s+(\d{1,2})$/);
        if (m) {
            cells[7] = m[1];
            cells[8] = m[2];
        }
    }
    return cells;
}

function normalizeBonitetShiftFrom13(cells) {
    // если бонитет уже есть — не трогаем
    if ((cells[12] || "").trim()) return cells;

    const c13 = normalizeSpaces(cells[13]);
    if (!c13) return cells;

    // ловим "5А ..." или "5а ..." или "1 ... 5 ..."
    const m = c13.match(/^(5[аА]|[1-5])\s+(.+)$/);
    if (!m) return cells;

    const b = normalizeBonitet(m[1]); // вернёт "5А" или "1".."5"
    if (!b) return cells;

    cells[12] = b;
    cells[13] = normalizeSpaces(m[2]); // остаток -> forestType/tlu
    return cells;
}

function normalizeSpacedIntegersInNumericCols(cells) {
    const join = (s) => {
        const t = String(s ?? "").trim();
        // "1 8" -> "18", "1 4 0" -> "140"
        if (!/^\d(?:\s+\d){1,2}$/.test(t)) return s;
        return t.replace(/\s+/g, "");
    };

    // числовые колонки, где такое реально встречается
    const idxs = [4, 5, 7, 8, 9, 10, 11, 12, 15, 16, 17, 18, 19, 20, 21, 22, 23];
    for (const i of idxs) {
        if (cells[i]) cells[i] = join(cells[i]);
    }
    return cells;
}

function normalizeSpeciesAgeHeightShift(cells) {
    // если age пустой, но есть числа в 8/9 — трактуем как age/height
    const a7 = (cells[7] || "").trim();
    const a8 = (cells[8] || "").trim();
    const a9 = (cells[9] || "").trim();

    if (!a7 && /^\d{1,3}$/.test(a8) && /^\d{1,2}$/.test(a9)) {
        const age = Number(a8);
        const h = Number(a9);
        if (age >= 1 && age <= 300 && h >= 1 && h <= 99) {
            cells[7] = a8;
            cells[8] = a9;
            cells[9] = ""; // diam отсутствует в исходной строке (как у "ОС 26 40 ...")
        }
    }
    return cells;
}

function normalizeSpeciesAgeHeightSplitFromC8(cells) {
    // нужен элемент породы, иначе лучше не трогать
    const el = (cells[6] || "").trim();
    if (!el) return cells;

    // если age уже есть — не трогаем
    if ((cells[7] || "").trim()) return cells;

    const s8 = String(cells[8] || "").trim();
    if (!s8) return cells;

    // "100 24" -> age=100, height=24
    const m = s8.match(/^(\d{1,3})\s+(\d{1,2})$/);
    if (!m) return cells;

    const age = Number(m[1]);
    const h = Number(m[2]);
    if (!(Number.isInteger(age) && age >= 1 && age <= 300)) return cells;
    if (!(Number.isInteger(h) && h >= 1 && h <= 99)) return cells;

    // важно: diam уже может быть в 9 — оставляем как есть
    cells[7] = String(age);
    cells[8] = String(h);
    return cells;
}

function normalizeSpeciesAgeHeightSplitFromC8(cells) {
    // нужен элемент породы, иначе лучше не трогать
    const el = (cells[6] || "").trim();
    if (!el) return cells;

    // если age уже есть — не трогаем
    if ((cells[7] || "").trim()) return cells;

    const s8 = String(cells[8] || "").trim();
    if (!s8) return cells;

    // CASE A: "100 24" -> age=100, height=24
    {
        const m = s8.match(/^(\d{1,3})\s+(\d{1,2})$/);
        if (m) {
            const age = Number(m[1]);
            const h = Number(m[2]);
            if (
                Number.isInteger(age) && age >= 1 && age <= 300 &&
                Number.isInteger(h) && h >= 1 && h <= 99
            ) {
                cells[7] = String(age);
                cells[8] = String(h);
                return cells;
            }
        }
    }

    // CASE B: "100 24 28" -> age=100, height=24, diam=28
    {
        const m = s8.match(/^(\d{1,3})\s+(\d{1,2})\s+(\d{1,3})$/);
        if (m) {
            const age = Number(m[1]);
            const h = Number(m[2]);
            const d = Number(m[3]);
            if (
                Number.isInteger(age) && age >= 1 && age <= 300 &&
                Number.isInteger(h) && h >= 1 && h <= 99 &&
                Number.isInteger(d) && d >= 1 && d <= 150
            ) {
                cells[7] = String(age);
                cells[8] = String(h);

                // diam кладём в 9 только если там пусто (чтобы не затереть уже правильно распознанное)
                if (!String(cells[9] || "").trim()) cells[9] = String(d);

                // чистим 8-ю от "пачки"
                // (мы уже положили нормальные значения в 7/8/9)
                // оставляем cells[8] как height
                return cells;
            }
        }
    }

    return cells;
}

/* ---------------- per-line normalizers ---------------- */

function normalizeHozMeropriyatiyaShiftFrom23(cells) {
    const c23 = (cells[23] || "").trim();
    const c24 = (cells[24] || "").trim();
    if (!c23) return cells;

    // 23 должна быть ЧИСЛОМ (захламленность ликвида), 24 — ТЕКСТ (хоз. мероприятия)
    const hasLetters = /[A-Za-zА-ЯЁа-яё]/.test(c23);
    const hasPercent = /%/.test(c23);

    // чистое число (без %) считаем валидным для 23
    const n23 = toFloatLoose(c23);
    const isPureNumeric23 = n23 != null && !hasLetters && !hasPercent;

    // если это не число (или это проценты/текст) — переносим в 24
    if (!isPureNumeric23) {
        cells[24] = c24 ? normalizeSpaces(`${c24} ${c23}`) : c23;
        cells[23] = "";
    }

    return cells;
}

function normalizeAreaSplitIntoC3(cells) {
    // применяем только для строк, где есть №выдела в col1
    if (!/^\d+$/.test((cells[1] || "").trim())) return cells;

    const a2 = (cells[2] || "").trim(); // площадь (иногда только целая часть)
    const c3 = (cells[3] || "").trim(); // иногда начинается с ".6 ..."

    if (!a2 || !c3) return cells;

    // кейс: "36" + ".6 7Б..."  => area=36.6, desc="7Б..."
    if (/^\d+$/.test(a2)) {
        const m = c3.match(/^([.,]\d+)(?:\s+(.*))?$/);
        if (m) {
            cells[2] = `${a2}${m[1].replace(",", ".")}`;
            cells[3] = (m[2] || "").trim();
            return cells;
        }
    }

    // кейс: "36." + "6 7Б..." => area=36.6, desc="7Б..."
    if (/^\d+[.,]$/.test(a2)) {
        const m = c3.match(/^(\d+)(?:\s+(.*))?$/);
        if (m) {
            cells[2] = `${a2}${m[1]}`.replace(",", ".");
            cells[3] = (m[2] || "").trim();
            return cells;
        }
    }

    return cells;
}


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

    if (/^\d{1,4}$/.test(t)) return true;
    if (t.includes("----------------------------------------------------------------")) return true;
    if (/^[-]{5,}$/.test(t)) return true;
    if (t.trim().startsWith(":")) return true;

    // robust header detection (разорванные слова)
    const compact = t.replace(/\s+/g, "").toLowerCase();
    if (compact.includes("категориязащитности") && compact.includes("квартал")) return true;

    // many ':' means header grid
    const colonCount = (t.match(/:/g) || []).length;
    if (colonCount >= 10) return true;

    return false;
}

/* ---------------- helpers ---------------- */

function parseShortMainNoIdFromTxt(txt) {
    const t = normalizeBrokenWords(txt).trim();

    const m = t.match(
        /^(\d{1,2}\s*[А-ЯЁA-Z]{1,6}(?:\s*\d{1,2}\s*[А-ЯЁA-Z]{1,6})*(?:\s*\+\s*[А-ЯЁA-Z]{1,6}(?:\s*,\s*[А-ЯЁA-Z]{1,6})*)*)\s+(\d{1,2})\s+([А-ЯЁA-Z]{1,6})\s+(\d{1,3})\s+(\d{1,2})\s+(\d{1,3})\s+(\d{1,2})\s+(\d+(?:[.,]\d+)?)$/i
    );
    if (!m) return null;

    return {
        comp: m[1].replace(/\s+/g, ""),
        yarusH: Number(m[2]),
        el: m[3].toUpperCase(),
        age: Number(m[4]),
        h: Number(m[5]),
        d: Number(m[6]),
        ageKlass: Number(m[7]),
        zapasNa1ga: Number(m[8].replace(",", ".")),
    };
}



function moveTluToLastMain(ctx, records, tlu) {
    if (!tlu) return false;
    if (ctx.lastMainIndex == null) return false;

    const prev = records[ctx.lastMainIndex];
    if (!prev || prev.kind !== "MAIN") return false;
    if (prev.vydel !== ctx.vydel) return false;

    if (!prev.tlu) {
        prev.tlu = tlu;
        return true;
    }
    return false;
}

function isEmptyRange(cells, a, b) {
    for (let i = a; i <= b; i++) if (cells[i]) return false;
    return true;
}
function hasAnyRange(cells, a, b) {
    for (let i = a; i <= b; i++) if (cells[i]) return true;
    return false;
}
// function looksLikePolnota(cells) {
//     return /^\d\.\d$/.test((cells[14] || "").replace(",", "."));
// }

function looksLikePolnota(cells) {
    return isPolnotaStr(cells[14]);
}

function looksLikeStandDescription(desc) {
    const s = (desc ?? "").toString().toLowerCase();
    return s.includes("насажд") || s.includes("культ") || s.includes("полог");
}

// function looksLikeMainNoId(cells) {
//     if ((cells[1] || "").trim()) return false;
//     if (!looksLikePolnota(cells)) return false;

//     const y = toIntStrict(cells[4]);
//     if (cells[4] && !(y != null && y >= 1 && y <= 5)) return false;

//     const el = (cells[6] || "").trim();
//     const hasElement = !!el && /^[А-ЯЁA-Z]{1,6}$/.test(el.replace(/\s+/g, ""));

//     const age = toIntStrict(cells[7]);
//     const h = toIntStrict(cells[8]);
//     const d = toIntStrict(cells[9]);

//     const hasDims = (age != null && age <= 300) || (h != null && h <= 99) || (d != null && d <= 150);

//     return hasElement || hasDims;
// }

// function looksLikeMainNoId(cells) {
//     if ((cells[1] || "").trim()) return false;

//     const coeff = (cells[2] || "").trim();
//     const coeffCompact = coeff.replace(/\s+/g, "").toUpperCase();

//     const coeffIsComp = !!coeffCompact && /\d/.test(coeffCompact) && /[А-ЯЁA-Z]/.test(coeffCompact);

//     // старая логика: если есть полното́та — работаем как раньше
//     if (looksLikePolnota(cells)) {
//         const y = toIntStrict(cells[4]);
//         if (cells[4] && !(y != null && y >= 1 && y <= 5)) return false;

//         const el = (cells[6] || "").trim();
//         const hasElement = !!el && /^[А-ЯЁA-Z]{1,6}$/.test(el.replace(/\s+/g, ""));

//         const age = toIntStrict(cells[7]);
//         const h = toIntStrict(cells[8]);
//         const d = toIntStrict(cells[9]);
//         const hasDims = (age != null && age <= 300) || (h != null && h <= 99) || (d != null && d <= 150);

//         return hasElement || hasDims;
//     }

//     // ✅ новая логика: polnota нет, но похоже на строку яруса (состав + метрики)
//     if (coeffIsComp) {
//         const el = (cells[6] || "").trim();
//         const hasElement = !!el && /^[А-ЯЁA-Z]{1,6}$/.test(el.replace(/\s+/g, ""));

//         const age = toIntStrict(cells[7]);
//         const h = toIntStrict(cells[8]);
//         const d = toIntStrict(cells[9]);
//         const hasDims = (age != null && age <= 300) || (h != null && h <= 99) || (d != null && d <= 150);

//         const hasAnyMetrics = hasAnyRange(cells, 5, 23); // справа что-то числовое/значимое

//         if (hasAnyMetrics && (hasElement || hasDims)) return true;
//     }

//     return false;
// }

function looksLikeMainNoId(cells) {
    // если слева уже есть №выдела — это не "без-Id"
    if ((cells[1] || "").trim()) return false;

    // 1) состав может сидеть и в col2 (coeff), и в col3 (description)
    const compCandidate = ((cells[2] || "").trim() || (cells[3] || "").trim());
    const compCompact = compCandidate.replace(/\s+/g, "").toUpperCase();

    // грубый признак состава: есть цифры и буквы
    const compLooksLike = !!compCompact && /\d/.test(compCompact) && /[А-ЯЁA-Z]/.test(compCompact);

    // 2) элемент (порода) в col6
    const el = (cells[6] || "").trim();
    const hasElement = !!el && /^[А-ЯЁA-Z]{1,6}$/.test(el.replace(/\s+/g, ""));

    // 3) размеры/возраст — учитываем склейки типа "120 27" в col8/col7
    const age = toIntStrict(cells[7]);
    const h = toIntStrict(cells[8]);
    const d = toIntStrict(cells[9]);

    let hasDims = (age != null && age <= 300) || (h != null && h <= 99) || (d != null && d <= 150);

    // если age/height склеены, например cells[8]="120 27"
    if (!hasDims) {
        const m8 = String(cells[8] || "").trim().match(/^(\d{1,3})\s+(\d{1,2})$/);
        if (m8) {
            const a = Number(m8[1]), hh = Number(m8[2]);
            if (Number.isFinite(a) && a <= 300) hasDims = true;
            if (Number.isFinite(hh) && hh <= 99) hasDims = true;
        }
        const m7 = String(cells[7] || "").trim().match(/^(\d{1,3})\s+(\d{1,2})$/);
        if (!hasDims && m7) {
            const a = Number(m7[1]), hh = Number(m7[2]);
            if (Number.isFinite(a) && a <= 300) hasDims = true;
            if (Number.isFinite(hh) && hh <= 99) hasDims = true;
        }
    }

    // 4) если есть полното́та — это сильный сигнал "MAIN"
    if (looksLikePolnota(cells)) {
        const y = toIntStrict(cells[4]);
        if (cells[4] && !(y != null && y >= 1 && y <= 5)) return false;
        return hasElement || hasDims || compLooksLike;
    }

    // 5) без полното́ты: всё равно считаем MAIN, если:
    //    - похоже на состав
    //    - есть хоть какие-то метрики справа
    //    - и есть порода или размеры
    if (compLooksLike) {
        const hasAnyMetrics = hasAnyRange(cells, 5, 23);
        if (hasAnyMetrics && (hasElement || hasDims)) return true;
    }

    return false;
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

function parseTotalsValuesFromCells(cells) {
    return {
        area: toNum(cells[2]),

        ZapasObshiy: toNum(cells[16]),
        zapasPoSostavu: toNum(cells[17]),
        klassTovarnosti: toIntStrict(cells[18]),

        suhostoy: toNum(cells[19]),
        redin: toNum(cells[20]),
        edinichnye: toNum(cells[21]),
        zakhl: toNum(cells[22]),
        zakhlLikvid: toNum(cells[23]),
    };
}

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

// function parseTotalsSpeciesRow(txt) {
//     const s = normalizeBrokenWords(String(txt ?? "").trim()).replace(/\s+/g, " ");
//     const m = s.match(/^([А-ЯЁA-Z]{1,6})(?:\s+(\d+(?:[.,]\d+)?))?$/i);
//     if (!m) return null;

//     const element = m[1].toUpperCase();
//     const zapas = m[2] != null ? Number(m[2].replace(",", ".")) : null;
//     if (m[2] != null && !Number.isFinite(zapas)) return null;

//     return { element, zapasPoSostavu: zapas };
// }

function parseTotalsSpeciesRow(txt) {
    const s = normalizeBrokenWords(String(txt ?? "").trim()).replace(/\s+/g, " ");

    // допускаем разрыв породы пробелом: "ОЛ СА 76"
    const m = s.match(/^([А-ЯЁA-Z\s]{1,20})(?:\s+(\d+(?:[.,]\d+)?))?$/i);
    if (!m) return null;

    const elementRaw = m[1].replace(/\s+/g, "").toUpperCase();
    if (!/^[А-ЯЁA-Z]{1,6}$/.test(elementRaw)) return null;

    const zapas = m[2] != null ? Number(m[2].replace(",", ".")) : null;
    if (m[2] != null && !Number.isFinite(zapas)) return null;

    return { element: elementRaw, zapasPoSostavu: zapas };
}

/* ---------------- normalize zapas / klass ---------------- */
// function normalizeZapasAndKlassTov(cells) {
//     const trySplit = (idx) => {
//         const s = (cells[idx] || "").trim();
//         if (!s) return false;

//         const m = s.match(/^(\d+(?:[.,]\d+)?)\s+(\d+)$/);
//         if (!m) return false;

//         const first = m[1].replace(",", ".");
//         const second = m[2];

//         if (idx === 17) {
//             cells[17] = first;
//             if (!cells[18]) cells[18] = second;
//             return true;
//         }

//         if (idx === 18) {
//             if (!cells[17]) cells[17] = first;
//             cells[18] = second;
//             return true;
//         }
//         return false;
//     };

//     if (!trySplit(17)) trySplit(18);
//     return cells;
// }

function normalizeKlassShiftFrom19(cells) {
    const k18 = toIntStrict(cells[18]);
    const k19 = toIntStrict(cells[19]);
    const isKlass = (n) => n != null && n >= 1 && n <= 4;

    // если в 19 стоит класс, а 18 пустая -> переносим в 18
    if (isKlass(k19) && !(cells[18] || "").trim()) {
        cells[18] = String(k19);
        cells[19] = "";
        return cells;
    }

    // если в 18 уже класс и в 19 тот же класс -> чистим 19
    if (isKlass(k18) && isKlass(k19) && k18 === k19) {
        cells[19] = "";
        return cells;
    }

    return cells;
}

function normalizeZapasAndKlassTov(cells) {
    const normNumStr = (s) => String(s ?? "").trim().replace(/,/g, ".");
    const toFloatStr = (s) => {
        const t = normNumStr(s);
        const m = t.match(/^\d+(?:\.\d+)?$/);
        return m ? t : null;
    };
    const toIntStr = (s) => {
        const t = normNumStr(s);
        return /^\d+$/.test(t) ? t : null;
    };
    const isKlass14 = (s) => {
        const n = toIntStrict(s);
        return n != null && n >= 1 && n <= 4;
    };

    // A) кейс: 17="167", 18=".1 1"  => 17="167.1", 18="1"
    {
        const z17 = toIntStr(cells[17]);
        const s18 = normNumStr(cells[18]);
        const m = s18.match(/^([.]\d+)\s+(\d+)$/);
        if (z17 && m) {
            const frac = m[1];        // ".1"
            const k = m[2];           // "1"
            if (isKlass14(k)) {
                cells[17] = `${z17}${frac}`; // "167.1"
                cells[18] = k;
                return cells;
            }
        }
    }

    // B) если 18 содержит "zapas klass" (например "1041. 2" или "1041 2")
    {
        const s18 = normNumStr(cells[18]);
        const m = s18.match(/^(\d+(?:\.\d+)?)\s+(\d+)$/);
        if (m) {
            const zap = m[1];
            const k = m[2];
            if (Number(zap) > 4 && isKlass14(k)) {
                // переносим zapas -> 17, klass -> 18
                if (!cells[17]) cells[17] = zap;
                else {
                    // если в 17 лежит класс (1..4), а zapas в 18 — меняем местами
                    if (isKlass14(cells[17])) cells[17] = zap;
                }
                cells[18] = k;
                return cells;
            }
        }
    }

    // C) если 17 содержит "zapas klass"
    {
        const s17 = normNumStr(cells[17]);
        const m = s17.match(/^(\d+(?:\.\d+)?)\s+(\d+)$/);
        if (m) {
            const zap = m[1];
            const k = m[2];
            if (Number(zap) > 4 && isKlass14(k)) {
                cells[17] = zap;
                if (!cells[18]) cells[18] = k;
                return cells;
            }
        }
    }

    // D) если класс уехал в 17, а запас в 18: 17 in [1..4] и 18 > 4 -> swap
    {
        const k17 = cells[17];
        const z18 = toFloatLoose(cells[18]);
        if (isKlass14(k17) && z18 != null && z18 > 4) {
            cells[17] = String(z18);
            cells[18] = String(toIntStrict(k17));
            return cells;
        }
    }

    // E) если в 18 лежит просто большой запас, а 17 пустой — перенесём в 17
    {
        const z18 = toFloatLoose(cells[18]);
        if (!cells[17] && z18 != null && z18 > 4 && !isKlass14(cells[18])) {
            cells[17] = String(z18);
            cells[18] = "";
            return cells;
        }
    }

    // F) финальная валидация: klassTovarnosti только 1..4, иначе чистим
    if (cells[18] && !isKlass14(cells[18])) {
        // иногда там мусор типа "1041. 2" — выше должны были разрулить
        // если не разрулили, лучше очистить класс, чем писать неверно
        cells[18] = "";
    }

    return cells;
}

/* ---------------- normalize MAIN ---------------- */
// function normalizeMainCells(cells) {
//     // 9: diam, 10: ageKlass
//     {
//         const s = (cells[9] || "").trim();
//         const m = s.match(/^(\d+(?:[.,]\d+)?)\s+(\d+)$/);
//         if (m) {
//             cells[9] = m[1].replace(",", ".");
//             if (!cells[10]) cells[10] = m[2];
//         }
//     }

//     // 17 may contain "ZapasObshiy zapasPoSostavu klass"
//     {
//         const s17 = (cells[17] || "").trim();
//         const nums = s17.match(/\d+(?:[.,]\d+)?\.?/g);
//         if (nums && nums.length >= 2) {
//             const n1 = nums[0].replace(",", ".");
//             const n2 = nums[1].replace(",", ".");
//             const n3 = nums[2] ? nums[2].replace(",", ".") : null;

//             if (!cells[16]) cells[16] = n1;
//             cells[17] = n2;
//             if (n3 && !cells[18]) cells[18] = n3;
//         }
//     }

//     // 16 may contain "464.3 325"
//     {
//         const s16 = (cells[16] || "").trim();
//         const m = s16.match(/^(\d+(?:[.,]\d+)?)\.?\s+(\d+(?:[.,]\d+)?)\.?$/);
//         if (m) {
//             cells[16] = m[1].replace(",", ".");
//             if (!cells[17]) cells[17] = m[2].replace(",", ".");
//         }
//     }

//     return cells;
// }

function rescueMainMetricsFromTxt(cells, txt) {
    const t = normalizeBrokenWords(String(txt ?? "")).trim();
    if (!t) return cells;

    const tokens = t.split(/\s+/);

    // ищем фрагмент "... <yarus 1..5> <yarusH 1..99> <EL> <AGE> <H> <D> <AK> <AG> <BON> ..."
    for (let i = 0; i <= tokens.length - 6; i++) {
        const yarus = toIntStrict(tokens[i]);
        const yarusH = toIntStrict(tokens[i + 1]);
        const el = tokens[i + 2];

        if (!(yarus != null && yarus >= 1 && yarus <= 5)) continue;
        if (!(yarusH != null && yarusH >= 1 && yarusH <= 99)) continue;
        if (!/^[А-ЯЁA-Z]{1,6}$/i.test(el)) continue;

        // собираем подряд идущие "числовые" токены после element,
        // останавливаемся на первом "словесном" (лесотип и т.п.), но НЕ на "5А"
        const seq = [];
        for (let j = i + 3; j < tokens.length; j++) {
            const tok = tokens[j];

            const isBonitetTok = /^5[аА]$|^[1-5]$/.test(tok);
            const isIntTok = /^\d{1,3}$/.test(tok);

            if (isBonitetTok || isIntTok) {
                seq.push(tok);
                if (seq.length >= 6) break; // нам достаточно
                continue;
            }

            // встретили слово (например "Е", "ЧЕР") — конец блока размеров
            break;
        }

        if (seq.length < 4) return cells; // слишком мало — не рискуем

        const [ageS, hS, dS, akS, agS, bS] = seq;

        const age = toIntStrict(ageS);
        const h = toIntStrict(hS);
        const d = toIntStrict(dS);
        const ak = toIntStrict(akS);
        const ag = toIntStrict(agS);
        const b = normalizeBonitet(bS);

        // заполняем только если пусто/невалидно
        const setIfMissing = (idx, val) => {
            if (val == null) return;
            if (!String(cells[idx] || "").trim()) cells[idx] = String(val);
        };

        // age/height/diam
        if (age != null && age >= 1 && age <= 300) setIfMissing(7, age);
        if (h != null && h >= 1 && h <= 99) setIfMissing(8, h);
        if (d != null && d >= 1 && d <= 150) {
            const curD = toIntStrict(cells[9]);
            const curAG = toIntStrict(cells[11]);
            const curB = normalizeBonitet(cells[12]); // строка

            const curBnum = toIntStrict(curB); // 1..5 или null

            const suspicious =
                curD == null ||                 // пусто/не число
                curD === curAG ||               // diam совпал с ageGroup (частый сдвиг)
                (curBnum != null && curD === curBnum); // diam совпал с bonitet 1..5

            if (suspicious) {
                cells[9] = String(d);
            }
        }

        // ageKlass/ageGroup
        if (ak != null && ak >= 1 && ak <= 12) setIfMissing(10, ak);
        if (ag != null && ag >= 1 && ag <= 10) setIfMissing(11, ag);

        // bonitet
        if (b) setIfMissing(12, b);

        return cells;
    }

    return cells;
}

function normalizeMainCells(cells) {
    const split2ints = (idxA, idxB, aMin, aMax, bMin, bMax) => {
        const s = String(cells[idxA] || "").trim();
        if (!s) return;

        // "80 17" / "17 18" / "2 4"
        const m = s.match(/^(\d{1,3})\s+(\d{1,3})$/);
        if (!m) return;

        const a = Number(m[1]);
        const b = Number(m[2]);

        const okA = Number.isInteger(a) && a >= aMin && a <= aMax;
        const okB = Number.isInteger(b) && b >= bMin && b <= bMax;
        if (!okA || !okB) return;

        // переносим во вторую колонку только если она пустая
        if (!String(cells[idxB] || "").trim()) {
            cells[idxA] = String(a);
            cells[idxB] = String(b);
        }
    };

    // 7 может содержать "age height"
    split2ints(7, 8, 1, 300, 1, 99);

    // 8 может содержать "height diam"
    split2ints(8, 9, 1, 99, 1, 150);

    // 9: diam, 10: ageKlass (ваша логика, оставляем)
    {
        const s = (cells[9] || "").trim();
        const m = s.match(/^(\d+(?:[.,]\d+)?)\s+(\d+)$/);
        if (m) {
            cells[9] = m[1].replace(",", ".");
            if (!cells[10]) cells[10] = m[2];
        }
    }

    // 10 может содержать "ageKlass ageGroup"
    split2ints(10, 11, 1, 12, 1, 10);

    // 11 может содержать "ageGroup bonitet(1..5)" если 12 пустая
    {
        const s11 = String(cells[11] || "").trim();
        if (s11 && !String(cells[12] || "").trim()) {
            const m = s11.match(/^(\d{1,2})\s+(\d{1,2})$/);
            if (m) {
                const g = Number(m[1]);
                const b = Number(m[2]);
                if (Number.isInteger(g) && g >= 1 && g <= 10 && Number.isInteger(b) && b >= 1 && b <= 5) {
                    cells[11] = String(g);
                    cells[12] = String(b);
                }
            }
        }
    }

    // 17 may contain "ZapasObshiy zapasPoSostavu klass" (как у вас)
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

    // 16 may contain "464.3 325" (как у вас)
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

    // запасы из 17 если 16 пустой
    if (!cells[16] && cells[17]) {
        const nums = String(cells[17]).replace(",", ".").match(/\d+(?:\.\d+)?\.?/g);
        if (nums && nums.length >= 2) {
            cells[16] = nums[0].replace(",", ".");
            cells[17] = nums[1].replace(",", ".");
            if (nums[2] && !cells[18]) cells[18] = nums[2];
        }
    }

    // если 16==17 и 18 выглядит как запас -> сдвиг
    {
        const z16 = getFloat(16);
        const z17 = getFloat(17);
        const z18 = getFloat(18);
        const looksZapas = (x) => x != null && x > 9;

        if (z16 != null && z17 != null && Math.abs(z16 - z17) < 1e-9 && looksZapas(z18)) {
            cells[17] = String(z18).replace(",", ".");
            const k18 = getInt(18);
            if (!(k18 != null && k18 >= 1 && k18 <= 4)) cells[18] = "";
        } else {
            const k18 = getInt(18);
            if (cells[18] && !(k18 != null && k18 >= 1 && k18 <= 4)) {
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

    const isKlass = (n) => Number.isInteger(n) && n >= 1 && n <= 4;
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

function extractLonelyElementFromTxt(txt) {
    const t = normalizeBrokenWords(String(txt ?? "")).trim();

    // ВАЖНО: только "ОС" / "Е" / "ИВ" без чисел
    const m = t.match(/^([А-ЯЁA-Z]{1,6})$/i);
    if (!m) return null;

    return canon.species(m[1].toUpperCase()) || null;
}

function formulaHasElement(desc, el) {
    const d = _internal.normKey(desc);
    const e = _internal.normKey(el);
    return !!d && !!e && d.includes(e);
}

function appendExtraElementToMainFormula(prevDesc, el, ctx) {
    const d = normalizeSpaces(prevDesc || "");
    if (!d) return d;

    // если уже есть — ничего не добавляем
    if (formulaHasElement(d, el)) return d;

    const sep = ctx.mainExtraEls && ctx.mainExtraEls.size > 0 ? "," : "+";
    ctx.mainExtraEls?.add(el);
    return d + sep + el;
}

// function appendExtraElementToMainFormula(prevDesc, el, ctx) {
//     const d = normalizeSpaces(prevDesc || "");
//     if (!d) return d;

//     // если уже есть такая порода в формуле — не добавляем
//     if (new RegExp(`\\b${el}\\b`, "i").test(d)) return d;

//     // первый дополнительный элемент через "+", последующие через ","
//     const sep = ctx.mainExtraEls && ctx.mainExtraEls.size > 0 ? "," : "+";
//     ctx.mainExtraEls?.add(el);

//     return d + sep + el;
// }

/* ---------------- SPECIES normalization ---------------- */
function normalizeSpeciesCells(cells) {
    if (cells[6]) cells[6] = String(cells[6]).replace(/\s+/g, "").toUpperCase();
    if (cells[7]) cells[7] = String(cells[7]).trim();

    // ОЛСА50
    {
        const el = (cells[6] || "").trim();
        const m = el.match(/^([А-ЯЁA-Z]{1,6})(\d{1,3})$/i);
        if (m) {
            cells[6] = m[1].toUpperCase();
            if (!cells[7]) cells[7] = m[2];
            // canonize species
            if (cells[6]) cells[6] = canon.species(cells[6]);
            return cells;
        }
    }

    // ОЛСА 50
    {
        const el = (cells[6] || "").trim();
        const m = el.match(/^([А-ЯЁA-Z]{1,6})\s+(\d{1,3})$/i);
        if (m) {
            cells[6] = m[1].toUpperCase();
            if (!cells[7]) cells[7] = m[2];
            if (cells[6]) cells[6] = canon.species(cells[6]);
            return cells;
        }
    }

    // ОЛС + "А 70"
    {
        const el = (cells[6] || "").trim().toUpperCase();
        const a7 = (cells[7] || "").trim();
        const m = a7.match(/^А\s*(\d{1,3})$/i);
        if (el === "ОЛС" && m) {
            cells[6] = "ОЛСА";
            cells[7] = m[1];
            if (cells[6]) cells[6] = canon.species(cells[6]);
            return cells;
        }
    }

    // clean digits from element
    {
        const el = (cells[6] || "").trim();
        if (/\d/.test(el)) {
            const m = el.match(/^([А-ЯЁA-Z]{1,6})/i);
            if (m) cells[6] = m[1].toUpperCase();
        }
    }

    // final canonize
    if (cells[6]) cells[6] = canon.species(cells[6]);

    return cells;
}

function strictDictMatch(dictMap, value) {
    if (!value) return null;
    const k = _internal.normKey(value);
    return dictMap.get(k) ?? null; // null если не нашли
}

// поиск канон. значения по включению в normKey(txt)
// выбираем самое длинное совпадение (ПросекиКварт > Просеки)
function findDictInText(dictMap, txt) {
    const keyTxt = _internal.normKey(txt);
    if (!keyTxt) return null;

    let best = null;
    for (const [k, canonVal] of dictMap.entries()) {
        if (k && keyTxt.includes(k)) {
            if (!best || k.length > best.k.length) best = { k, canonVal };
        }
    }
    return best ? best.canonVal : null;
}

// function getObjectDescriptionFromDict(cells, txt) {
//     return (
//         strictDictMatch(DICTS.OBJECT_KIND, cells?.[3]) ||
//         strictDictMatch(DICTS.OBJECT_KIND, cells?.[24]) ||
//         strictDictMatch(DICTS.OBJECT_KIND, normalizeSpaces(`${cells?.[3] || ""} ${cells?.[24] || ""}`)) ||
//         strictDictMatch(DICTS.OBJECT_KIND, txt) ||
//         findDictInText(DICTS.OBJECT_KIND, txt)
//     );
// }

// function getObjectDescriptionFromDict(cells, txt) {
//     const d1 = strictDictMatch(DICTS.OBJECT_KIND, cells?.[3]);
//     if (d1) return d1;

//     const d2 = strictDictMatch(DICTS.OBJECT_KIND, txt);
//     if (d2) return d2;

//     return null;
// }

function getObjectDescriptionFromDict(cells, txt) {
    return (
        strictDictMatch(DICTS.OBJECT_KIND, cells?.[3]) ||
        strictDictMatch(DICTS.OBJECT_KIND, cells?.[24]) ||
        strictDictMatch(DICTS.OBJECT_KIND, normalizeSpaces(`${cells?.[3] || ""} ${cells?.[24] || ""}`)) ||
        strictDictMatch(DICTS.OBJECT_KIND, txt) ||
        findDictInText(DICTS.OBJECT_KIND, txt)
    );
}

// function getVydelHeaderDescriptionFromDict(cells, txt) {
//     // 1) cells[3]
//     const d1 = strictDictMatch(DICTS.VYDEL_HEADER_KIND, cells?.[3]);
//     if (d1) return d1;

//     // 2) cells[24] (часто туда улетает "Культуры")
//     const d24 = strictDictMatch(DICTS.VYDEL_HEADER_KIND, cells?.[24]);
//     if (d24) return d24;

//     // 3) весь текст строки
//     const d2 = strictDictMatch(DICTS.VYDEL_HEADER_KIND, txt);
//     if (d2) return d2;

//     // 4) комбинированно (на случай если слова разделены между 3 и 24)
//     const combo = normalizeSpaces(`${cells?.[3] || ""} ${cells?.[24] || ""}`);
//     const d3 = strictDictMatch(DICTS.VYDEL_HEADER_KIND, combo);
//     if (d3) return d3;

//     return null;
// }

function getVydelHeaderDescriptionFromDict(cells, txt) {
    const c3 = cells?.[3] || "";
    const c24 = cells?.[24] || "";
    const combo = normalizeSpaces(`${c3} ${c24}`);
    const fullTxt = String(txt || "");

    return (
        // 1) строгие совпадения
        strictDictMatch(DICTS.VYDEL_HEADER_KIND, c3) ||
        strictDictMatch(DICTS.VYDEL_HEADER_KIND, c24) ||
        strictDictMatch(DICTS.VYDEL_HEADER_KIND, combo) ||
        strictDictMatch(DICTS.VYDEL_HEADER_KIND, fullTxt) ||

        // 2) ✅ поиск по ВХОЖДЕНИЮ (когда в хвосте есть лишние слова: "… Проходные рубки")
        findDictInText(DICTS.VYDEL_HEADER_KIND, combo) ||
        findDictInText(DICTS.VYDEL_HEADER_KIND, fullTxt) ||

        null
    );
}


/* ---------------- forestType vs tlu ---------------- */
function parseForestTypeAndTlu(raw13) {
    const s = normalizeSpaces(raw13);
    if (!s) return { forestType: null, tlu: null };

    const key = _internal.normKey(s);

    if (DICTS.FOREST_GROUP.has(key)) {
        return { forestType: canon.forestGroup(s), tlu: null };
    }
    if (DICTS.TLU.has(key)) {
        return { forestType: null, tlu: canon.tlu(s) };
    }

    const parts = s.split(/\s+/g);

    if (parts.length >= 2) {
        const g2 = parts.slice(0, 2).join(" ");
        if (DICTS.FOREST_GROUP.has(_internal.normKey(g2))) {
            const rest = parts.slice(2).join(" ");
            let tlu = null;
            if (rest && DICTS.TLU.has(_internal.normKey(rest))) tlu = canon.tlu(rest);
            return { forestType: canon.forestGroup(g2), tlu };
        }
    }

    if (parts.length === 1 && parts[0].length <= 3) {
        const maybe = parts[0];
        if (DICTS.TLU.has(_internal.normKey(maybe))) {
            return { forestType: null, tlu: canon.tlu(maybe) };
        }
    }

    return { forestType: s, tlu: null };
}

/* ---------------- mapping ---------------- */
function mapCellsToNamed(cells) {
    const { forestType, tlu } = parseForestTypeAndTlu(cells[13]);
    const safeForestType = forestType && /^\d/.test(forestType) ? null : forestType;

    return {
        vydel: toNum(cells[1]),
        area: toNum(cells[2]),
        description: cells[3] || null,

        yarus: toNum(cells[4]),
        yarusHigth: toNum(cells[5]),
        element: cells[6] ? canon.species(cells[6]) : null,
        age: toNum(cells[7]),
        higth: toNum(cells[8]),
        diam: toNum(cells[9]),
        ageKlass: toNum(cells[10]),
        ageGroup: toNum(cells[11]),
        bonitet: normalizeBonitet(cells[12]),

        forestType: safeForestType,
        tlu,

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

/* ---------------- record builder ---------------- */
const EMPTY_FIELDS = {
    area: null, description: null,
    yarus: null, yarusHigth: null, element: null, age: null, higth: null, diam: null,
    ageKlass: null, ageGroup: null, bonitet: null,
    forestType: null, tlu: null, polnota: null,
    zapasNa1ga: null, ZapasObshiy: null, zapasPoSostavu: null,
    klassTovarnosti: null, suhostoy: null, redin: null, edinichnye: null,
    zakhl: null, zakhlLikvid: null, hozMeropriyatiya: null,
    note: null, raw: null,
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
        cells.map((x, i) => (i === 0 ? "" : x ? `${i}:${x}` : "")).filter(Boolean).join(" | ")
    );
}

/* ---------------- classification ---------------- */
function classifyRow(cells, ctx) {
    const c1 = cells[1], c2 = cells[2], c3 = cells[3];
    const leftText = normalizeBrokenWords(cells.slice(1, 7).filter(Boolean).join(" "));

    // totals headers (fallback; основные totals ловим по txt выше)
    if (/^Итого по категории/i.test(leftText)) { ctx.totalsMode = "CATEGORY"; ctx.totalsExpectSpecies = false; ctx.inSingleTrees = false; return "TOTAL_CATEGORY"; }
    if (/^Итого по кварталу/i.test(leftText)) { ctx.totalsMode = "QUARTER"; ctx.totalsExpectSpecies = false; ctx.inSingleTrees = false; return "TOTAL_QUARTER"; }
    if (/^По составляющим породам/i.test(leftText)) { ctx.totalsExpectSpecies = true; ctx.inSingleTrees = false; return "TOTAL_SPECIES_HEADER"; }

    // NOTE
    if (leftText.match(/^(подрост|подлесок|Болота:|Земли линейного протяжения:|Расчистка просек)/i)) {
        ctx.inSingleTrees = false;
        return "NOTE";
    }

    // SINGLE_TREES marker
    if (/^Единичные деревья/i.test(leftText)) { ctx.inSingleTrees = true; return "NOTE"; }

    // SINGLE_TREES line
    if (ctx.inSingleTrees) {
        const whole = cells.slice(1).filter(Boolean).join(" ").toLowerCase();
        if (whole.includes("подрост") || whole.includes("подлесок")) { ctx.inSingleTrees = false; return "NOTE"; }

        const coeff = (cells[2] || "").trim();
        const hasNumbersRight = /\d/.test(cells.slice(5).filter(Boolean).join(" "));
        if (/^\d{1,2}[А-ЯЁA-Z]{1,3}$/i.test(coeff) && hasNumbersRight) { ctx.inSingleTrees = false; return "SINGLE_TREES"; }
        ctx.inSingleTrees = false;
    }

    // MAIN / OBJECT
    if (/^\d+$/.test(c1) && /^\d+(\.\d+)?$/.test(c2)) {
        if (looksLikePolnota(cells)) return "MAIN";
        if (c3) return "OBJECT";
    }

    // SPECIES
    if (isEmptyRange(cells, 1, 5) && hasAnyRange(cells, 6, 24)) return "SPECIES";

    return "TEXT";
}

/* ---------------- main parse (with progress) ---------------- */
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
        pendingTotal: null, // { kind, titleTxt, page, rowNo }

        lastMainIndex: null,
        pendingMain: null, // { vydel, area, description, page, rowNo, composition? }
        mainExtraEls: new Set(), // элементы, добавленные в формулу текущего MAIN
        emittedLonelySpecies: new Set(),
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
            cat = canon.category(cat); // ✅ канонизация категории
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

            // --- totals by txt (не зависят от grid) ---
            {
                const totalsKind = isTotalsTitleByTxt(txt);

                if (totalsKind === "TOTAL_CATEGORY" || totalsKind === "TOTAL_QUARTER") {
                    // 1) пишем заголовок отдельной строкой
                    pushRecord(records, buildRecord({ ctx, rowNo, kind: `${totalsKind}_HEADER`, page: p, vydel: ctx.vydel }, {
                        description: txt,
                        raw: txt,
                    }));
                    // 2) ждём числовую строку
                    ctx.pendingTotal = { kind: totalsKind, titleTxt: txt, page: p, rowNo };
                    continue;
                }

                if (totalsKind === "TOTAL_SPECIES_HEADER") {
                    pushRecord(records, buildRecord({ ctx, rowNo, kind: "TOTAL_SPECIES_HEADER", page: p, vydel: ctx.vydel }, {
                        description: txt,
                        raw: txt,
                    }));
                    ctx.totalsExpectSpecies = true;
                    continue;
                }
            }

            // pendingTotal: capture numeric line after "Итого ..."
            // pendingTotal: capture numeric line after "Итого ..."
            if (ctx.pendingTotal) {
                const t = txt.trim();

                // пробуем сначала через GRID (самый правильный способ для колонок 19-23)
                let parsed = null;

                if (grid) {
                    let tc = splitLineIntoCells(line, grid);
                    tc = normalizeSpacedDecimals(tc);

                    // на всякий случай применим ваши фиксы границ 17/18 и 18/19
                    tc = normalizeZapasAndKlassTov(tc);
                    tc = normalizeKlassShiftFrom19(tc);

                    // если действительно есть цифры в правой части — считаем, что это строка итогов
                    if (hasAnyRange(tc, 15, 23)) {
                        parsed = parseTotalsValuesFromCells(tc);
                    }
                }

                // fallback: старая эвристика по тексту (когда grid не найден)
                if (!parsed && /^\d+(?:[.,]\d+)?(?:\s+\d+(?:[.,]\d+)?)+$/.test(t)) {
                    const p2 = parseTotalsNumbersFromTwoLines(ctx.pendingTotal.titleTxt, t);
                    parsed = { area: p2.area, ZapasObshiy: p2.ZapasObshiy, edinichnye: p2.edinichnye };
                }

                if (parsed) {
                    pushRecord(records, buildRecord({
                        ctx, rowNo, kind: `${ctx.pendingTotal.kind}_VALUES`, page: p, vydel: ctx.vydel
                    }, {
                        area: parsed.area ?? null,
                        ZapasObshiy: parsed.ZapasObshiy ?? null,
                        zapasPoSostavu: parsed.zapasPoSostavu ?? null,
                        klassTovarnosti: parsed.klassTovarnosti ? String(parsed.klassTovarnosti) : null,

                        suhostoy: parsed.suhostoy ?? null,
                        redin: parsed.redin ?? null,
                        edinichnye: parsed.edinichnye ?? null,
                        zakhl: parsed.zakhl ?? null,
                        zakhlLikvid: parsed.zakhlLikvid ?? null,

                        raw: `nums:${t}`,
                    }));

                    ctx.pendingTotal = null;
                    continue;
                } else {
                    ctx.pendingTotal = null;
                }
            }

            // totals species rows: parse by txt in totalsExpectSpecies mode
            if (ctx.totalsExpectSpecies) {
                const parsed = parseTotalsSpeciesRow(txt);
                if (parsed) {
                    pushRecord(records, buildRecord({ ctx, rowNo, kind: "TOTAL_SPECIES_ROW", page: p }, {
                        element: canon.species(parsed.element),
                        zapasPoSostavu: parsed.zapasPoSostavu,
                        raw: txt,
                    }));
                    continue;
                }

                // ВАЖНО: если началась обычная табличная строка (выдел/площадь), выходим из режима итогов
                if (/^\d{1,4}\s+\d+(?:[.,]\d+)?\b/.test(txt.trim())) {
                    ctx.totalsExpectSpecies = false;
                    // не continue — дальше строка обработается обычной логикой
                }
            }

            // --- for cultures: catch separate composition line after VYDEL_HEADER ---
            if (ctx.pendingMain && looksLikeStandDescription(ctx.pendingMain.description)) {
                const compLine = isPureCompositionLine(txt);
                if (compLine) {
                    ctx.pendingMain.composition = compLine;
                    continue;
                }
            }

            // --- split to cells ---
            let cells = splitLineIntoCells(line, grid);
            cells = normalizeSpacedDecimals(cells);
            cells = normalizeSpacedIntegersInNumericCols(cells);
            cells = normalizeVydelAreaSplitAcrossC1C2(cells);
            cells = normalizeAreaSplitIntoC3(cells);   // <-- ДО normalizeStartColumns
            cells = normalizeStartColumns(cells);
            cells = normalizeObjectDescriptionSplit(cells);

            cells = normalizeBonitetShiftFrom13(cells);

            cells = normalizeZapasAndKlassTov(cells);
            cells = normalizeKlassShiftFrom19(cells);
            cells = normalizeHozMeropriyatiyaShiftFrom23(cells);
            cells = normalizePolnotaAndZapasNa1gaShift(cells);
            cells = normalizeAgeHeightSplitFromC7(cells);
            

            // restore vydel-header tail from txt (e.g. "Культуры лесные")
            {
                const hdr = extractVydelAreaTail(txt);
                if (hdr) {
                    const hasMainMetrics = hasAnyRange(cells, 4, 18);
                    if (!hasMainMetrics) {
                        if (!cells[1]) cells[1] = String(hdr.vydel);
                        if (!cells[2]) cells[2] = String(hdr.area);
                        cells[3] = hdr.tail;
                    }
                }
            }

            let kind = classifyRow(cells, ctx);

            // FIX: MAIN description из txt (чтобы не терялся хвост "2Е+Е", "2Б" и т.п.)
            if (kind === "MAIN") {
                const d = extractMainDescriptionFromTxt(txt, cells);
                if (d) cells[3] = d;
            }

            // Склейка "хвостов" породной формулы (ОС / Е 140) в предыдущий MAIN
            if (ctx.lastMainIndex != null && ctx.vydel != null) {
                const prev = records[ctx.lastMainIndex];
                if (prev && prev.kind === "MAIN" && prev.vydel === ctx.vydel) {
                    const el = extractLonelyElementFromTxt(txt);

                    // важно: не трогаем нормальные строки (подрост/подлесок/итоги/новый выдел)
                    const isNotNote = !/^(подрост|подлесок|единичные деревья|итого|по составляющим породам)/i.test(txt.trim());
                    const isNotNewVydel = !/^\d{1,4}\s+\d+(?:[.,]\d+)?\b/.test(txt.trim());

                    // и чтобы это не было обычной SPECIES-строкой с метриками справа
                    const hasRightMetrics = hasAnyRange(cells, 7, 24);

                    // if (el && isNotNote && isNotNewVydel && !hasRightMetrics) {
                    //     const prevDesc = normalizeSpaces(prev.description || "");

                    //     // если элемент уже есть в формуле — просто пропускаем строку
                    //     if (formulaHasElement(prevDesc, el)) {
                    //         prev.raw = clampCellText(normalizeSpaces([prev.raw, `| skip-el:${el}`].join(" ")));
                    //         continue;
                    //     }

                    //     prev.description = appendExtraElementToMainFormula(prevDesc, el, ctx);
                    //     prev.raw = clampCellText(normalizeSpaces([prev.raw, `| +el:${el}`].join(" ")));
                    //     continue;
                    // }

                    if (el && isNotNote && isNotNewVydel && !hasRightMetrics) {
                        const prevDesc = normalizeSpaces(prev.description || "");
                        const key = `${ctx.vydel}:${el}`;

                        // если элемент уже есть в формуле — ничего не делаем
                        if (formulaHasElement(prevDesc, el)) {
                            prev.raw = clampCellText(normalizeSpaces([prev.raw, `| skip-el:${el}`].join(" ")));
                            continue;
                        }

                        // добавляем элемент в формулу
                        prev.description = appendExtraElementToMainFormula(prevDesc, el, ctx);
                        prev.raw = clampCellText(normalizeSpaces([prev.raw, `| +el:${el}`].join(" ")));

                        // ✅ element-only пишем ТОЛЬКО если в исходной строке есть цифры
                        // (иначе строки "ОС" / "ОЛСА" не будут плодить мусор)
                        const hasDigitsInTxt = /\d/.test(txt);
                        if (hasDigitsInTxt && !ctx.emittedLonelySpecies.has(key)) {
                            pushRecord(
                                records,
                                buildRecord({ ctx, rowNo, kind: "SPECIES", page: p, vydel: ctx.vydel }, {
                                    element: el,
                                    note: "element-only",
                                    raw: `lonely:${txt}`,
                                })
                            );
                            ctx.emittedLonelySpecies.add(key);
                        }

                        continue;
                    }
                    
                }
            }

            // reset pendingMain on new vydel
            if (ctx.pendingMain && /^\d+$/.test(cells[1] || "") && toNum(cells[1]) !== ctx.pendingMain.vydel) ctx.pendingMain = null;

            // OBJECT -> pendingMain + VYDEL_HEADER
            if (kind === "OBJECT") {
                const desc = (cells[3] || "").trim();
                const hasMainMetrics = hasAnyRange(cells, 4, 18);
                if (desc && !looksLikeRealObjectName(desc) && !hasMainMetrics) {
                    const vyd = toNum(cells[1]);
                    const ar = toNum(cells[2]);

                    const hdrDesc = getVydelHeaderDescriptionFromDict(cells, txt) ?? "";

                    const hm = cells[24] ? canon.hozMerop(cells[24]) : null;

                    // пишем VYDEL_HEADER со строгим description (справочник или пусто)
                    pushRecord(records, buildRecord({ ctx, rowNo, kind: "VYDEL_HEADER", page: p, vydel: vyd }, {
                        area: ar,
                        description: hdrDesc,
                        // чтобы не терять исходный текст строки — можно писать в note (по желанию)
                        hozMeropriyatiya: hm,
                        note: hdrDesc ? null : normalizeSpaces(txt),
                        raw: cellsRaw(cells),
                    }));

                    // ВАЖНО: pendingMain.description должен быть тем же, иначе дальше (looksLikeStandDescription) будет хуже работать
                    ctx.pendingMain = { vydel: vyd, area: ar, description: hdrDesc, page: p, rowNo, composition: null };

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
                        // для "Насажд/Культ/Полог" description должна быть ТОЛЬКО породная формула
                        if (ctx.pendingMain.composition) {
                            cells[3] = ctx.pendingMain.composition;
                        } else {
                            const compFromTxt = extractCompositionFromTxtAnywhere(txt);
                            if (compFromTxt) cells[3] = compFromTxt;
                            else {
                                const comp = extractCompositionToken(d1);
                                cells[3] = comp ? comp : d1;
                            }
                        }
                    } else {
                        cells[3] = normalizeSpaces(`${d0} ${d1}`.trim());
                    }

                    ctx.pendingMain = null;
                    kind = "MAIN";
                }
            }

            // Additional yarus MAIN without N/area
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

            // MAIN normalization
            if (kind === "MAIN") {
                // НЕ режем description: она может собираться из нескольких строк
                // (максимум можно убрать лишние пробелы)
                if (cells[3]) cells[3] = normalizeSpaces(cells[3]);

                cells = normalizeMainCells(cells);
                cells = normalizeZapasAndKlassTov(cells);
                cells = normalizeMainByRanges(cells);

                // ✅ восстановление age/height/diam/ageKlass/ageGroup/bonitet из txt
                cells = rescueMainMetricsFromTxt(cells, txt);
                cells = normalizeMainByRanges(cells);

                cells = normalizeMainZapasFromText(cells, txt);

                

                cells = normalizeKlassShiftFrom19(cells);
                cells = normalizeHozMeropriyatiyaShiftFrom23(cells);

                cells = normalizePolnotaAndZapasNa1gaShift(cells);  // ✅
                cells = normalizeAgeHeightSplitFromC7(cells);        // ✅
            }

            if (kind === "SINGLE_TREES") cells = normalizeSingleTreesCells(cells);

            // MAIN/OBJECT write
            if (kind === "MAIN" || kind === "OBJECT") {
                ctx.vydel = toNum(cells[1]);
                ctx.inSingleTrees = false;

                // RULE: для OBJECT description только из справочника или пусто
                let objectNote = null;
                if (kind === "OBJECT") {
                    const dictDesc = getObjectDescriptionFromDict(cells, txt);
                    if (dictDesc) {
                        cells[3] = dictDesc;
                        objectNote = null; // ✅ не дублируем исходную строку для распознанных OBJECT
                    } else {
                        cells[3] = "";
                        objectNote = normalizeSpaces(txt) || null; // для неизвестных OBJECT оставляем исходник
                    }
                }

                if (cells[24]) cells[24] = canon.hozMerop(cells[24]);

                const named = mapCellsToNamed(cells);
                pushRecord(records, buildRecord({ ctx, rowNo, kind, page: p, vydel: named.vydel }, {
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
                    tlu: named.tlu, // ✅
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
                    note: kind === "OBJECT" ? objectNote : null,
                }));

                ctx.pendingMain = null;
                ctx.mainExtraEls = new Set();
                ctx.emittedLonelySpecies = new Set();
                ctx.lastMainIndex = records.length - 1;
                continue;
            }

            // below only inside vydel
            if (ctx.vydel == null) continue;

            // TEXT dual: "+..." and species on right
            if (kind === "TEXT" && ctx.lastMainIndex != null) {
                const c3 = (cells[3] || "").trim();
                const hasSpeciesRight = hasAnyRange(cells, 6, 24);

                if (/^[+,]/.test(c3) && hasSpeciesRight) {
                    const prev = records[ctx.lastMainIndex];
                    if (prev && prev.kind === "MAIN" && prev.vydel === ctx.vydel) {
                        prev.description = normalizeSpaces([prev.description, c3].filter(Boolean).join(" "));
                        prev.raw = clampCellText(normalizeSpaces([prev.raw, `| cont:${c3}`].join(" ")));
                    }

                    let sc = cells.slice();
                    for (let i = 1; i <= 5; i++) sc[i] = "";
                    sc[3] = "";
                    sc = normalizeZapasAndKlassTov(sc);
                    sc = normalizeKlassShiftFrom19(sc);
                    sc = normalizeHozMeropriyatiyaShiftFrom23(sc);
                    sc = normalizeSpeciesCells(sc);

                    if (sc[24]) sc[24] = canon.hozMerop(sc[24]);

                    const named = mapCellsToNamed(sc);

                    const moved = moveTluToLastMain(ctx, records, named.tlu);
                    if (moved) named.tlu = null;

                    pushRecord(records, buildRecord({ ctx, rowNo, kind: "SPECIES", page: p, vydel: ctx.vydel }, {
                        element: named.element,
                        age: named.age,
                        higth: named.higth,
                        diam: named.diam,
                        forestType: named.forestType,
                        tlu: named.tlu, // ✅
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
                        raw: cellsRaw(sc),
                    }));

                    continue;
                }
            }

            if (kind === "TEXT" && ctx.vydel != null && !ctx.pendingMain) {
                const pm = parseShortMainNoIdFromTxt(txt);
                if (pm) {
                    // делаем MAIN без vydel/area (добавочный ярус)
                    cells = Array.from({ length: 25 }, () => "");
                    cells[1] = String(ctx.vydel);
                    cells[3] = pm.comp;
                    cells[4] = "1";                  // по умолчанию ярус 1 (если у вас иначе — скорректируем)
                    cells[5] = String(pm.yarusH);
                    cells[6] = pm.el;
                    cells[7] = String(pm.age);
                    cells[8] = String(pm.h);
                    cells[9] = String(pm.d);
                    cells[10] = String(pm.ageKlass);
                    cells[15] = String(pm.zapasNa1ga);

                    kind = "MAIN";
                }
            }

            // NOTE/TEXT
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

                // IMPORTANT: note берём из txt (правильный порядок слов)
                pushRecord(records, buildRecord({ ctx, rowNo, kind, page: p, vydel: ctx.vydel }, {
                    note: txt,
                    raw: cellsRaw(cells),
                }));
                continue;
            }

            // SPECIES / SINGLE_TREES
            if (kind === "SPECIES" || kind === "SINGLE_TREES") {
                if (kind === "SPECIES") { 
                    cells = normalizeZapasAndKlassTov(cells);
                    cells = normalizeKlassShiftFrom19(cells);
                    cells = normalizeHozMeropriyatiyaShiftFrom23(cells);
                    cells = normalizeSpeciesCells(cells);

                    cells = normalizeSpeciesAgeHeightSplitFromC8(cells);

                    cells = normalizeSpacedIntegersInNumericCols(cells); // ✅ если не вызываете глобально
                    cells = normalizeSpeciesAgeHeightShift(cells);  

                    cells = normalizePolnotaAndZapasNa1gaShift(cells);  // ✅
                    cells = normalizeAgeHeightSplitFromC7(cells);        // ✅
                }
                if (cells[24]) cells[24] = canon.hozMerop(cells[24]);
                
                const named = mapCellsToNamed(cells);

                if (kind === "SPECIES") {
                    const moved = moveTluToLastMain(ctx, records, named.tlu);
                    if (moved) named.tlu = null; // чистим в SPECIES
                }

                let description = null;
                let note = null;
                if (kind === "SINGLE_TREES") {
                    const coeff = (cells[2] || "").trim();
                    description = coeff || null;
                    note = null; // ✅ не дублируем "Единичные деревья"
                }

                pushRecord(records, buildRecord({ ctx, rowNo, kind, page: p, vydel: ctx.vydel }, {
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
                    tlu: named.tlu, // ✅
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
                }));
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
        { header: "kind", key: "kind", width: 18 },
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
        { header: "tlu", key: "tlu", width: 8 },
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