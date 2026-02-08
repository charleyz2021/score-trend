import * as XLSX from "xlsx";
import { parseSheetToBlocks, type ParsedBlock, type ParsedColumn } from "./ExcelParse";

export type Row = Record<string, any>;
export type ExamId = string;

export type ExamScope = "auto" | "school" | "class";

export type MetricId =
    | "total"
    | "chinese"
    | "math"
    | "english"
    | "physics"
    | "chemistry"
    | "biology"
    | "history"
    | "geography"
    | "politics"
    | "classRank"
    | "schoolRank";

export type MetricCols = Partial<Record<MetricId, string>>;

export type ExamConfig = {
    id: ExamId;
    fileName: string;
    sheetName: string;
    examName: string;

    blocks: ParsedBlock[];
    blockId: string;

    cols: ParsedColumn[];
    rows: any[];

    idCol: string;        // 可空
    nameCol?: string;     // 必须能识别到，否则 fatal
    classCol?: string;    // 可空
    totalCol?: string;    // 可空：允许导入但无法画总分

    metricCols?: MetricCols;

    scope: ExamScope;
    inferredClass: string;
    overrideClass: string;

    fatalErrors: string[];
    warnings: string[];
};

// ========== utils ==========
export function normalizeHeader(s: any) {
    return String(s ?? "")
        .trim()
        .replace(/（[a-zA-Z]+列）$/, "")
        .replace(/\([a-zA-Z]+列\)$/, "")
        .replace(/\s+/g, "")
        .replace(/（/g, "(")
        .replace(/）/g, ")")
        .toLowerCase();
}

export function toNumber(v: any): number | null {
    if (v === null || v === undefined) return null;
    if (typeof v === "number" && Number.isFinite(v)) return v;
    const s = String(v).trim();
    if (!s) return null;
    const cleaned = s.replace(/[^\d.\-]/g, "");
    if (!cleaned) return null;
    const n = Number(cleaned);
    return Number.isFinite(n) ? n : null;
}

export function isLikelyChineseName(v: any) {
    const s = String(v ?? "").trim();
    return /^[\u4e00-\u9fa5]{2,5}$/.test(s);
}

export function makeExamId(fileName: string, sheetName: string) {
    return `${fileName}::${sheetName}`;
}

// ========== column guessing ==========
function findAllNameCols(cols: ParsedColumn[]) {
    const hits: string[] = [];
    for (const c of cols) {
        const h = normalizeHeader(c.label);
        if (h.includes("姓名") || h.includes("名字")) hits.push(c.key);
    }
    return hits;
}

function guessIdCol(cols: ParsedColumn[]) {
    const keys = ["学生id", "学号", "studentid", "student_id", "id"];
    for (const c of cols) {
        const h = normalizeHeader(c.label);
        if (keys.some(k => h.includes(k))) return c.key;
    }
    return "";
}

function guessClassCol(cols: ParsedColumn[]) {
    const includeExact = ["班级", "班级号", "行政班", "班号", "班别", "班级编号"];
    const exclude = ["排名", "名次", "rank", "班级排", "年级排", "班排", "年排", "校排", "级排"];

    const isRankLike = (h: string) =>
        exclude.some(k => h.includes(k)) ||
        (h.includes("排") && (h.includes("班") || h.includes("年") || h.includes("校") || h.includes("级")));

    for (const c of cols) {
        const h = normalizeHeader(c.label);
        if (isRankLike(h)) continue;
        if (includeExact.includes(h)) return c.key;
    }
    for (const c of cols) {
        const h = normalizeHeader(c.label);
        if (isRankLike(h)) continue;
        if (h.includes("班级号") || h.includes("行政班") || h.includes("班号") || h === "班级" || h.includes("班级")) {
            return c.key;
        }
    }
    return "";
}

function guessNameCol(rows: Row[], cols: ParsedColumn[]) {
    const prefer = cols.find(c => {
        const h = normalizeHeader(c.label);
        return h.includes("姓名") || h.includes("名字");
    });

    let best = prefer?.key ?? cols[0]?.key ?? "";
    let bestScore = -1;

    for (const c of cols) {
        const h = normalizeHeader(c.label);
        let score = 0;
        if (h.includes("姓名") || h.includes("名字")) score += 6;

        let seen = 0, hit = 0;
        for (const r of rows.slice(0, 450)) {
            const s = String(r[c.key] ?? "").trim();
            if (!s) continue;
            seen++;
            if (isLikelyChineseName(s)) hit++;
        }
        if (seen) score += (hit / seen) * 10;
        if (score > bestScore) { bestScore = score; best = c.key; }
    }

    return best;
}

function guessTotalCol(rows: Row[], cols: ParsedColumn[]) {
    function scoreTotal(c: ParsedColumn) {
        const h = normalizeHeader(c.label);
        let score = 0;

        const good = [
            "总分",
            "总分得分", "总分_得分",
            "总分原始分", "总分_原始分",
            "总分原始", "总分_原始",
        ];
        if (good.some(g => h.includes(g))) score += 20;

        const isRankLike =
            h.includes("排名") ||
            h.includes("名次") ||
            h.includes("班级排") ||
            h.includes("年级排") ||
            (h.includes("排") && (h.includes("班") || h.includes("年") || h.includes("校") || h.includes("级")));
        if (isRankLike) return -9999;

        const isMetaLike =
            h === "班级" || h.includes("班级号") || h.includes("行政班") || h.includes("班号") ||
            h.includes("考号") || h.includes("准考证") ||
            h === "姓名" || h.includes("名字");
        if (isMetaLike) return -9999;

        if (h.includes("赋分")) score -= 50;

        let seen = 0;
        let numeric = 0;
        let inRange = 0;
        let veryLarge = 0;

        for (const r of rows.slice(0, 450)) {
            const v = r[c.key];
            if (v === null || v === undefined || String(v).trim() === "") continue;
            seen++;

            const n = toNumber(v);
            if (n !== null) {
                numeric++;
                if (n >= 0 && n <= 1200) inRange++;
                if (n >= 100000) veryLarge++;
            }
        }

        if (seen > 0) {
            score += (numeric / seen) * 5;
            score += (inRange / seen) * 30;
            score -= (veryLarge / seen) * 200;
        } else {
            score -= 50;
        }

        return score;
    }

    let best = "";
    let bestScore = -1e9;

    for (const c of cols) {
        const s = scoreTotal(c);
        if (s > bestScore) { bestScore = s; best = c.key; }
    }

    if (bestScore < 10) return "";
    return best;
}

function hasAnyNonEmpty(rows: any[], colKey: string | undefined, limit = 400) {
  if (!colKey) return false;
  for (const r of rows.slice(0, limit)) {
    const v = String(r[colKey] ?? "").trim();
    if (v) return true;
  }
  return false;
}

// ========== metric detection（排名强排除） ==========
// 鲁棒 + 班排/年排强排除
function detectMetricCols(cols: { key: string; label: string }[]) {
    const metricCols: any = {};

    const norm = (s: any) =>
        String(s ?? "")
            .trim()
            .toLowerCase()
            .replace(/\s+/g, "")
            .replace(/[（）()【】\[\]_\-—·.。:：/\\]/g, "");

    // 是否“包含某关键词”（用于：语文得分 / 语文_得分 / 语文得分_8 等）
    const has = (h: string, kw: string) => h.includes(norm(kw));

    for (const c of cols) {
        const h = norm(c.label);

        // ========= 排名：强排除 =========
        // 年级/学校排名：必须含 年级/学校/校/级 + 排名/名次/排；且强排除 班/班级
        const isSchoolRank =
            (has(h, "年级") || has(h, "学校") || has(h, "校") || has(h, "级")) &&
            (has(h, "排名") || has(h, "名次") || has(h, "排")) &&
            !has(h, "班");

        // 班级排名：必须含 班/班级 + 排名/名次/排；且强排除 年级/学校/校/级
        const isClassRank =
            has(h, "班") &&
            (has(h, "排名") || has(h, "名次") || has(h, "排")) &&
            !has(h, "年级") &&
            !has(h, "学校") &&
            !has(h, "校");

        if (!metricCols.schoolRank && isSchoolRank) {
            metricCols.schoolRank = c.key;
            continue;
        }
        if (!metricCols.classRank && isClassRank) {
            metricCols.classRank = c.key;
            continue;
        }

        // ========= 分数：允许 “xx得分/xx_得分/xx得分_8” =========
        // 总分：允许 总分/总分得分/总分原始…
        if (!metricCols.total && (has(h, "总分") || has(h, "总分得分") || has(h, "总分原始"))) {
            metricCols.total = c.key;
            continue;
        }

        // 科目：按模板截图列
        if (!metricCols.chinese && has(h, "语文")) { metricCols.chinese = c.key; continue; }
        if (!metricCols.math && has(h, "数学")) { metricCols.math = c.key; continue; }
        if (!metricCols.english && has(h, "英语")) { metricCols.english = c.key; continue; }
        if (!metricCols.physics && has(h, "物理")) { metricCols.physics = c.key; continue; }
        if (!metricCols.chemistry && has(h, "化学")) { metricCols.chemistry = c.key; continue; }
        if (!metricCols.biology && has(h, "生物")) { metricCols.biology = c.key; continue; }
        if (!metricCols.politics && has(h, "政治")) { metricCols.politics = c.key; continue; }
        if (!metricCols.history && has(h, "历史")) { metricCols.history = c.key; continue; }
        if (!metricCols.geography && has(h, "地理")) { metricCols.geography = c.key; continue; }
    }

    return metricCols;
}

// ========== block quality ==========
function scoreBlockQuality(block: ParsedBlock) {
    const cols = block.columns;
    const rows = block.rows;
    if (!cols.length || !rows.length) return -999;

    const headers = cols.map(c => normalizeHeader(c.label)).join("|");
    let score = 0;
    if (headers.includes("姓名") || headers.includes("名字")) score += 15;
    if (headers.includes("总分")) score += 15;

    const nameCol = guessNameCol(rows, cols);
    const totalCol = guessTotalCol(rows, cols);

    let seen = 0, nameHit = 0, totalHit = 0;
    for (const r of rows.slice(0, 500)) {
        const n = String(r[nameCol] ?? "").trim();
        const t = totalCol ? toNumber(r[totalCol]) : null;
        if (!n && t === null) continue;
        seen++;
        if (isLikelyChineseName(n)) nameHit++;
        if (t !== null) totalHit++;
    }
    if (seen) {
        score += (nameHit / seen) * 20;
        score += (totalHit / seen) * 15;
    }
    if (cols.length >= 6) score += 3;
    if (cols.length < 4) score -= 10;

    return score;
}

// ========== class inference ==========
export type ClassInferenceResult = {
    classSet: string[];
    nameToClasses: Map<string, Set<string>>;
};

function buildClassIndex(exams: ExamConfig[]): ClassInferenceResult {
    const nameToClasses = new Map<string, Set<string>>();
    const classSet = new Set<string>();

    for (const ex of exams) {
        if (!hasClassData(ex) || !ex.classCol || !ex.nameCol) continue;
        for (const r of ex.rows) {
            const name = String(r[ex.nameCol] ?? "").trim();
            const cls = String(r[ex.classCol] ?? "").trim();
            if (!name || !cls) continue;
            classSet.add(cls);
            if (!nameToClasses.has(name)) nameToClasses.set(name, new Set());
            nameToClasses.get(name)!.add(cls);
        }
    }

    return { classSet: Array.from(classSet).sort(), nameToClasses };
}

function inferSheetClass(ex: ExamConfig, idx: ClassInferenceResult) {
    if (hasClassData(ex)) return ex.inferredClass || "";

    const N = 6;
    const names: string[] = [];
    if (!ex.nameCol) return "未知班级";
    for (const r of ex.rows) {
        const n = String(r[ex.nameCol] ?? "").trim();
        if (!n) continue;
        names.push(n);
        if (names.length >= N) break;
    }

    if (names.length < 3) return "未知班级";

    const vote = new Map<string, number>();
    let hit = 0;

    for (const n of names) {
        const classes = idx.nameToClasses.get(n);
        if (!classes || classes.size === 0) continue;
        hit++;
        for (const c of classes) vote.set(c, (vote.get(c) ?? 0) + 1);
    }

    if (hit < 3 || vote.size === 0) return "未知班级";

    let bestC = "";
    let bestV = -1;
    for (const [c, v] of vote.entries()) {
        if (v > bestV) { bestV = v; bestC = c; }
    }

    const ratio = bestV / hit;
    if (ex.rows.length >=150 && ratio < 0.85) return "全校";
    if (ratio >= 0.7) return bestC;
    return "未知班级";
}

// ========== import main ==========
export async function importFiles(files: File[]): Promise<ExamConfig[]> {
    const all: ExamConfig[] = [];

    for (const f of files) {
        const ab = await f.arrayBuffer();
        const wb = XLSX.read(ab, { type: "array" });

        for (const sheetName of wb.SheetNames) {
            if (sheetName.trim() === "说明") continue; // 跳过说明页
            const blocks = parseSheetToBlocks(wb, sheetName);
            if (!blocks.length) continue;

            // pick best block
            let best = blocks[0];
            let bestScore = -Infinity;
            for (const b of blocks) {
                const s = scoreBlockQuality(b);
                if (s > bestScore) { bestScore = s; best = b; }
            }

            const cols = best.columns;
            const rows = best.rows;

            const id = makeExamId(f.name, sheetName);

            const idCol = guessIdCol(cols);
            let classCol = guessClassCol(cols);
            if (classCol && !hasAnyNonEmpty(rows, classCol)) classCol = "";

            const nameCols = findAllNameCols(cols);
            const fatalErrors: string[] = [];
            const warnings: string[] = [];

            if (nameCols.length >= 2) {
                fatalErrors.push(`检测到多个“姓名”列（${nameCols.length}列）。请合并为一列后再导入。`);
            }

            const nameCol = nameCols[0] || guessNameCol(rows, cols);
            if (!nameCol) fatalErrors.push(`未识别到“姓名”列（支持“姓名/姓 名/名字”）。`);

            let fixedClassCol = classCol;
            if (fixedClassCol) {
                const N = Math.min(300, rows.length);
                let seen = 0;
                let nonEmpty = 0;

                for (const r of rows.slice(0, N)) {
                    // 只统计有姓名的行，避免空行干扰
                    const nameV = nameCol ? String(r[nameCol] ?? "").trim() : "";
                    if (!nameV) continue;

                    seen++;
                    const v = String(r[fixedClassCol] ?? "").trim();
                    if (v) nonEmpty++;
                }

                // seen 太少不判断；否则非空比例过低 => 认为“班级列无效”
                const ratio = seen ? nonEmpty / seen : 0;

                // 阈值可以调：0.1/0.2 都行。0.1 更保守，避免误判
                if (seen >= 10 && ratio < 0.1) {
                    // warnings.push(`检测到“班级”列但几乎全为空（非空 ${(ratio * 100).toFixed(0)}%），已按“无班级列”处理，将使用自动推断/班级覆盖。`);
                    fixedClassCol = ""; // 或 undefined
                }
            }

            const totalCol = guessTotalCol(rows, cols);

            const metricCols0 = detectMetricCols(cols);
            const metricCols: MetricCols = { ...metricCols0, total: metricCols0.total || totalCol };

            // eslint-disable-next-line no-console
            console.log(`[metricCols] ${sheetName}`, metricCols);

            if (totalCol) {
                let seen = 0, tiny = 0;
                for (const r of rows.slice(0, 300)) {
                    const n = toNumber(r[totalCol]);
                    if (n === null) continue;
                    seen++;
                    if (n >= 0 && n <= 20) tiny++;
                }
                if (seen >= 10 && tiny / seen >= 0.7) {
                    warnings.push(`总分列疑似异常（大量值<=20），可能误选为排名列，请检查该sheet表头。`);
                }
            }

            all.push({
                id,
                fileName: f.name,
                sheetName,
                examName: sheetName,

                blocks,
                blockId: best.blockId,
                cols,
                rows,

                idCol,
                nameCol,
                classCol,
                totalCol,

                metricCols,

                scope: "auto",
                inferredClass: "",
                overrideClass: "",

                fatalErrors,
                warnings,
            });
        }
    }

    // 班级推断
    const idx = buildClassIndex(all);
    for (const ex of all) ex.inferredClass = inferSheetClass(ex, idx);

    return all;
}

export function getEffectiveClass(ex: ExamConfig) {
    const c = String(ex.overrideClass ?? "").trim();
    if (c) return c;

    if (ex.classCol) {
        const set = new Set<string>();
        const N = 300;
        for (const r of ex.rows.slice(0, N)) {
            const raw = r[ex.classCol];
            if (raw === null || raw === undefined) continue;
            const v = String(raw).trim();
            if (!v) continue;
            set.add(v);
            if (set.size >= 2) break;
        }

        if (set.size === 1) return Array.from(set)[0];
        if (set.size >= 2) return "全校";
        return "未知班级";
    }

    return ex.inferredClass || "未知班级";
}

export function hasClassData(ex: ExamConfig) {
  const col = ex.classCol;
  if (!col) return false;

  const N = 300;
  for (const r of ex.rows.slice(0, N)) {
    const v = String(r[col] ?? "").trim();
    if (v) return true;
  }
  return false;
}