import * as XLSX from "xlsx";

export type ParsedColumn = {
  key: string;      // 唯一key（用于数据）
  label: string;    // 展示给用户看的：姓名（P列）
  colIndex: number; // 0-based
};

export type ParsedBlock = {
  blockId: string;      // e.g. "block-1"
  rangeLabel: string;   // e.g. "块1 (A~H列)"
  columns: ParsedColumn[];
  rows: Record<string, any>[];
  meta: {
    startCol: number;
    endCol: number;
    headerStartRow: number;
    headerRowsUsed: number;
  };
};

export type SheetParseResult = {
  sheetName: string;
  blocks: ParsedBlock[];
};

function getCellValueMerged(sheet: XLSX.WorkSheet, r0: number, c0: number) {
  // r0/c0：0-based
  const addr = XLSX.utils.encode_cell({ r: r0, c: c0 });
  const cell = sheet[addr];
  if (cell && cell.v !== undefined && cell.v !== null && String(cell.v).trim() !== "") return cell.v;

  const merges = (sheet["!merges"] ?? []) as XLSX.Range[];
  for (const m of merges) {
    if (r0 >= m.s.r && r0 <= m.e.r && c0 >= m.s.c && c0 <= m.e.c) {
      const topLeft = XLSX.utils.encode_cell({ r: m.s.r, c: m.s.c });
      const tl = sheet[topLeft];
      return tl ? tl.v : null;
    }
  }
  return null;
}

function applyMergesToMatrix(ws: XLSX.WorkSheet, matrix: any[][], maxRows = 30) {
  const merges = (ws["!merges"] ?? []) as XLSX.Range[];
  if (!merges.length) return matrix;

  // 确保 matrix 至少有 maxRows 行
  while (matrix.length < maxRows) matrix.push([]);

  for (const m of merges) {
    // 只回填前 maxRows 行（足够用于识别表头）
    if (m.s.r >= maxRows) continue;

    const tl = getCellValueMerged(ws, m.s.r, m.s.c);
    if (tl === null || tl === undefined || String(tl).trim() === "") continue;

    const rEnd = Math.min(m.e.r, maxRows - 1);
    for (let r = m.s.r; r <= rEnd; r++) {
      const row = matrix[r] ?? (matrix[r] = []);
      for (let c = m.s.c; c <= m.e.c; c++) {
        // 只填空的，避免覆盖真实值
        if (row[c] === undefined || row[c] === null || String(row[c]).trim() === "") {
          row[c] = tl;
        }
      }
    }
  }

  return matrix;
}

function colLetter(n: number) {
  let s = "";
  let x = n + 1;
  while (x > 0) {
    const r = (x - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    x = Math.floor((x - 1) / 26);
  }
  return s;
}

function norm(s: any) {
  return String(s ?? "")
    .trim()
    .replace(/\s+/g, "")
    .replace(/（/g, "(")
    .replace(/）/g, ")")
    .toLowerCase();
}

function isEmptyCell(v: any) {
  return v === null || v === undefined || String(v).trim() === "";
}

function fillForward(arr: any[]) {
  const out = [...arr];
  let last = "";
  for (let i = 0; i < out.length; i++) {
    const v = String(out[i] ?? "").trim();
    if (v) last = v;
    else out[i] = last;
  }
  return out;
}

function isSeparatorCol(matrix: any[][], col: number, topRows = 3) {
  for (let r = 0; r < Math.min(topRows, matrix.length); r++) {
    if (!isEmptyCell(matrix[r]?.[col])) return false;
  }
  return true;
}

function detectHeaderStartRow(matrix: any[][]) {
  const LIMIT = Math.min(18, matrix.length);

  const hasKeyword = (row: any[]) => {
    const s = row.map((v) => String(v ?? "")).join("|");
    return /姓名|名字|学号|考号|学籍|班级|总分|合计|总计|分数|排名|名次|扣分|原始分|赋分|语文|数学|英语|物理|化学|生物|政治|历史|地理/i.test(
      s
    );
  };

  const nonEmptyCount = (row: any[]) =>
    row.reduce((acc, v) => acc + (String(v ?? "").trim() ? 1 : 0), 0);

  let bestRow = 0;
  let bestScore = -1;

  for (let r = 0; r < LIMIT; r++) {
    const row = matrix[r] ?? [];
    const ne = nonEmptyCount(row);

    // 标题行：只有 1 个较长字符串
    const nonEmptyCells = row.filter((v) => String(v ?? "").trim());
    const looksLikeTitle =
      nonEmptyCells.length === 1 && String(nonEmptyCells[0]).trim().length >= 6;
    if (looksLikeTitle) continue;

    let score = 0;
    score += ne;
    if (hasKeyword(row)) score += 12;

    if (ne < 3 && !hasKeyword(row)) continue;

    if (score > bestScore) {
      bestScore = score;
      bestRow = r;
    }
  }

  return bestRow;
}

/**
 * 从二维表里推断表头（支持1~2行表头）
 */
function buildHeaders(view: any[][]) {
  const row0 = view[0] ?? [];
  const row1 = view[1] ?? [];
  const W = Math.max(row0.length, row1.length);

  const r0 = fillForward([...Array(W)].map((_, i) => row0[i] ?? ""));
  const r1 = [...Array(W)].map((_, i) => row1[i] ?? ""); // ✅ row1 不 forward

  const HEADER_KEYWORDS =
    /姓名|名字|学号|考号|学籍|班级|总分|合计|总计|分数|排名|名次|扣分|原始分|赋分|语文|数学|英语|物理|化学|生物|政治|历史|地理/i;

  const nonEmptyCells = (arr: any[]) =>
    arr.map((v) => String(v ?? "").trim()).filter((s) => s.length > 0);

  const looksLikeChineseName = (s: string) => /^[\u4e00-\u9fa5]{2,5}$/.test(s);
  const looksLikeNumber = (s: string) => /^-?\d+(\.\d+)?$/.test(s);

  const c1 = nonEmptyCells(r1);
  const kwHits = c1.filter((s) => HEADER_KEYWORDS.test(s)).length;
  const nameHits = c1.filter((s) => looksLikeChineseName(s)).length;
  const numHits = c1.filter((s) => looksLikeNumber(s)).length;

  const ratio = (x: number) => (c1.length ? x / c1.length : 0);

  const row1LooksLikeData =
    kwHits === 0 && (ratio(nameHits) >= 0.25 || ratio(numHits) >= 0.6);

  const row1HasStrongHeader = c1.some((s) =>
    /姓名|原始分|赋分|校内排名|班级排名|年级排名|总分|合计|名次|排名/i.test(s)
  );
  const row1LooksLikeHeader = kwHits >= 1 || row1HasStrongHeader;

  const nonEmptyCount0 = nonEmptyCells(r0).length;
  const nonEmptyCount1 = nonEmptyCells(r1).length;

  const row0Sparse = nonEmptyCount0 <= Math.max(2, Math.floor(W * 0.25));
  const row1Rich = nonEmptyCount1 >= Math.max(4, Math.floor(W * 0.35));

  const row0HasGroupKeyword = nonEmptyCells(r0).some((s) => HEADER_KEYWORDS.test(s));
  const groupHeaderPattern = (row0Sparse && row1Rich) || (row0HasGroupKeyword && row1Rich);

  const useTwoRows = !row1LooksLikeData && (row1LooksLikeHeader || groupHeaderPattern);

  const headers: string[] = [];
  for (let c = 0; c < W; c++) {
    const a = String(r0[c] ?? "").trim();
    const b = String(r1[c] ?? "").trim();

    let h = "";
    if (useTwoRows) {
      if (a && b && a !== b) h = `${a}_${b}`;
      else h = a || b;
    } else {
      h = a;
    }

    if (!h) h = `未命名`;
    headers.push(h);
  }

  return { headers, headerRowsUsed: useTwoRows ? 2 : 1 };
}

/**
 * 按“空列”把一个sheet切成多个块（解决左右并排表）
 */
function splitBlocks(view: any[][], headers: string[]) {
  const W = headers.length;
  const cuts: number[] = [];
  for (let c = 0; c < W; c++) {
    if (isSeparatorCol(view, c, 2)) cuts.push(c);
  }

  const ranges: Array<{ start: number; end: number }> = [];
  let start = 0;
  let c = 0;
  while (c < W) {
    if (cuts.includes(c)) {
      if (c - 1 >= start) ranges.push({ start, end: c - 1 });
      while (c < W && cuts.includes(c)) c++;
      start = c;
    } else {
      c++;
    }
  }
  if (start < W) ranges.push({ start, end: W - 1 });

  return ranges.filter((r) => r.end - r.start + 1 >= 3);
}

function makeUniqueColumns(headers: string[], start: number, end: number) {
  const count = new Map<string, number>();
  const cols: ParsedColumn[] = [];
  for (let c = start; c <= end; c++) {
    const base = headers[c];
    const keyBase = norm(base) || "col";
    const n = (count.get(keyBase) ?? 0) + 1;
    count.set(keyBase, n);

    const key = `${keyBase}__${c}`;
    const letter = colLetter(c);
    const label = `${base}（${letter}列）`;
    cols.push({ key, label, colIndex: c });
  }
  return cols;
}

export function parseSheetToBlocks(wb: XLSX.WorkBook, sheetName: string): ParsedBlock[] {
  const ws = wb.Sheets[sheetName];
  let matrix = XLSX.utils.sheet_to_json<any[]>(ws, { header: 1, defval: "" }) as any[][];
  if (!matrix.length) return [];

  // 先把合并单元格的值回填到 matrix（至少前30行足够识别表头）
  matrix = applyMergesToMatrix(ws, matrix, 30);

  const headerStartRow = detectHeaderStartRow(matrix);

  const view = matrix.slice(headerStartRow);

  const { headers, headerRowsUsed } = buildHeaders(view);
  const ranges = splitBlocks(view, headers);
  const dataStartRow = headerRowsUsed;

  return ranges.map((r, idx) => {
    const columns = makeUniqueColumns(headers, r.start, r.end);

    const rows: Record<string, any>[] = [];
    for (let i = dataStartRow; i < view.length; i++) {
      const row = view[i] ?? [];
      let anyVal = false;
      const obj: Record<string, any> = {};

      for (const col of columns) {
        const v = row[col.colIndex];
        if (!isEmptyCell(v)) anyVal = true;
        obj[col.key] = v;
      }
      if (anyVal) rows.push(obj);
    }

    const rangeLabel = `块${idx + 1} (${colLetter(r.start)}~${colLetter(r.end)}列)`;
    return {
      blockId: `block-${idx + 1}`,
      rangeLabel,
      columns,
      rows,
      meta: {
        startCol: r.start,
        endCol: r.end,
        headerStartRow,
        headerRowsUsed,
      },
    };
  });
}
