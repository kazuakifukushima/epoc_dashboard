/**
 * クライアント側 Excel パーサー
 * ブラウザで Excel を読み込み、ダッシュボード用 JSON に変換する
 */
import * as XLSX from "xlsx";

const SYMPTOM_DISEASE_BOUNDARY_COL = 32; // AF列以降を疾患

function cellText(v: unknown): string {
  if (v == null || v === undefined) return "";
  const s = String(v).trim();
  return s === "nan" ? "" : s;
}

function safeInt(v: unknown): number {
  if (v == null || v === "") return 0;
  const n = Number(v);
  return isNaN(n) ? 0 : Math.floor(n);
}

function safeFloat(v: unknown): number | null {
  if (v == null || v === "") return null;
  const n = Number(v);
  return isNaN(n) ? null : n;
}

export type FileType = "symptom_disease" | "evaluation" | "case_summary" | "symptom_disease_simplified";

export function detectWorkbookType(wb: XLSX.WorkBook, fileTypeHint?: FileType): FileType {
  if (fileTypeHint === "evaluation") return "evaluation";
  if (fileTypeHint === "case_summary") return "case_summary";

  const sheetNames = wb.SheetNames || [];
  if (sheetNames.some((n) => n.includes("研修医評価") || n.includes("指導医評価"))) {
    return "evaluation";
  }

  for (const name of sheetNames) {
    const ws = wb.Sheets[name];
    if (!ws || !ws["!ref"]) continue;
    const data = XLSX.utils.sheet_to_json<string[]>(ws, { header: 1, defval: "" }) as string[][];
    const texts = data.slice(0, 10).flat().map(cellText).filter(Boolean);
    const joined = texts.join(" ");

    if (joined.includes("研修医氏名") && joined.includes("経験すべき症候")) {
      return "symptom_disease";
    }
    if (joined.includes("研修医氏名") && (texts.includes("未") || texts.includes("済"))) {
      return "case_summary";
    }
    if (joined.includes("研修医氏名")) {
      const firstDataRow = data.find((r, i) => i > 0 && r.some((c) => cellText(c)));
      if (firstDataRow) {
        const vals = firstDataRow.slice(2).map(cellText);
        const hasJumi = vals.some((v) => v === "済" || v === "未");
        if (hasJumi) return "case_summary";
        const hasNum = vals.some((v) => v && !isNaN(Number(v)));
        if (hasNum) return "symptom_disease_simplified";
      }
    }
  }
  return "case_summary"; // デフォルト
}

export function parseExcelToDashboard(
  file: File,
  fileTypeHint?: FileType
): Promise<{ dashboardData: any; filename?: string; excelBase64?: string }> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = e.target?.result;
        if (!(data instanceof ArrayBuffer)) {
          reject(new Error("ファイルの読み込みに失敗しました"));
          return;
        }
        const wb = XLSX.read(data, { type: "array" });
        const detected = detectWorkbookType(wb, fileTypeHint);

        let dashboardData: any;
        if (detected === "evaluation") {
          dashboardData = parseEvaluation(wb);
          resolve({ dashboardData });
          return;
        }
        if (detected === "symptom_disease") {
          reject(new Error("経験症候・疾患（複数シート形式）はクライアント変換に対応していません。npm run dev でサーバー起動するか、JSON を読み込んでください。"));
          return;
        }
        if (detected === "case_summary") {
          dashboardData = parseCaseSummary(wb);
        } else {
          dashboardData = parseSymptomDiseaseSimplified(wb);
        }
        resolve({ dashboardData });
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = () => reject(new Error("ファイルの読み込みに失敗しました"));
    reader.readAsArrayBuffer(file);
  });
}

const BASE_COLS = ["研修医UMIN ID", "研修医氏名", "施設名", "診療科名", "研修開始日", "研修終了日", "主／並行", "指導医UMIN ID", "指導医氏名"];
const EXCLUDED_SHEET_KEYWORDS = ["メディカルスタッフ", "患者・家族等", "患者家族"];
const RENAME_MAP: Record<string, string> = {
  研修施設: "施設名",
  診療科: "診療科名",
  評価者氏名: "指導医氏名",
  "評価者UMIN ID": "指導医UMIN ID",
};

function getScoreColumns(cols: string[]): string[] {
  return cols.filter((c) => /^[ABC]-\d+\./.test(String(c)));
}

function classifyCompetency(col: string): string {
  if (col.startsWith("A-")) return "A_プロフェッショナリズム";
  if (col.startsWith("B-")) return "B_資質・能力";
  if (col.startsWith("C-")) return "C_基本的診療業務";
  return "その他";
}

function parseEvaluation(wb: XLSX.WorkBook): any {
  const evalDfs: Record<string, Record<string, unknown>[]> = {};

  for (const sheetName of wb.SheetNames || []) {
    if (EXCLUDED_SHEET_KEYWORDS.some((kw) => sheetName.includes(kw))) continue;

    const ws = wb.Sheets[sheetName];
    if (!ws || !ws["!ref"]) continue;

    const data = XLSX.utils.sheet_to_json<string[]>(ws, { header: 1, defval: "" }) as unknown[][];
    const headerRowIdx = data.findIndex(
      (row) =>
        Array.isArray(row) &&
        row.some((c) => cellText(c) === "研修医UMIN ID") &&
        row.some((c) => cellText(c) === "研修医氏名")
    );
    if (headerRowIdx < 0) continue;

    const header = (data[headerRowIdx] as unknown[]).map((c) => {
      let h = cellText(c);
      return RENAME_MAP[h] || h;
    });
    const rows = data.slice(headerRowIdx + 1) as unknown[][];

    const df = rows
      .map((row) => {
        const obj: Record<string, unknown> = {};
        header.forEach((h, i) => {
          obj[h] = row[i];
        });
        return obj;
      })
      .filter((r) => Object.values(r).some((v) => cellText(v)))
      .filter((r) => cellText(r["研修医UMIN ID"]) !== "研修医UMIN ID");

    let sourceName = "研修医評価";
    if (sheetName.includes("指導")) sourceName = "指導医評価";
    else if (sheetName.includes("研修医") || sheetName.includes("自己")) sourceName = "研修医評価";
    else sourceName = sheetName;

    if (evalDfs[sourceName]) sourceName = `${sourceName}_${sheetName}`;
    evalDfs[sourceName] = df;
  }

  const longAll: Array<Record<string, unknown>> = [];
  for (const [sourceName, df] of Object.entries(evalDfs)) {
    const baseCols = BASE_COLS.filter((c) => df[0] && c in df[0]);
    const scoreCols = getScoreColumns(Object.keys(df[0] || {}));

    for (const row of df) {
      for (const col of scoreCols) {
        const score = safeFloat(row[col]);
        const base: Record<string, unknown> = {};
        baseCols.forEach((c) => (base[c] = row[c]));
        longAll.push({
          ...base,
          評価項目: col,
          評価点: row[col],
          評価点_num: score,
          評価群: classifyCompetency(col),
          評価元: sourceName,
        });
      }
    }
  }

  if (longAll.length === 0) {
    throw new Error("評価データが見つかりませんでした。研修医評価・指導医評価シートを確認してください。");
  }

  const residentSummary = makeEvalResidentSummary(longAll);
  const departmentSummary = makeEvalDepartmentSummary(longAll);
  const evaluatorSummary = makeEvalEvaluatorSummary(longAll);
  const radarSource = makeEvalRadarSource(longAll);
  const itemScores = makeEvalItemScores(longAll);
  const rotationStatus = makeEvalRotationStatus(longAll);
  const alerts = makeEvalAlerts(longAll);

  const alertsSupervisor = alerts.filter((a) => a.評価元 === "指導医評価");
  const avgScore =
    residentSummary.length > 0
      ? residentSummary.reduce((s, r) => s + ((r.平均点_全体 as number) ?? 0), 0) / residentSummary.length
      : null;

  return {
    type: "evaluation",
    stats: {
      total_residents: residentSummary.length,
      avg_overall_score: avgScore != null ? Math.round(avgScore * 1000) / 1000 : null,
      total_low_evals: alertsSupervisor.length,
      total_zero_evals: alerts.filter((a) => a.状態 === "0点").length,
    },
    residents: residentSummary,
    departments: departmentSummary,
    evaluators: evaluatorSummary,
    radar: radarSource,
    item_scores: itemScores,
    rotations: rotationStatus,
    alerts,
  };
}

function makeEvalResidentSummary(longAll: Array<Record<string, unknown>>): any[] {
  const byResident = new Map<string, Array<Record<string, unknown>>>();
  for (const r of longAll) {
    const key = `${cellText(r["研修医UMIN ID"])}|${cellText(r["研修医氏名"])}`;
    if (!byResident.has(key)) byResident.set(key, []);
    byResident.get(key)!.push(r);
  }

  const rows: any[] = [];
  for (const [key, sub] of byResident) {
    const [rid, rname] = key.split("|");
    const valid = sub.filter(evalValidFilter);
    const validSupervisor = valid.filter((r) => r.評価元 === "指導医評価");

    const avgAll = valid.length ? valid.reduce((s, r) => s + ((r.評価点_num as number) ?? 0), 0) / valid.length : null;
    const avgA =
      valid.filter((r) => r.評価群 === "A_プロフェッショナリズム").length > 0
        ? valid
            .filter((r) => r.評価群 === "A_プロフェッショナリズム")
            .reduce((s, r) => s + ((r.評価点_num as number) ?? 0), 0) /
          valid.filter((r) => r.評価群 === "A_プロフェッショナリズム").length
        : null;
    const avgB =
      valid.filter((r) => r.評価群 === "B_資質・能力").length > 0
        ? valid
            .filter((r) => r.評価群 === "B_資質・能力")
            .reduce((s, r) => s + ((r.評価点_num as number) ?? 0), 0) /
          valid.filter((r) => r.評価群 === "B_資質・能力").length
        : null;
    const avgC =
      valid.filter((r) => r.評価群 === "C_基本的診療業務").length > 0
        ? valid
            .filter((r) => r.評価群 === "C_基本的診療業務")
            .reduce((s, r) => s + ((r.評価点_num as number) ?? 0), 0) /
          valid.filter((r) => r.評価群 === "C_基本的診療業務").length
        : null;
    const minScore = valid.length ? Math.min(...valid.map((r) => (r.評価点_num as number) ?? 0)) : null;
    const low2 = validSupervisor.filter((r) => ((r.評価点_num as number) ?? 0) < 2).length;
    const low1 = validSupervisor.filter((r) => ((r.評価点_num as number) ?? 0) <= 1).length;
    const zeroCount = sub.filter((r) => (r.評価点_num as number) === 0).length;

    rows.push({
      "研修医UMIN ID": rid,
      研修医氏名: rname,
      評価件数: valid.length,
      平均点_全体: avgAll != null ? Math.round(avgAll * 1000) / 1000 : null,
      平均点_A: avgA != null ? Math.round(avgA * 1000) / 1000 : null,
      平均点_B: avgB != null ? Math.round(avgB * 1000) / 1000 : null,
      平均点_C: avgC != null ? Math.round(avgC * 1000) / 1000 : null,
      最低点: minScore != null ? Math.round(minScore * 1000) / 1000 : null,
      "2点未満件数": low2,
      "1点以下件数": low1,
      "0点件数": zeroCount,
    });
  }

  rows.sort((a, b) => {
    const aAvg = (a.平均点_全体 as number) ?? 999;
    const bAvg = (b.平均点_全体 as number) ?? 999;
    if (aAvg !== bAvg) return aAvg - bAvg;
    if ((b["2点未満件数"] as number) !== (a["2点未満件数"] as number))
      return (b["2点未満件数"] as number) - (a["2点未満件数"] as number);
    return (b["0点件数"] as number) - (a["0点件数"] as number);
  });
  return rows;
}

function makeEvalDepartmentSummary(longAll: Array<Record<string, unknown>>): any[] {
  const valid = longAll.filter(evalValidFilter);
  if (valid.length === 0) return [];

  const byDept = new Map<string, number[]>();
  for (const r of valid) {
    const key = `${cellText(r.施設名)}|${cellText(r.診療科名)}|${cellText(r.評価元)}`;
    if (!byDept.has(key)) byDept.set(key, []);
    byDept.get(key)!.push((r.評価点_num as number) ?? 0);
  }

  return Array.from(byDept.entries()).map(([key, scores]) => {
    const [施設名, 診療科名, 評価元] = key.split("|");
    const avg = scores.reduce((a, b) => a + b, 0) / scores.length;
    return {
      施設名,
      診療科名,
      評価元,
      評価件数: scores.length,
      平均点: Math.round(avg * 1000) / 1000,
      最低点: Math.min(...scores),
      最高点: Math.max(...scores),
    };
  });
}

function makeEvalEvaluatorSummary(longAll: Array<Record<string, unknown>>): any[] {
  const valid = longAll.filter(evalValidFilter);
  if (valid.length === 0) return [];

  const byEval = new Map<string, number[]>();
  for (const r of valid) {
    const key = `${cellText(r["指導医UMIN ID"])}|${cellText(r["指導医氏名"])}|${cellText(r.評価元)}`;
    if (!byEval.has(key)) byEval.set(key, []);
    byEval.get(key)!.push((r.評価点_num as number) ?? 0);
  }

  return Array.from(byEval.entries()).map(([key, scores]) => {
    const [指導医UMIN_ID, 指導医氏名, 評価元] = key.split("|");
    const avg = scores.reduce((a, b) => a + b, 0) / scores.length;
    return {
      "指導医UMIN ID": 指導医UMIN_ID,
      指導医氏名,
      評価元,
      評価件数: scores.length,
      平均点: Math.round(avg * 1000) / 1000,
    };
  });
}

function makeEvalRadarSource(longAll: Array<Record<string, unknown>>): any[] {
  const valid = longAll.filter(evalValidFilter);
  if (valid.length === 0) return [];

  const byGroup = new Map<string, number[]>();
  for (const r of valid) {
    const key = `${cellText(r["研修医UMIN ID"])}|${cellText(r["研修医氏名"])}|${cellText(r.評価元)}|${cellText(r.評価群)}`;
    if (!byGroup.has(key)) byGroup.set(key, []);
    byGroup.get(key)!.push((r.評価点_num as number) ?? 0);
  }

  return Array.from(byGroup.entries()).map(([key, scores]) => {
    const [研修医UMIN_ID, 研修医氏名, 評価元, 評価群] = key.split("|");
    const avg = scores.reduce((a, b) => a + b, 0) / scores.length;
    return {
      "研修医UMIN ID": 研修医UMIN_ID,
      研修医氏名,
      評価元,
      評価群,
      平均点: Math.round(avg * 1000) / 1000,
    };
  });
}

function makeEvalItemScores(longAll: Array<Record<string, unknown>>): any[] {
  const valid = longAll.filter(evalValidFilter);
  if (valid.length === 0) return [];

  const byItem = new Map<string, number[]>();
  for (const r of valid) {
    const key = `${cellText(r["研修医UMIN ID"])}|${cellText(r["研修医氏名"])}|${cellText(r.評価元)}|${cellText(r.評価群)}|${cellText(r.評価項目)}`;
    if (!byItem.has(key)) byItem.set(key, []);
    byItem.get(key)!.push((r.評価点_num as number) ?? 0);
  }

  return Array.from(byItem.entries()).map(([key, scores]) => {
    const parts = key.split("|");
    const avg = scores.reduce((a, b) => a + b, 0) / scores.length;
    return {
      "研修医UMIN ID": parts[0],
      研修医氏名: parts[1],
      評価元: parts[2],
      評価群: parts[3],
      評価項目: parts[4],
      平均点: Math.round(avg * 1000) / 1000,
      件数: scores.length,
    };
  });
}

function makeEvalRotationStatus(longAll: Array<Record<string, unknown>>): any[] {
  const matchCols = ["研修医UMIN ID", "研修医氏名", "施設名", "診療科名", "研修開始日", "研修終了日"].filter(
    (c) => longAll[0] && c in longAll[0]
  );

  const byRot = new Map<string, Record<string, unknown>>();
  for (const r of longAll) {
    const key = matchCols.map((c) => cellText(r[c])).join("|");
    if (!byRot.has(key)) {
      byRot.set(key, {});
      matchCols.forEach((c) => (byRot.get(key)![c] = r[c]));
      if ("主／並行" in (longAll[0] || {})) byRot.get(key)!["主／並行"] = r["主／並行"];
      if (longAll[0] && "指導医UMIN ID" in longAll[0]) byRot.get(key)!["指導医UMIN ID"] = r["指導医UMIN ID"];
      if (longAll[0] && "指導医氏名" in longAll[0]) byRot.get(key)!["指導医氏名"] = r["指導医氏名"];
    }
  }

  const aggBySource = new Map<string, { avg: number; cnt: number }[]>();
  for (const r of longAll) {
    const key = matchCols.map((c) => cellText(r[c])).join("|");
    const score = r.評価点_num as number | null;
    if (score == null || score <= 0) continue;
    const src = cellText(r.評価元);
    const subKey = `${key}|${src}`;
    if (!aggBySource.has(subKey)) aggBySource.set(subKey, []);
    aggBySource.get(subKey)!.push({ avg: score, cnt: 1 });
  }

  const rotations: any[] = [];
  for (const [key, row] of byRot) {
    const parts = key.split("|");
    const r: any = { ...row };
    for (const src of ["研修医評価", "指導医評価"]) {
      const subKey = `${key}|${src}`;
      const items = aggBySource.get(subKey) || [];
      const cnt = items.reduce((s, x) => s + x.cnt, 0);
      const avg = cnt > 0 ? items.reduce((s, x) => s + x.avg * x.cnt, 0) / cnt : null;
      r[`${src}_平均点`] = avg != null ? Math.round(avg * 1000) / 1000 : null;
      r[`${src}_件数`] = cnt;
    }
    r.研修医評価あり = (r.研修医評価_件数 as number) > 0;
    r.指導医評価あり = (r.指導医評価_件数 as number) > 0;
    r.入力状況 =
      r.研修医評価あり && r.指導医評価あり
        ? "両方入力済"
        : r.研修医評価あり
          ? "研修医のみ"
          : r.指導医評価あり
            ? "指導医のみ"
            : "未入力";
    rotations.push(r);
  }
  return rotations;
}

function evalValidFilter(r: Record<string, unknown>): boolean {
  const n = r.評価点_num as number | null;
  return n != null && n > 0;
}

function makeEvalAlerts(longAll: Array<Record<string, unknown>>): any[] {
  const low = longAll.filter((r) => {
    const n = r.評価点_num as number | null;
    return n != null && n > 0 && n < 2;
  });
  return low.map((r) => ({
    評価元: r.評価元,
    研修医氏名: r["研修医氏名"],
    "研修医UMIN ID": r["研修医UMIN ID"],
    施設名: r.施設名,
    診療科名: r.診療科名,
    指導医氏名: r["指導医氏名"],
    評価群: r.評価群,
    評価項目: r.評価項目,
    評価点: r.評価点_num as number,
    状態: "低評価(<2)",
  }));
}

function parseCaseSummary(wb: XLSX.WorkBook): any {
  const ws = wb.Sheets[wb.SheetNames[0]];
  if (!ws) throw new Error("シートが見つかりません");
  const rows = XLSX.utils.sheet_to_json<Record<string, unknown>>(ws, { defval: "" });
  if (!rows.length) throw new Error("データがありません");

  const cols = Object.keys(rows[0] || {});
  if (!cols.includes("研修医氏名")) {
    throw new Error("ヘッダーに「研修医氏名」が見つかりませんでした。病歴要約等入力状況一覧ではない可能性があります。");
  }

  const excludeCols = ["研修医氏名", "UMIN ID"];
  const itemCols = cols.filter((c) => !excludeCols.includes(c) && !String(c).startsWith("Unnamed"));
  const colToIdx: Record<string, number> = {};
  cols.forEach((c, i) => (colToIdx[c] = i));

  const itemStats: Record<string, { 済: number; 未: number }> = {};
  itemCols.forEach((c) => (itemStats[c] = { 済: 0, 未: 0 }));

  const summaryRows: any[] = [];
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const residentName = cellText(row["研修医氏名"]);
    if (!residentName) continue;

    const uminId = cellText(row["UMIN ID"] || "");
    let completed = 0,
      uncompleted = 0;

    for (const col of itemCols) {
      const val = cellText(row[col]);
      if (val === "済") {
        completed++;
        itemStats[col].済++;
      } else {
        uncompleted++;
        itemStats[col].未++;
      }
    }

    summaryRows.push({
      研修医氏名: residentName,
      "UMIN ID": uminId,
      入力済件数: completed,
      未入力件数: uncompleted,
      全体進捗率: itemCols.length ? Math.round((completed / itemCols.length) * 1000) / 1000 : 0,
      対象項目数: itemCols.length,
    });
  }

  summaryRows.sort((a, b) => b.入力済件数 - a.入力済件数 || String(a.研修医氏名).localeCompare(String(b.研修医氏名)));

  const items = itemCols.map((col) => {
    const s = itemStats[col];
    const total = s.済 + s.未;
    return {
      項目名: col,
      済_件数: s.済,
      未_件数: s.未,
      入力率: total > 0 ? Math.round((s.済 / total) * 1000) / 1000 : 0,
    };
  });
  items.sort((a, b) => b.済_件数 - a.済_件数);

  const itemDetails: any[] = [];
  for (const row of rows) {
    const residentName = cellText(row["研修医氏名"]);
    if (!residentName) continue;
    for (const col of itemCols) {
      const val = cellText(row[col]);
      const status = val === "済" ? "済" : "未";
      const excelCol = (colToIdx[col] ?? 0) + 1;
      const 区分 = excelCol >= SYMPTOM_DISEASE_BOUNDARY_COL ? "疾患" : "症候";
      itemDetails.push({ 研修医氏名: residentName, 項目名: col, 区分, 状態: status });
    }
  }

  const avgProgress =
    summaryRows.length > 0
      ? summaryRows.reduce((s, r) => s + (r.全体進捗率 || 0), 0) / summaryRows.length
      : 0;

  return {
    type: "case_summary",
    stats: {
      total_residents: summaryRows.length,
      avg_overall_progress: Math.round(avgProgress * 1000) / 1000,
      total_items: itemCols.length,
    },
    residents: summaryRows,
    items,
    item_details: itemDetails,
  };
}

function parseSymptomDiseaseSimplified(wb: XLSX.WorkBook): any {
  const ws = wb.Sheets[wb.SheetNames[0]];
  if (!ws) throw new Error("シートが見つかりません");
  const rows = XLSX.utils.sheet_to_json<Record<string, unknown>>(ws, { defval: "" });
  if (!rows.length) throw new Error("データがありません");

  const cols = Object.keys(rows[0] || {});
  if (!cols.includes("研修医氏名")) {
    throw new Error("ヘッダーに「研修医氏名」が見つかりませんでした。");
  }

  const excludeCols = ["研修医氏名", "UMIN ID"];
  const itemCols = cols.filter((c) => !excludeCols.includes(c) && !String(c).startsWith("Unnamed"));
  const colToIdx: Record<string, number> = {};
  cols.forEach((c, i) => (colToIdx[c] = i));

  const summaryRows: any[] = [];
  const alertRows: any[] = [];
  const heatmapData: any[] = [];

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const residentName = cellText(row["研修医氏名"]);
    if (!residentName) continue;

    const uminId = cellText(row["UMIN ID"] || "");
    const totalItems = itemCols.length;
    let experiencedCount = 0;
    const symptomCount = itemCols.filter((c) => (colToIdx[c] ?? 0) + 1 < SYMPTOM_DISEASE_BOUNDARY_COL).length;
    const diseaseCount = totalItems - symptomCount;
    let symptomExp = 0,
      diseaseExp = 0;

    for (const col of itemCols) {
      const count = safeInt(row[col]);
      const excelCol = (colToIdx[col] ?? 0) + 1;
      const section = excelCol >= SYMPTOM_DISEASE_BOUNDARY_COL ? "疾患" : "症候";

      heatmapData.push({
        研修医氏名: residentName,
        研修医UMIN_ID: uminId,
        区分: section,
        項目名: col,
        経験数: count,
        承認数: count,
        病歴要約提出: 0,
        病歴要約確認: 0,
        外科手術要約提出: 0,
        外科手術要約確認: 0,
        シート名: "簡易版一覧",
      });

      if (count > 0) {
        experiencedCount++;
        if (section === "症候") symptomExp++;
        else diseaseExp++;
      } else {
        alertRows.push({
          研修医氏名: residentName,
          研修医UMIN_ID: uminId,
          区分: section,
          項目名: col,
          経験数: 0,
          承認数: 0,
          病歴要約提出: 0,
          病歴要約確認: 0,
          外科手術要約提出: 0,
          外科手術要約確認: 0,
          状態: "未経験",
        });
      }
    }

    summaryRows.push({
      研修医氏名: residentName,
      研修医UMIN_ID: uminId,
      総項目数: totalItems,
      経験済項目数: experiencedCount,
      承認済項目数: experiencedCount,
      症候達成率: symptomCount > 0 ? Math.round((symptomExp / symptomCount) * 1000) / 1000 : 0,
      疾患達成率: diseaseCount > 0 ? Math.round((diseaseExp / diseaseCount) * 1000) / 1000 : 0,
      全体達成率: totalItems > 0 ? Math.round((experiencedCount / totalItems) * 1000) / 1000 : 0,
      承認率: totalItems > 0 ? Math.round((experiencedCount / totalItems) * 1000) / 1000 : 0,
      病歴要約未提出件数: 0,
      病歴要約未確認件数: 0,
      外科手術要約未提出件数: 0,
      外科手術要約未確認件数: 0,
      要対応件数: totalItems - experiencedCount,
    });
  }

  summaryRows.sort((a, b) => b.要対応件数 - a.要対応件数 || (a.全体達成率 || 0) - (b.全体達成率 || 0));

  const rawDf = heatmapData.map((r) => ({
    ...r,
    研修医UMIN_ID: r.研修医UMIN_ID,
  }));

  const items = rawDf.map((r) => ({
    研修医氏名: r.研修医氏名,
    "研修医UMIN ID": r.研修医UMIN_ID,
    区分: r.区分,
    項目名: r.項目名,
    経験数: r.経験数,
    承認数: r.承認数,
    状態: r.経験数 > 0 ? "達成" : "未経験",
  }));

  const residents = summaryRows.map((r) => {
    const { 研修医UMIN_ID, ...rest } = r;
    return { ...rest, "研修医UMIN ID": 研修医UMIN_ID };
  });
  const alerts = alertRows.map((r) => {
    const { 研修医UMIN_ID, ...rest } = r;
    return { ...rest, "研修医UMIN ID": 研修医UMIN_ID };
  });

  const avgRate = residents.length > 0 ? residents.reduce((s, r) => s + (r.全体達成率 || 0), 0) / residents.length : 0;

  return {
    type: "symptom_disease",
    stats: {
      total_residents: residents.length,
      avg_overall_rate: Math.round(avgRate * 1000) / 1000,
      avg_approval_rate: Math.round(avgRate * 1000) / 1000,
      total_alerts: alerts.length,
    },
    residents,
    alerts,
    items,
  };
}
