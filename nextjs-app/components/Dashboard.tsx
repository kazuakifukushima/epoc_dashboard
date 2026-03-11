"use client";

import React from "react";
import {
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip,
  ResponsiveContainer, Legend, Cell,
  RadarChart, Radar, PolarGrid, PolarAngleAxis, PolarRadiusAxis,
  PieChart, Pie,
} from "recharts";
import {
  AlertCircle, FileText, Activity, Users, ClipboardCheck,
  CheckSquare, TrendingUp, TrendingDown, ChevronRight,
  ChevronLeft, ChevronUp, ChevronDown, AlertTriangle, Award, BookOpen, Stethoscope,
  BarChart2, ShieldCheck, Check, X, Download,
} from "lucide-react";
import * as XLSX from "xlsx";

// ============================================================
// Design Tokens
// ============================================================
const C = {
  blue: "#1d4ed8",
  blueSoft: "#eff6ff",
  blueBorder: "#bfdbfe",
  purple: "#7c3aed",
  purpleSoft: "#f5f3ff",
  purpleBorder: "#ddd6fe",
  green: "#16a34a",
  greenSoft: "#f0fdf4",
  greenBorder: "#bbf7d0",
  amber: "#d97706",
  amberSoft: "#fffbeb",
  amberBorder: "#fde68a",
  red: "#dc2626",
  redSoft: "#fef2f2",
  redBorder: "#fecaca",
  slate900: "#0f172a",
  slate700: "#374151",
  slate500: "#64748b",
  slate300: "#cbd5e1",
  slate100: "#f1f5f9",
  white: "#ffffff",
};

// ============================================================
// Shared Primitives
// ============================================================
function StatCard({
  title, value, sub, icon: Icon, color, bg,
}: {
  title: string; value: string | number; sub?: string;
  icon: any; color: string; bg: string;
}) {
  return (
    <div style={{
      background: C.white, borderRadius: 16, padding: "20px 24px",
      boxShadow: "0 1px 3px rgba(0,0,0,0.06), 0 4px 12px rgba(0,0,0,0.04)",
      display: "flex", alignItems: "flex-start", gap: 16,
    }}>
      <div style={{
        width: 44, height: 44, borderRadius: 12,
        background: bg, display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0,
      }}>
        <Icon size={22} color={color} />
      </div>
      <div style={{ minWidth: 0 }}>
        <div style={{ fontSize: 12, color: C.slate500, fontWeight: 500, marginBottom: 4 }}>{title}</div>
        <div style={{ fontSize: 26, fontWeight: 700, color: C.slate900, lineHeight: 1.1 }}>{value}</div>
        {sub && <div style={{ fontSize: 12, color: C.slate500, marginTop: 4 }}>{sub}</div>}
      </div>
    </div>
  );
}

function ProgressBar({ pct, color = C.blue, height = 8 }: { pct: number; color?: string; height?: number }) {
  const clamped = Math.min(100, Math.max(0, pct));
  return (
    <div style={{ background: C.slate100, borderRadius: 99, height, overflow: "hidden", minWidth: 60 }}>
      <div style={{
        width: `${clamped}%`, height: "100%",
        background: color, borderRadius: 99,
        transition: "width 0.4s ease",
      }} />
    </div>
  );
}

function Badge({ label, color, bg }: { label: string; color: string; bg: string }) {
  return (
    <span style={{
      display: "inline-flex", alignItems: "center",
      padding: "2px 10px", borderRadius: 99,
      fontSize: 11, fontWeight: 600,
      color, background: bg,
    }}>{label}</span>
  );
}

function ScoreChip({ score }: { score: number | null | undefined }) {
  if (score == null) return <span style={{ color: C.slate500 }}>—</span>;
  const color = score === 0 ? C.red : score < 2 ? C.amber : score >= 3.5 ? C.green : C.slate700;
  const bg = score === 0 ? C.redSoft : score < 2 ? C.amberSoft : score >= 3.5 ? C.greenSoft : C.slate100;
  return (
    <span style={{
      display: "inline-block", padding: "2px 10px", borderRadius: 99,
      fontSize: 13, fontWeight: 700, color, background: bg,
    }}>{score.toFixed(2)}</span>
  );
}

function SectionTitle({ icon: Icon, label, color = C.slate700 }: { icon?: any; label: string; color?: string }) {
  return (
    <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 16 }}>
      {Icon && <Icon size={18} color={color} />}
      <span style={{ fontSize: 15, fontWeight: 700, color: C.slate900 }}>{label}</span>
    </div>
  );
}

function Card({ children, style }: { children: React.ReactNode; style?: React.CSSProperties }) {
  return (
    <div style={{
      background: C.white, borderRadius: 16, padding: 24,
      boxShadow: "0 1px 3px rgba(0,0,0,0.06), 0 4px 12px rgba(0,0,0,0.04)",
      ...style,
    }}>{children}</div>
  );
}

function ResidentSelector({
  names, selected, onChange, accentColor,
}: {
  names: string[]; selected: string; onChange: (v: string) => void; accentColor: string;
}) {
  return (
    <div style={{
      display: "flex", alignItems: "center", gap: 8,
      background: C.white, padding: "8px 14px", borderRadius: 10,
      border: `1.5px solid ${C.slate300}`,
      boxShadow: "0 1px 3px rgba(0,0,0,0.06)",
    }}>
      <Users size={14} color={C.slate500} />
      <select
        value={selected}
        onChange={(e) => onChange(e.target.value)}
        style={{
          border: "none", outline: "none", background: "transparent",
          fontSize: 13, color: C.slate700, cursor: "pointer", fontWeight: 600,
          fontFamily: "inherit",
        }}
      >
        <option value="all">全研修医サマリー</option>
        {names.map(n => <option key={n} value={n}>{n}</option>)}
      </select>
    </div>
  );
}

// ============================================================
// Tooltip
// ============================================================
const CustomTooltip = ({ active, payload, label }: any) => {
  if (!active || !payload?.length) return null;
  return (
    <div style={{
      background: C.white, borderRadius: 10, padding: "10px 14px",
      boxShadow: "0 8px 24px rgba(0,0,0,0.12)", border: `1px solid ${C.slate100}`, fontSize: 13,
    }}>
      <div style={{ fontWeight: 700, marginBottom: 6, color: C.slate900 }}>{label}</div>
      {payload.map((p: any, i: number) => (
        <div key={i} style={{ display: "flex", alignItems: "center", gap: 6, color: p.color, marginBottom: 2 }}>
          <div style={{ width: 8, height: 8, borderRadius: 2, background: p.color }} />
          <span style={{ color: C.slate700 }}>{p.name}:</span>
          <span style={{ fontWeight: 600 }}>{typeof p.value === "number" ? p.value.toLocaleString() : p.value}</span>
        </div>
      ))}
    </div>
  );
};

// ============================================================
// Excel形式: 症候と疾患で表を2つに分けて表示
// ============================================================
function ResidentItemPivotTable({ items }: { items: any[] }) {
  const [sortOrder, setSortOrder] = React.useState<Record<string, "desc" | "asc">>({ 症候: "desc", 疾患: "desc" });

  const statusStyle = (status: string) => {
    if (status === "達成") return { bg: C.greenSoft, color: C.green };
    if (status === "未経験") return { bg: "#f1f5f9", color: C.slate500 };
    return { bg: C.amberSoft, color: C.amber };
  };

  const { byCategory, residents, statusMap, rateMap } = React.useMemo(() => {
    const byCat = new Map<string, string[]>();
    for (const r of items) {
      const 区分 = r["区分"] || "症候";
      const 項目名 = r["項目名"];
      if (!byCat.has(区分)) byCat.set(区分, []);
      const list = byCat.get(区分)!;
      if (!list.includes(項目名)) list.push(項目名);
    }
    for (const list of byCat.values()) list.sort();

    const byCategory = Array.from(byCat.entries())
      .sort((a, b) => (a[0] === "症候" ? 0 : 1) - (b[0] === "症候" ? 0 : 1))
      .map(([区分, 項目]) => ({ 区分, 項目 }));

    const residents = Array.from(new Set(items.map((i: any) => i["研修医氏名"]))).filter(Boolean).sort();

    const statusMap = new Map<string, string>();
    const rateMap = new Map<string, number>();
    for (const r of items) {
      const key = `${r["研修医氏名"]}\t${r["区分"]}\t${r["項目名"]}`;
      statusMap.set(key, r["状態"] || "未経験");
    }
    for (const name of residents) {
      for (const { 区分, 項目 } of byCategory) {
        let achieved = 0;
        for (const 項目名 of 項目) {
          const status = statusMap.get(`${name}\t${区分}\t${項目名}`) || "—";
          if (status === "達成") achieved++;
        }
        const rate = 項目.length > 0 ? Math.round((achieved / 項目.length) * 100) : 0;
        rateMap.set(`${name}\t${区分}`, rate);
      }
    }

    return { byCategory, residents, statusMap, rateMap };
  }, [items]);

  if (items.length === 0) {
    return <div style={{ textAlign: "center", padding: 24, color: C.slate500 }}>項目データがありません</div>;
  }

  const categoryStyle = (区分: string) => ({
    bg: 区分 === "症候" ? C.blueSoft : C.purpleSoft,
    color: 区分 === "症候" ? C.blue : C.purple,
    borderColor: 区分 === "症候" ? C.blueBorder : C.purpleBorder,
  });

  const rateColor = (rate: number) => rate >= 80 ? C.green : rate >= 50 ? C.amber : C.red;

  const toggleSort = (区分: string) => {
    setSortOrder((prev) => ({ ...prev, [区分]: prev[区分] === "desc" ? "asc" : "desc" }));
  };

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 24 }}>
      {byCategory.map(({ 区分, 項目 }) => {
        const style = categoryStyle(区分);
        const order = sortOrder[区分] ?? "desc";
        const sortedResidents = [...residents].sort((a, b) => {
          const ra = rateMap.get(`${a}\t${区分}`) ?? 0;
          const rb = rateMap.get(`${b}\t${区分}`) ?? 0;
          if (ra !== rb) return order === "desc" ? rb - ra : ra - rb;
          return a.localeCompare(b);
        });

        return (
          <div key={区分} style={{ border: `2px solid ${style.borderColor}`, borderRadius: 12, overflow: "hidden", background: style.bg }}>
            <div style={{ padding: "12px 16px", fontWeight: 700, fontSize: 14, color: style.color, background: style.bg, borderBottom: `2px solid ${style.borderColor}` }}>
              {区分}
            </div>
            <div style={{ overflowX: "auto", background: C.white }}>
              <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
                <thead>
                  <tr style={{ background: C.slate100 }}>
                    <th style={{ padding: "10px 12px", textAlign: "left", fontWeight: 600, color: C.slate700, fontSize: 12, whiteSpace: "nowrap", position: "sticky", left: 0, background: C.slate100, zIndex: 2, minWidth: 120 }}>研修医</th>
                    <th
                      style={{ padding: "10px 12px", textAlign: "center", fontWeight: 600, color: C.slate700, fontSize: 12, whiteSpace: "nowrap", position: "sticky", left: 120, background: C.slate100, zIndex: 2, minWidth: 72, boxShadow: "2px 0 4px rgba(0,0,0,0.06)", cursor: "pointer", userSelect: "none", borderRadius: 4 }}
                      onClick={() => toggleSort(区分)}
                      onMouseEnter={(e) => { e.currentTarget.style.background = "#e2e8f0"; }}
                      onMouseLeave={(e) => { e.currentTarget.style.background = C.slate100; }}
                      title={order === "desc" ? "クリックで低い順に並べ替え" : "クリックで高い順に並べ替え"}
                    >
                      <span style={{ display: "inline-flex", alignItems: "center", gap: 4 }}>
                        達成率
                        {order === "desc" ? <ChevronDown size={14} /> : <ChevronUp size={14} />}
                      </span>
                    </th>
                    {項目.map((項目名) => (
                      <th key={項目名} style={{ padding: "10px 8px", textAlign: "center", fontWeight: 600, color: C.slate700, fontSize: 11, minWidth: 70 }}>{項目名}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {sortedResidents.map((name) => {
                    const rate = rateMap.get(`${name}\t${区分}`) ?? 0;
                    return (
                    <tr key={name} style={{ borderBottom: `1px solid ${C.slate100}` }}>
                      <td style={{ padding: "8px 12px", fontWeight: 600, color: C.blue, whiteSpace: "nowrap", position: "sticky", left: 0, background: C.white, zIndex: 1, minWidth: 120 }}>{name}</td>
                      <td style={{ padding: "8px 12px", textAlign: "center", fontWeight: 700, color: rateColor(rate), position: "sticky", left: 120, background: C.white, zIndex: 1, minWidth: 72, boxShadow: "2px 0 4px rgba(0,0,0,0.06)" }}>
                        {rate}%
                      </td>
                      {項目.map((項目名) => {
                        const status = statusMap.get(`${name}\t${区分}\t${項目名}`) || "—";
                        const { bg, color } = status === "—" ? { bg: "transparent", color: C.slate300 } : statusStyle(status);
                        return (
                          <td key={項目名} style={{ padding: "6px 8px", textAlign: "center", minWidth: 70 }}>
                            {status !== "—" ? (
                              <span style={{ display: "inline-block", padding: "2px 6px", borderRadius: 4, fontSize: 11, fontWeight: 500, background: bg, color }}>{status}</span>
                            ) : (
                              <span style={{ color: C.slate300 }}>—</span>
                            )}
                          </td>
                        );
                      })}
                    </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          </div>
        );
      })}
    </div>
  );
}

// ============================================================
// Resident Item Table (個別研修医用・縦長リスト形式)
// ============================================================
function ResidentItemGrid({ items, residentName }: { items: any[]; residentName: string }) {
  const statusStyle = (status: string) => {
    if (status === "達成") return { bg: C.greenSoft, color: C.green };
    if (status === "未経験") return { bg: "#f1f5f9", color: C.slate500 };
    return { bg: C.amberSoft, color: C.amber };
  };

  return (
    <div style={{ background: C.white, borderRadius: 12, overflow: "hidden", border: `1px solid ${C.slate100}` }}>
      {items.length === 0 ? (
        <div style={{ textAlign: "center", padding: 24, color: C.slate500 }}>項目データがありません</div>
      ) : (
        <div style={{ overflowX: "auto" }}>
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
            <thead>
              <tr style={{ background: C.slate100 }}>
                <th style={{ padding: "10px 12px", textAlign: "left", fontWeight: 600, color: C.slate700, fontSize: 12 }}>区分</th>
                <th style={{ padding: "10px 12px", textAlign: "left", fontWeight: 600, color: C.slate700, fontSize: 12 }}>項目名</th>
                <th style={{ padding: "10px 12px", textAlign: "left", fontWeight: 600, color: C.slate700, fontSize: 12 }}>状態</th>
                <th style={{ padding: "10px 12px", textAlign: "right", fontWeight: 600, color: C.slate700, fontSize: 12 }}>経験数</th>
                <th style={{ padding: "10px 12px", textAlign: "right", fontWeight: 600, color: C.slate700, fontSize: 12 }}>承認数</th>
              </tr>
            </thead>
            <tbody>
              {items.map((item: any, i: number) => {
                const { bg, color } = statusStyle(item["状態"] || "未経験");
                return (
                  <tr key={i} style={{ borderBottom: `1px solid ${C.slate100}` }}>
                    <td style={{ padding: "8px 12px", color: item["区分"] === "症候" ? C.blue : C.purple, fontWeight: 500 }}>{item["区分"]}</td>
                    <td style={{ padding: "8px 12px", color: C.slate900 }}>{item["項目名"]}</td>
                    <td style={{ padding: "8px 12px" }}>
                      <span style={{ display: "inline-block", padding: "2px 8px", borderRadius: 6, fontSize: 12, fontWeight: 500, background: bg, color }}>{item["状態"]}</span>
                    </td>
                    <td style={{ padding: "8px 12px", textAlign: "right", color: C.slate700 }}>{item["経験数"] ?? "—"}</td>
                    <td style={{ padding: "8px 12px", textAlign: "right", color: C.slate700 }}>{item["承認数"] ?? "—"}</td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}

// ============================================================
// Symptom / Disease Dashboard
// ============================================================
function SymptomDiseaseDashboard({ data }: { data: any }) {
  const [selected, setSelected] = React.useState("all");

  const names = Array.from(new Set(
    (data.residents || []).map((r: any) => r["研修医氏名"])
  )).filter(Boolean).sort() as string[];

  const allResidents = [...(data.residents || [])].sort(
    (a: any, b: any) => (b["要対応件数"] || 0) - (a["要対応件数"] || 0)
  );

  const filteredResidents = selected === "all"
    ? allResidents
    : allResidents.filter((r: any) => r["研修医氏名"] === selected);

  const filteredAlerts = data.alerts
    .filter((a: any) => selected === "all" ? true : a["研修医氏名"] === selected)
    .sort((a: any, b: any) => a["研修医氏名"].localeCompare(b["研修医氏名"]));

  const selectedResident = selected !== "all"
    ? data.residents.find((r: any) => r["研修医氏名"] === selected)
    : null;

  const stats = React.useMemo(() => {
    if (selected === "all") {
      return {
        label: `${data.stats.total_residents} 名`,
        rate: (data.stats.avg_overall_rate * 100).toFixed(1),
        alerts: data.stats.total_alerts,
      };
    }
    const r = selectedResident;
    if (!r) return { label: selected, rate: "0.0", alerts: 0 };
    const rate = r["総項目数"] > 0 ? ((r["経験済項目数"] || 0) / r["総項目数"]) * 100 : 0;
    return { label: selected, rate: rate.toFixed(1), alerts: r["要対応件数"] || 0 };
  }, [selected, data, selectedResident]);

  const barData = filteredResidents.slice(0, selected === "all" ? 20 : undefined).map((r: any) => ({
    name: r["研修医氏名"] || "不明",
    経験済: r["経験済項目数"] || 0,
    未経験: (r["総項目数"] || 0) - (r["経験済項目数"] || 0),
    要対応: r["要対応件数"] || 0,
    達成率: r["総項目数"] > 0 ? Math.round(((r["経験済項目数"] || 0) / r["総項目数"]) * 100) : 0,
  }));

  const rateColor = (rate: number) =>
    rate >= 80 ? C.green : rate >= 50 ? C.amber : C.red;

  return (
    <div>
      {/* Dashboard Header */}
      <div style={{ marginBottom: 24, display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 12 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          <div style={{ width: 40, height: 40, background: C.blueSoft, borderRadius: 10, display: "flex", alignItems: "center", justifyContent: "center" }}>
            <Activity size={22} color={C.blue} />
          </div>
          <div>
            <h2 style={{ margin: 0, fontSize: 18, fontWeight: 700, color: C.slate900 }}>経験症候・疾患 状況ダッシュボード</h2>
            <div style={{ fontSize: 12, color: C.slate500 }}>症例入力・経験項目の達成状況をリアルタイムに把握</div>
          </div>
        </div>
        <ResidentSelector names={names} selected={selected} onChange={setSelected} accentColor={C.blue} />
      </div>

      {/* Stat Cards */}
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(200px, 1fr))", gap: 16, marginBottom: 28 }}>
        <StatCard title={selected === "all" ? "対象研修医数" : "選択中研修医"} value={stats.label} icon={Users} color={C.blue} bg={C.blueSoft} />
        <StatCard
          title="経験達成率"
          value={`${stats.rate}%`}
          sub={parseFloat(stats.rate) >= 80 ? "目標達成" : parseFloat(stats.rate) >= 50 ? "進行中" : "要指導"}
          icon={TrendingUp} color={rateColor(parseFloat(stats.rate))} bg={C.greenSoft}
        />
        <StatCard
          title="要対応件数"
          value={stats.alerts}
          sub={stats.alerts > 0 ? "未提出・未承認あり" : "対応不要"}
          icon={AlertTriangle} color={stats.alerts > 0 ? C.red : C.green} bg={stats.alerts > 0 ? C.redSoft : C.greenSoft}
        />
      </div>

      {/* 研修医ごとに統一した達成/未達成カード */}
      {(() => {
        const items = (data.items || []) as any[];
        const hasItems = items.length > 0;
        const residentsToShow = selected === "all" ? allResidents : selectedResident ? [selectedResident] : [];

        return (
          <>
            {hasItems ? (
              <div style={{ display: "flex", flexDirection: "column", gap: 24 }}>
                {selected === "all" ? (
                  /* 全研修医サマリー: Excel形式（研修医=第1列、項目名=第2列以降、セル=達成状態） */
                  <Card style={{ border: `1.5px solid ${C.blueBorder}` }}>
                    <div style={{ marginBottom: 16 }}>
                      <div style={{ fontSize: 12, color: C.slate500, marginBottom: 8, display: "flex", gap: 16 }}>
                        <span><Check size={12} style={{ verticalAlign: "middle", marginRight: 4 }} />緑 = 達成</span>
                        <span><X size={12} style={{ verticalAlign: "middle", marginRight: 4 }} />灰 = 未経験</span>
                        <span><AlertTriangle size={12} style={{ verticalAlign: "middle", marginRight: 4 }} />黄 = 要対応</span>
                      </div>
                      <ResidentItemPivotTable items={items} />
                    </div>
                  </Card>
                ) : (
                  /* 個別研修医: カード形式 */
                  residentsToShow.map((r: any) => {
                    const residentItems = items.filter((i: any) => i["研修医氏名"] === r["研修医氏名"]);
                    const rate = r["総項目数"] > 0 ? ((r["経験済項目数"] || 0) / r["総項目数"]) * 100 : 0;
                    return (
                      <Card key={r["研修医氏名"]} style={{ border: `1.5px solid ${C.blueBorder}` }}>
                        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", flexWrap: "wrap", gap: 16, marginBottom: 20 }}>
                          <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
                            <div style={{ width: 48, height: 48, borderRadius: 99, background: C.blue, display: "flex", alignItems: "center", justifyContent: "center" }}>
                              <span style={{ color: "white", fontWeight: 700, fontSize: 18 }}>{(r["研修医氏名"] || "?")[0]}</span>
                            </div>
                            <div>
                              <div style={{ fontWeight: 700, fontSize: 16, color: C.slate900 }}>{r["研修医氏名"]}</div>
                              <div style={{ fontSize: 12, color: C.slate500 }}>症候・疾患の達成状況</div>
                            </div>
                          </div>
                          <div style={{ display: "flex", alignItems: "center", gap: 16, flexWrap: "wrap" }}>
                            <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                              <span style={{ fontSize: 12, color: C.slate500 }}>達成率</span>
                              <ProgressBar pct={rate} color={rateColor(rate)} height={8} />
                              <span style={{ fontSize: 14, fontWeight: 700, color: rateColor(rate), minWidth: 44 }}>{rate.toFixed(0)}%</span>
                            </div>
                            <div style={{ display: "flex", gap: 12 }}>
                              <Badge label={`経験済 ${r["経験済項目数"] || 0}`} color={C.green} bg={C.greenSoft} />
                              <Badge label={`/ 総数 ${r["総項目数"]}`} color={C.slate700} bg={C.slate100} />
                              {r["要対応件数"] > 0 ? (
                                <Badge label={`要対応 ${r["要対応件数"]}`} color={C.red} bg={C.redSoft} />
                              ) : (
                                <Badge label="✓ 対応済" color={C.green} bg={C.greenSoft} />
                              )}
                            </div>
                          </div>
                        </div>
                        <div style={{ fontSize: 11, color: C.slate500, marginBottom: 8, display: "flex", gap: 16 }}>
                          <span><Check size={12} style={{ verticalAlign: "middle", marginRight: 4 }} />緑 = 達成</span>
                          <span><X size={12} style={{ verticalAlign: "middle", marginRight: 4 }} />灰 = 未経験</span>
                          <span><AlertTriangle size={12} style={{ verticalAlign: "middle", marginRight: 4 }} />黄 = 要対応</span>
                        </div>
                        <ResidentItemGrid items={residentItems} residentName={r["研修医氏名"]} />
                      </Card>
                    );
                  })
                )}
              </div>
            ) : (
              /* items がない場合の従来レイアウト（後方互換） */
              <>
                <div style={{ display: "grid", gridTemplateColumns: selected === "all" ? "1fr 1fr" : "1fr", gap: 24 }}>
                  {selected === "all" && (
                    <Card>
                      <SectionTitle icon={BarChart2} label="研修医別 経験達成状況" color={C.blue} />
                      <ResponsiveContainer width="100%" height={320}>
                        <BarChart data={barData} layout="vertical" margin={{ top: 4, right: 16, left: 8, bottom: 4 }}>
                          <CartesianGrid strokeDasharray="3 3" horizontal={false} stroke="#f1f5f9" />
                          <XAxis type="number" tick={{ fontSize: 11, fill: C.slate500 }} />
                          <YAxis dataKey="name" type="category" width={70} tick={{ fontSize: 11, fill: C.slate700 }} />
                          <Tooltip content={<CustomTooltip />} />
                          <Legend wrapperStyle={{ fontSize: 12 }} />
                          <Bar dataKey="経験済" name="経験済" stackId="a" fill={C.green} radius={[0, 0, 0, 0]} />
                          <Bar dataKey="未経験" name="未経験" stackId="a" fill="#e2e8f0" radius={[0, 4, 4, 0]} />
                        </BarChart>
                      </ResponsiveContainer>
                    </Card>
                  )}
                  <Card>
                    <SectionTitle icon={Users} label={selected === "all" ? "研修医別 達成状況サマリー" : "詳細アラート"} color={C.blue} />
                    {selected === "all" ? (
                      <div style={{ overflowX: "auto" }}>
                        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
                          <thead>
                            <tr style={{ borderBottom: `2px solid ${C.slate100}` }}>
                            {["研修医", "達成率", "経験済 / 総数", "要対応"].map(h => (
                              <th key={h} style={{ padding: "10px 12px", textAlign: h === "研修医" ? "left" : "right", color: C.slate500, fontWeight: 600, fontSize: 12, whiteSpace: "nowrap" }}>{h}</th>
                            ))}
                          </tr>
                          </thead>
                          <tbody>
                            {allResidents.map((r: any, i: number) => {
                              const rate = r["総項目数"] > 0 ? ((r["経験済項目数"] || 0) / r["総項目数"]) * 100 : 0;
                              return (
                                <tr key={i} style={{ borderBottom: `1px solid ${C.slate100}`, cursor: "pointer" }}
                                  onMouseEnter={e => (e.currentTarget.style.background = C.blueSoft)}
                                  onMouseLeave={e => (e.currentTarget.style.background = "transparent")}
                                  onClick={() => setSelected(r["研修医氏名"])}
                                >
                                  <td style={{ padding: "10px 12px", fontWeight: 600, color: C.blue }}>
                                    <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
                                      {r["研修医氏名"]}
                                      <ChevronRight size={12} color={C.slate300} />
                                    </div>
                                  </td>
                                  <td style={{ padding: "10px 12px" }}>
                                    <div style={{ display: "flex", alignItems: "center", gap: 8, justifyContent: "flex-end" }}>
                                      <ProgressBar pct={rate} color={rateColor(rate)} height={6} />
                                      <span style={{ fontSize: 12, fontWeight: 700, color: rateColor(rate), minWidth: 38, textAlign: "right" }}>{rate.toFixed(0)}%</span>
                                    </div>
                                  </td>
                                  <td style={{ padding: "10px 12px", textAlign: "right", color: C.slate700 }}>
                                    <span style={{ color: C.green, fontWeight: 600 }}>{r["経験済項目数"] || 0}</span>
                                    <span style={{ color: C.slate300 }}> / </span>
                                    {r["総項目数"]}
                                  </td>
                                  <td style={{ padding: "10px 12px", textAlign: "right" }}>
                                    {r["要対応件数"] > 0
                                      ? <Badge label={`⚠ ${r["要対応件数"]}`} color={C.red} bg={C.redSoft} />
                                      : <Badge label="✓ 対応済" color={C.green} bg={C.greenSoft} />
                                    }
                                  </td>
                                </tr>
                              );
                            })}
                          </tbody>
                        </table>
                      </div>
                    ) : (
                      <AlertTable alerts={filteredAlerts} showName={false} />
                    )}
                  </Card>
                </div>
                {selected === "all" && filteredAlerts.length > 0 && (
                  <Card style={{ marginTop: 24 }}>
                    <SectionTitle icon={AlertCircle} label={`アラート一覧 — 要対応 ${filteredAlerts.length} 件`} color={C.red} />
                    <AlertTable alerts={filteredAlerts} showName={true} />
                  </Card>
                )}
              </>
            )}
          </>
        );
      })()}
    </div>
  );
}

function AlertTable({ alerts, showName }: { alerts: any[]; showName: boolean }) {
  const displayed = alerts.slice(0, 60);
  const rest = alerts.length - displayed.length;
  return (
    <div style={{ overflowX: "auto" }}>
      {displayed.length === 0 ? (
        <div style={{ textAlign: "center", padding: "32px 0", color: C.green }}>
          <CheckSquare size={32} style={{ marginBottom: 8 }} />
          <div style={{ fontWeight: 600 }}>アラートはありません（全て対応済み）</div>
        </div>
      ) : (
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
          <thead>
            <tr style={{ borderBottom: `2px solid ${C.slate100}` }}>
              {showName && <th style={{ padding: "10px 12px", textAlign: "left", color: C.slate500, fontWeight: 600, fontSize: 12 }}>研修医</th>}
              <th style={{ padding: "10px 12px", textAlign: "left", color: C.slate500, fontWeight: 600, fontSize: 12 }}>区分</th>
              <th style={{ padding: "10px 12px", textAlign: "left", color: C.slate500, fontWeight: 600, fontSize: 12 }}>項目名</th>
              <th style={{ padding: "10px 12px", textAlign: "left", color: C.slate500, fontWeight: 600, fontSize: 12 }}>状態</th>
            </tr>
          </thead>
          <tbody>
            {displayed.map((alert: any, i: number) => (
              <tr key={i} style={{ borderBottom: `1px solid ${C.slate100}`, background: i % 2 === 0 ? C.white : "#fafafa" }}>
                {showName && <td style={{ padding: "9px 12px", fontWeight: 600, color: C.slate700 }}>{alert["研修医氏名"]}</td>}
                <td style={{ padding: "9px 12px" }}>
                  <Badge
                    label={alert["区分"]}
                    color={alert["区分"] === "症候" ? C.blue : C.purple}
                    bg={alert["区分"] === "症候" ? C.blueSoft : C.purpleSoft}
                  />
                </td>
                <td style={{ padding: "9px 12px", color: C.slate700 }}>{alert["項目名"]}</td>
                <td style={{ padding: "9px 12px" }}>
                  <Badge label={`⚠ ${alert["状態"]}`} color={C.red} bg={C.redSoft} />
                </td>
              </tr>
            ))}
            {rest > 0 && (
              <tr>
                <td colSpan={showName ? 4 : 3} style={{ padding: "10px 12px", textAlign: "center", color: C.slate500, fontSize: 12 }}>
                  ...他 {rest} 件
                </td>
              </tr>
            )}
          </tbody>
        </table>
      )}
    </div>
  );
}

// ============================================================
// Rotation Timeline Component
// ============================================================
const STATUS_META: Record<string, { color: string; bg: string; border: string; icon: string }> = {
  "両方入力済": { color: C.green,  bg: C.greenSoft,  border: C.greenBorder,  icon: "✓" },
  "研修医のみ": { color: C.amber,  bg: C.amberSoft,  border: C.amberBorder,  icon: "△" },
  "指導医のみ": { color: C.amber,  bg: C.amberSoft,  border: C.amberBorder,  icon: "△" },
  "未入力":     { color: C.red,    bg: C.redSoft,    border: C.redBorder,    icon: "✗" },
};

function RotationTimeline({ rows, residentName }: { rows: any[]; residentName: string }) {
  // Count by status
  const counts = rows.reduce<Record<string, number>>((acc, r) => {
    const s = r["入力状況"] ?? "未入力";
    acc[s] = (acc[s] || 0) + 1;
    return acc;
  }, {});

  const hasGap = (counts["研修医のみ"] || 0) + (counts["指導医のみ"] || 0) + (counts["未入力"] || 0) > 0;

  const formatDate = (d: string | null) => {
    if (!d) return "—";
    // "YYYY-MM-DD" → "YYYY/MM/DD"
    return d.replace(/-/g, "/");
  };

  return (
    <Card style={{ marginBottom: 24, border: hasGap ? `1.5px solid ${C.amberBorder}` : `1.5px solid ${C.greenBorder}` }}>
      {/* Header */}
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 20, flexWrap: "wrap", gap: 12 }}>
        <SectionTitle icon={Activity} label={`${residentName}先生 — ローテーション別 入力状況`} color={C.purple} />
        <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
          {Object.entries(counts).map(([status, cnt]) => {
            const m = STATUS_META[status] ?? STATUS_META["未入力"];
            return (
              <div key={status} style={{
                display: "flex", alignItems: "center", gap: 5,
                padding: "5px 12px", borderRadius: 99,
                background: m.bg, border: `1px solid ${m.border}`,
                fontSize: 12, fontWeight: 600, color: m.color,
              }}>
                {m.icon} {status} {cnt}件
              </div>
            );
          })}
        </div>
      </div>

      {/* Summary bar */}
      {hasGap && (
        <div style={{
          padding: "10px 16px", borderRadius: 10, marginBottom: 20,
          background: C.amberSoft, border: `1px solid ${C.amberBorder}`,
          display: "flex", alignItems: "center", gap: 8, fontSize: 13,
        }}>
          <AlertTriangle size={14} color={C.amber} />
          <span style={{ color: C.amber, fontWeight: 600 }}>
            入力状況に乖離があります。指導医への入力依頼が必要なローテーションを確認してください。
          </span>
        </div>
      )}

      {/* Rotation Table */}
      <div style={{ overflowX: "auto" }}>
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
          <thead>
            <tr style={{ borderBottom: `2px solid ${C.slate100}` }}>
              {[
                { label: "診療科", align: "left" as const },
                { label: "施設名", align: "left" as const },
                { label: "研修期間", align: "left" as const },
                { label: "区分", align: "center" as const },
                { label: "指導医", align: "left" as const },
                { label: "研修医評価", align: "center" as const },
                { label: "指導医評価", align: "center" as const },
                { label: "入力状況", align: "center" as const },
              ].map(h => (
                <th key={h.label} style={{ padding: "10px 12px", textAlign: h.align, color: C.slate500, fontWeight: 600, fontSize: 12, whiteSpace: "nowrap" }}>{h.label}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {rows.map((row: any, i: number) => {
              const status = row["入力状況"] ?? "未入力";
              const m = STATUS_META[status] ?? STATUS_META["未入力"];
              const hasResidentEval = row["研修医評価あり"];
              const hasSupervisorEval = row["指導医評価あり"];
              const rowBg = status === "両方入力済" ? C.white
                : status === "未入力" ? "#fff9f9"
                : "#fffcf0";

              return (
                <tr key={i} style={{ borderBottom: `1px solid ${C.slate100}`, background: rowBg }}>
                  {/* 診療科 */}
                  <td style={{ padding: "11px 12px", fontWeight: 700, color: C.slate900 }}>
                    <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
                      <div style={{ width: 6, height: 6, borderRadius: 99, background: m.color, flexShrink: 0 }} />
                      {row["診療科名"] || "—"}
                    </div>
                  </td>
                  {/* 施設名 */}
                  <td style={{ padding: "11px 12px", color: C.slate500, fontSize: 12 }}>
                    {row["施設名"] || "—"}
                  </td>
                  {/* 研修期間 */}
                  <td style={{ padding: "11px 12px", whiteSpace: "nowrap", fontSize: 12, color: C.slate700 }}>
                    <div style={{ display: "flex", alignItems: "center", gap: 4 }}>
                      <span style={{ fontWeight: 600 }}>{formatDate(row["研修開始日"])}</span>
                      <span style={{ color: C.slate300 }}>→</span>
                      <span style={{ fontWeight: 600 }}>{formatDate(row["研修終了日"])}</span>
                    </div>
                  </td>
                  {/* 主/並行 */}
                  <td style={{ padding: "11px 12px", textAlign: "center" }}>
                    {row["主／並行"] ? (
                      <Badge
                        label={String(row["主／並行"])}
                        color={String(row["主／並行"]).includes("主") ? C.blue : C.slate500}
                        bg={String(row["主／並行"]).includes("主") ? C.blueSoft : C.slate100}
                      />
                    ) : <span style={{ color: C.slate300 }}>—</span>}
                  </td>
                  {/* 指導医 */}
                  <td style={{ padding: "11px 12px", fontSize: 12, color: C.slate700 }}>
                    {row["指導医氏名"] || "—"}
                  </td>
                  {/* 研修医評価 */}
                  <td style={{ padding: "11px 12px", textAlign: "center" }}>
                    <EvalStatusCell
                      entered={hasResidentEval}
                      score={row["研修医評価_平均点"]}
                      label="研修医"
                    />
                  </td>
                  {/* 指導医評価 */}
                  <td style={{ padding: "11px 12px", textAlign: "center" }}>
                    <EvalStatusCell
                      entered={hasSupervisorEval}
                      score={row["指導医評価_平均点"]}
                      label="指導医"
                    />
                  </td>
                  {/* 入力状況バッジ */}
                  <td style={{ padding: "11px 12px", textAlign: "center" }}>
                    <span style={{
                      display: "inline-flex", alignItems: "center", gap: 4,
                      padding: "4px 12px", borderRadius: 99, fontSize: 12, fontWeight: 700,
                      color: m.color, background: m.bg, border: `1px solid ${m.border}`,
                      whiteSpace: "nowrap",
                    }}>
                      {m.icon} {status}
                    </span>
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    </Card>
  );
}

function EvalStatusCell({ entered, score, label }: { entered: boolean; score: number | null; label: string }) {
  if (!entered) {
    return (
      <div style={{ display: "flex", flexDirection: "column", alignItems: "center", gap: 2 }}>
        <span style={{
          display: "inline-flex", alignItems: "center", justifyContent: "center",
          width: 22, height: 22, borderRadius: 99, background: C.redSoft,
          fontSize: 12, fontWeight: 700, color: C.red,
        }}>✗</span>
        <span style={{ fontSize: 10, color: C.red }}>未入力</span>
      </div>
    );
  }
  const s = score ?? null;
  const color = s == null ? C.slate500 : s === 0 ? C.red : s < 2 ? C.amber : s >= 3.5 ? C.green : C.slate700;
  const bg = s == null ? C.slate100 : s === 0 ? C.redSoft : s < 2 ? C.amberSoft : s >= 3.5 ? C.greenSoft : C.slate100;
  return (
    <div style={{ display: "flex", flexDirection: "column", alignItems: "center", gap: 2 }}>
      <span style={{
        display: "inline-flex", alignItems: "center", justifyContent: "center",
        width: 22, height: 22, borderRadius: 99, background: C.greenSoft,
        fontSize: 12, fontWeight: 700, color: C.green,
      }}>✓</span>
      {s != null && (
        <span style={{ fontSize: 11, fontWeight: 700, color, background: bg, padding: "1px 6px", borderRadius: 99 }}>
          {s.toFixed(1)}
        </span>
      )}
    </div>
  );
}

// ============================================================
// Evaluation Dashboard
// ============================================================
function EvaluationDashboard({ data }: { data: any }) {
  const [selected, setSelected] = React.useState("all");

  const names = Array.from(new Set(
    data.residents.map((r: any) => r["研修医氏名"])
  )).filter(Boolean).sort() as string[];

  const selectedResident = selected !== "all"
    ? data.residents.find((r: any) => r["研修医氏名"] === selected)
    : null;

  const stats = React.useMemo(() => {
    if (selected === "all") {
      return {
        label: `${data.stats.total_residents} 名`,
        score: data.stats.avg_overall_score?.toFixed(2) ?? "—",
        low: data.stats.total_low_evals,
      };
    }
    const r = selectedResident;
    if (!r) return { label: selected, score: "—", low: 0 };
    return {
      label: selected,
      score: r["平均点_全体"]?.toFixed(2) ?? "—",
      low: r["2点未満件数"] || 0,
    };
  }, [selected, data, selectedResident]);

  const [itemTab, setItemTab] = React.useState<"A" | "B" | "C">("A");

  // Radar data for individual resident (competency group level)
  const radarData = React.useMemo(() => {
    if (!selectedResident || !data.radar) return null;
    const radarRows = data.radar.filter((r: any) => r["研修医氏名"] === selected);
    const competencies = [
      { key: "A_プロフェッショナリズム", label: "A: プロフェッショナリズム" },
      { key: "B_資質・能力", label: "B: 資質・能力" },
      { key: "C_基本的診療業務", label: "C: 基本的診療業務" },
    ];
    return competencies.map(({ key, label }) => {
      const resident = radarRows.find((r: any) => r["評価群"] === key && r["評価元"] === "研修医評価");
      const supervisor = radarRows.find((r: any) => r["評価群"] === key && r["評価元"] === "指導医評価");
      return {
        subject: label,
        研修医: resident?.["平均点"] ?? 0,
        指導医: supervisor?.["平均点"] ?? 0,
      };
    });
  }, [selected, data.radar, selectedResident]);

  // Item-level radar data per competency group tab
  const COMPETENCY_GROUP_MAP: Record<"A" | "B" | "C", string> = {
    A: "A_プロフェッショナリズム",
    B: "B_資質・能力",
    C: "C_基本的診療業務",
  };
  const itemRadarData = React.useMemo(() => {
    if (!selectedResident || !data.item_scores) return null;
    const groupKey = COMPETENCY_GROUP_MAP[itemTab];
    const rows = (data.item_scores as any[]).filter(
      (r) => r["研修医氏名"] === selected && r["評価群"] === groupKey && (r["平均点"] ?? 0) > 0
    );
    const items = Array.from(new Set(rows.map((r) => r["評価項目"] as string))).sort();
    if (items.length === 0) return null;
    return items.map((item) => {
      const res = rows.find((r) => r["評価項目"] === item && r["評価元"] === "研修医評価");
      const sup = rows.find((r) => r["評価項目"] === item && r["評価元"] === "指導医評価");
      return {
        subject: item,
        研修医: res?.["平均点"] ?? null,
        指導医: sup?.["平均点"] ?? null,
      };
    });
  }, [selected, data.item_scores, itemTab, selectedResident]);

  // Per-resident supervisor/resident scores from radar data (for all-view table)
  const residentScoreMap = React.useMemo(() => {
    if (!data.radar) return {} as Record<string, { sup: Record<string, number>; res: Record<string, number> }>;
    const map: Record<string, { sup: Record<string, number>; res: Record<string, number> }> = {};
    for (const r of data.radar) {
      const name = r["研修医氏名"];
      if (!name) continue;
      if (!map[name]) map[name] = { sup: {}, res: {} };
      const score = r["平均点"] ?? 0;
      const g = r["評価群"] === "A_プロフェッショナリズム" ? "A"
              : r["評価群"] === "B_資質・能力" ? "B"
              : r["評価群"] === "C_基本的診療業務" ? "C" : null;
      if (!g) continue;
      if (r["評価元"] === "指導医評価") map[name].sup[g] = score;
      else if (r["評価元"] === "研修医評価") map[name].res[g] = score;
    }
    return map;
  }, [data.radar]);

  // Alert rows for selected — 指導医評価のみ対象
  const filteredAlerts = React.useMemo(() =>
    data.alerts.filter((a: any) =>
      a["評価元"] === "指導医評価" &&
      (selected === "all" ? true : a["研修医氏名"] === selected)
    ),
  [data.alerts, selected]);

  // Rotation rows for selected resident
  const rotationRows = React.useMemo(() => {
    if (!data.rotations) return [];
    if (selected === "all") return data.rotations;
    return data.rotations.filter((r: any) => r["研修医氏名"] === selected);
  }, [data.rotations, selected]);

  // Per-resident rotation summary for all-view table
  const rotationSummaryByResident = React.useMemo(() => {
    if (!data.rotations) return {};
    const map: Record<string, { total: number; both: number; gap: number; none: number }> = {};
    for (const r of data.rotations) {
      const name = r["研修医氏名"];
      if (!map[name]) map[name] = { total: 0, both: 0, gap: 0, none: 0 };
      map[name].total++;
      const s = r["入力状況"];
      if (s === "両方入力済") map[name].both++;
      else if (s === "研修医のみ" || s === "指導医のみ") map[name].gap++;
      else map[name].none++;
    }
    return map;
  }, [data.rotations]);

  // Month-based rotation matrix for all-view cross table
  const rotationMonthMatrix = React.useMemo(() => {
    if (!data.rotations) return { months: [], residentNames: [], matrix: {} as Record<string, Record<string, any[]>> };

    // 月末7日以内の開始日 → 翌月扱い（例: 3/28開始 → 4月から）
    const normalizeStartMonth = (dateStr: string): string => {
      const d = new Date(dateStr);
      if (isNaN(d.getTime())) return (dateStr || "").substring(0, 7);
      const daysInMonth = new Date(d.getFullYear(), d.getMonth() + 1, 0).getDate();
      if (d.getDate() >= daysInMonth - 6) {
        const next = new Date(d.getFullYear(), d.getMonth() + 1, 1);
        return `${next.getFullYear()}-${String(next.getMonth() + 1).padStart(2, "0")}`;
      }
      return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
    };

    // 月初7日以内の終了日 → 前月扱い（例: 6/3終了 → 5月まで）
    const normalizeEndMonth = (dateStr: string): string => {
      const d = new Date(dateStr);
      if (isNaN(d.getTime())) return (dateStr || "").substring(0, 7);
      if (d.getDate() <= 7) {
        const prev = new Date(d.getFullYear(), d.getMonth(), 0);
        return `${prev.getFullYear()}-${String(prev.getMonth() + 1).padStart(2, "0")}`;
      }
      return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
    };

    // 正規化済み開始月〜終了月の全月リストを返す
    const expandMonths = (start: string, end: string): string[] => {
      const startM = normalizeStartMonth(start);
      const endM   = end ? normalizeEndMonth(end) : startM;
      if (!startM) return [];
      // 正規化の結果 endM < startM になる場合（超短期ローテーション）は startM のみ
      const result: string[] = [];
      const cur  = new Date(`${startM}-01`);
      const last = new Date(`${(endM >= startM ? endM : startM)}-01`);
      while (cur <= last) {
        result.push(`${cur.getFullYear()}-${String(cur.getMonth() + 1).padStart(2, "0")}`);
        cur.setMonth(cur.getMonth() + 1);
      }
      return result;
    };

    // ① 事前クラスタリング：(研修医, 診療科, 正規化開始月) を同一ローテーションとみなして統合
    //    終了日はデータ入力ずれによる過剰展開を防ぐため最小値を採用
    const clusterMap = new Map<string, any>();
    for (const r of data.rotations) {
      const startMonth = normalizeStartMonth(r["研修開始日"] || "");
      const key = `${r["研修医氏名"]}||${r["診療科名"]}||${startMonth}`;
      if (!clusterMap.has(key)) {
        clusterMap.set(key, { ...r });
      } else {
        const rep = clusterMap.get(key)!;
        // 評価フラグは OR で統合
        rep["研修医評価あり"] = rep["研修医評価あり"] || r["研修医評価あり"];
        rep["指導医評価あり"] = rep["指導医評価あり"] || r["指導医評価あり"];
        // 終了日は最小値（保守的）で過剰展開を防ぐ
        if (r["研修終了日"] && rep["研修終了日"] && r["研修終了日"] < rep["研修終了日"]) {
          rep["研修終了日"] = r["研修終了日"];
        }
      }
    }
    const clustered = Array.from(clusterMap.values());

    // ② 月一覧と研修医名一覧を構築
    const monthSet = new Set<string>();
    for (const r of clustered) {
      for (const m of expandMonths(r["研修開始日"] ?? "", r["研修終了日"] ?? "")) monthSet.add(m);
    }
    const months = Array.from(monthSet).sort();

    const residentNames: string[] = [];
    const seenNames = new Set<string>();
    for (const r of clustered) {
      const name = r["研修医氏名"];
      if (name && !seenNames.has(name)) { seenNames.add(name); residentNames.push(name); }
    }
    residentNames.sort();

    // ③ マトリクス構築（クラスタ済みデータを全期間月に展開）
    const matrix: Record<string, Record<string, any[]>> = {};
    for (const r of clustered) {
      const name = r["研修医氏名"];
      if (!name) continue;
      for (const month of expandMonths(r["研修開始日"] ?? "", r["研修終了日"] ?? "")) {
        if (!matrix[name]) matrix[name] = {};
        if (!matrix[name][month]) matrix[name][month] = [];
        if (!matrix[name][month].find((x: any) => x["診療科名"] === r["診療科名"])) {
          matrix[name][month].push(r);
        }
      }
    }
    return { months, residentNames, matrix };
  }, [data.rotations]);

  const scoreColor = (s: number) => s === 0 ? C.red : s < 2 ? C.amber : s >= 3.5 ? C.green : C.slate700;

  return (
    <div>
      <div style={{ marginBottom: 24, display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 12 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          <div style={{ width: 40, height: 40, background: C.purpleSoft, borderRadius: 10, display: "flex", alignItems: "center", justifyContent: "center" }}>
            <FileText size={22} color={C.purple} />
          </div>
          <div>
            <h2 style={{ margin: 0, fontSize: 18, fontWeight: 700, color: C.slate900 }}>研修評価 ダッシュボード</h2>
            <div style={{ fontSize: 12, color: C.slate500 }}>研修医・指導医評価の集計・傾向分析</div>
          </div>
        </div>
        <ResidentSelector names={names} selected={selected} onChange={setSelected} accentColor={C.purple} />
      </div>

      {/* Stat Cards */}
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(200px, 1fr))", gap: 16, marginBottom: 28 }}>
        <StatCard title={selected === "all" ? "評価対象研修医" : "選択中研修医"} value={stats.label} icon={Users} color={C.purple} bg={C.purpleSoft} />
        <StatCard
          title="平均評価点"
          value={stats.score}
          sub={parseFloat(stats.score) >= 3 ? "良好" : parseFloat(stats.score) >= 2 ? "標準" : "要注意"}
          icon={Award} color={parseFloat(stats.score) >= 3 ? C.green : C.amber} bg={C.greenSoft}
        />
        <StatCard
          title="低評価件数 (<2点)"
          value={stats.low}
          sub={stats.low > 0 ? "個別指導を検討" : "低評価なし"}
          icon={AlertTriangle} color={stats.low > 0 ? C.red : C.green} bg={stats.low > 0 ? C.redSoft : C.greenSoft}
        />
      </div>

      {/* Individual Resident Radar Chart */}
      {selectedResident && radarData && (
        <Card style={{ marginBottom: 24, border: `1.5px solid ${C.purpleBorder}` }}>
          <div style={{ display: "flex", alignItems: "center", gap: 12, marginBottom: 20 }}>
            <div style={{ width: 48, height: 48, borderRadius: 99, background: C.purple, display: "flex", alignItems: "center", justifyContent: "center" }}>
              <span style={{ color: "white", fontWeight: 700, fontSize: 18 }}>
                {(selectedResident["研修医氏名"] || "?")[0]}
              </span>
            </div>
            <div>
              <div style={{ fontWeight: 700, fontSize: 16, color: C.slate900 }}>{selectedResident["研修医氏名"]}</div>
              <div style={{ fontSize: 12, color: C.slate500 }}>コンピテンシー別レーダーチャート（研修医自己評価 vs 指導医評価）</div>
            </div>
          </div>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 24, alignItems: "center" }}>
            {/* Radar */}
            <div style={{ height: 280 }}>
              <ResponsiveContainer width="100%" height="100%">
                <RadarChart data={radarData} margin={{ top: 16, right: 32, bottom: 16, left: 32 }}>
                  <PolarGrid stroke="#e2e8f0" />
                  <PolarAngleAxis dataKey="subject" tick={{ fontSize: 11, fill: C.slate700, fontWeight: 500 }} />
                  <PolarRadiusAxis angle={90} domain={[0, 5]} tick={{ fontSize: 10, fill: C.slate500 }} tickCount={6} />
                  <Radar name="研修医自己評価" dataKey="研修医" stroke={C.blue} fill={C.blue} fillOpacity={0.2} strokeWidth={2} dot={{ r: 4, fill: C.blue }} />
                  <Radar name="指導医評価" dataKey="指導医" stroke={C.purple} fill={C.purple} fillOpacity={0.15} strokeWidth={2} dot={{ r: 4, fill: C.purple }} />
                  <Legend wrapperStyle={{ fontSize: 12 }} />
                  <Tooltip content={<CustomTooltip />} />
                </RadarChart>
              </ResponsiveContainer>
            </div>
            {/* Score Table */}
            <div>
              <div style={{ fontSize: 13, fontWeight: 700, color: C.slate700, marginBottom: 10 }}>コンピテンシー別スコア（研修医 vs 指導医）</div>
              <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
                <thead>
                  <tr style={{ borderBottom: `2px solid ${C.slate100}` }}>
                    {["", "研修医", "指導医", "差"].map(h => (
                      <th key={h} style={{ padding: "6px 8px", textAlign: h === "" ? "left" : "right", color: C.slate500, fontWeight: 600, whiteSpace: "nowrap" }}>{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {radarData.map((row: any, i: number) => {
                    const diff = ((row["指導医"] || 0) - (row["研修医"] || 0));
                    const diffColor = Math.abs(diff) >= 1 ? C.red : diff > 0.3 ? C.green : C.slate500;
                    return (
                      <tr key={i} style={{ borderBottom: `1px solid ${C.slate100}` }}>
                        <td style={{ padding: "8px", fontSize: 11, color: C.slate700, maxWidth: 100 }}>{row.subject.replace(/^[A-C]: /, "")}</td>
                        <td style={{ padding: "8px", textAlign: "right" }}><ScoreChip score={row["研修医"]} /></td>
                        <td style={{ padding: "8px", textAlign: "right" }}><ScoreChip score={row["指導医"]} /></td>
                        <td style={{ padding: "8px", textAlign: "right" }}>
                          <span style={{ fontSize: 12, fontWeight: 700, color: diffColor }}>
                            {diff > 0 ? "+" : ""}{diff.toFixed(2)}
                          </span>
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
              <div style={{ marginTop: 12, padding: "10px 14px", background: C.slate100, borderRadius: 10 }}>
                <div style={{ fontSize: 11, color: C.slate500, marginBottom: 2 }}>評価件数</div>
                <div style={{ fontSize: 18, fontWeight: 700, color: C.slate900 }}>{selectedResident["評価件数"] || 0} 件</div>
              </div>
            </div>
          </div>
        </Card>
      )}

      {/* Competency Detail Tabs (item-level radar) */}
      {selectedResident && data.item_scores && data.item_scores.length > 0 && (
        <Card style={{ marginBottom: 24, border: `1.5px solid ${C.purpleBorder}` }}>
          <div style={{ fontWeight: 700, fontSize: 14, color: C.slate700, marginBottom: 14 }}>
            コンピテンシー詳細（評価項目別レーダーチャート）
          </div>
          {/* Tabs */}
          <div style={{ display: "flex", gap: 8, marginBottom: 20 }}>
            {(["A", "B", "C"] as const).map((tab) => {
              const labels: Record<string, string> = {
                A: "A: プロフェッショナリズム",
                B: "B: 資質・能力",
                C: "C: 基本的診療業務",
              };
              const active = itemTab === tab;
              return (
                <button
                  key={tab}
                  onClick={() => setItemTab(tab)}
                  style={{
                    padding: "6px 16px", borderRadius: 8, border: "none", cursor: "pointer",
                    fontSize: 12, fontWeight: active ? 700 : 500,
                    background: active ? C.purple : C.slate100,
                    color: active ? C.white : C.slate500,
                    transition: "all 0.15s",
                  }}
                >
                  {labels[tab]}
                </button>
              );
            })}
          </div>
          {itemRadarData ? (
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 24, alignItems: "center" }}>
              {/* Item Radar Chart */}
              <div style={{ height: Math.max(280, itemRadarData.length * 30 + 80) }}>
                <ResponsiveContainer width="100%" height="100%">
                  <RadarChart data={itemRadarData} margin={{ top: 20, right: 40, bottom: 20, left: 40 }}>
                    <PolarGrid stroke="#e2e8f0" />
                    <PolarAngleAxis dataKey="subject" tick={{ fontSize: 10, fill: C.slate700 }} />
                    <PolarRadiusAxis angle={90} domain={[0, 5]} tick={{ fontSize: 9, fill: C.slate500 }} tickCount={6} />
                    <Radar name="研修医自己評価" dataKey="研修医" stroke={C.blue} fill={C.blue} fillOpacity={0.2} strokeWidth={2} dot={{ r: 3, fill: C.blue }} connectNulls />
                    <Radar name="指導医評価" dataKey="指導医" stroke={C.purple} fill={C.purple} fillOpacity={0.15} strokeWidth={2} dot={{ r: 3, fill: C.purple }} connectNulls />
                    <Legend wrapperStyle={{ fontSize: 11 }} />
                    <Tooltip content={<CustomTooltip />} />
                  </RadarChart>
                </ResponsiveContainer>
              </div>
              {/* Item Score Table */}
              <div style={{ overflowY: "auto", maxHeight: 360 }}>
                <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 11 }}>
                  <thead>
                    <tr style={{ borderBottom: `2px solid ${C.slate100}` }}>
                      {["評価項目", "研修医", "指導医", "差"].map(h => (
                        <th key={h} style={{ padding: "5px 8px", textAlign: h === "評価項目" ? "left" : "right", color: C.slate500, fontWeight: 600, whiteSpace: "nowrap" }}>{h}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {itemRadarData.map((row: any, i: number) => {
                      const res = row["研修医"] ?? null;
                      const sup = row["指導医"] ?? null;
                      const diff = res !== null && sup !== null ? sup - res : null;
                      const diffColor = diff === null ? C.slate300 : Math.abs(diff) >= 1 ? C.red : diff > 0.3 ? C.green : C.slate500;
                      return (
                        <tr key={i} style={{ borderBottom: `1px solid ${C.slate100}` }}>
                          <td style={{ padding: "6px 8px", color: C.slate700, wordBreak: "break-all" }}>{row.subject}</td>
                          <td style={{ padding: "6px 8px", textAlign: "right" }}>
                            {res !== null ? <ScoreChip score={res} /> : <span style={{ color: C.slate300 }}>—</span>}
                          </td>
                          <td style={{ padding: "6px 8px", textAlign: "right" }}>
                            {sup !== null ? <ScoreChip score={sup} /> : <span style={{ color: C.slate300 }}>—</span>}
                          </td>
                          <td style={{ padding: "6px 8px", textAlign: "right" }}>
                            {diff !== null
                              ? <span style={{ fontSize: 11, fontWeight: 700, color: diffColor }}>{diff > 0 ? "+" : ""}{diff.toFixed(2)}</span>
                              : <span style={{ color: C.slate300 }}>—</span>}
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            </div>
          ) : (
            <div style={{ textAlign: "center", padding: "32px 0", color: C.slate500, fontSize: 13 }}>
              このコンピテンシーグループの評価データがありません
            </div>
          )}
        </Card>
      )}

      {/* Rotation Timeline — individual view */}
      {selected !== "all" && rotationRows.length > 0 && (
        <RotationTimeline rows={rotationRows} residentName={selected} />
      )}

      {selected === "all" ? (
        <>
        <Card>
          <SectionTitle icon={Users} label="研修医別 評価・入力状況サマリー" color={C.purple} />
          {/* Legend */}
          <div style={{ display: "flex", gap: 16, marginBottom: 14, flexWrap: "wrap" }}>
            {[
              { dot: C.blue, label: "研修医自己評価スコア" },
              { dot: C.purple, label: "指導医評価スコア" },
              { dot: C.green, label: "両方入力済" },
              { dot: C.amber, label: "片方のみ" },
              { dot: C.red, label: "未入力あり / 低評価あり" },
            ].map(({ dot, label }) => (
              <div key={label} style={{ display: "flex", alignItems: "center", gap: 5, fontSize: 11, color: C.slate500 }}>
                <div style={{ width: 8, height: 8, borderRadius: 99, background: dot }} />
                {label}
              </div>
            ))}
          </div>
          <div style={{ overflowX: "auto" }}>
            <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13, minWidth: 700 }}>
              <thead style={{ position: "sticky", top: 0, background: C.white }}>
                <tr style={{ borderBottom: `2px solid ${C.slate100}` }}>
                  <th style={{ padding: "10px 12px", textAlign: "left", color: C.slate500, fontWeight: 600, fontSize: 12 }}>研修医</th>
                  <th style={{ padding: "10px 12px", textAlign: "center", color: C.blue, fontWeight: 600, fontSize: 12, whiteSpace: "nowrap" }}>自己評価</th>
                  <th style={{ padding: "10px 12px", textAlign: "center", color: C.purple, fontWeight: 600, fontSize: 12, whiteSpace: "nowrap" }}>指導医評価</th>
                  <th style={{ padding: "10px 12px", textAlign: "center", color: C.slate500, fontWeight: 600, fontSize: 12 }}>A</th>
                  <th style={{ padding: "10px 12px", textAlign: "center", color: C.slate500, fontWeight: 600, fontSize: 12 }}>B</th>
                  <th style={{ padding: "10px 12px", textAlign: "center", color: C.slate500, fontWeight: 600, fontSize: 12 }}>C</th>
                  <th style={{ padding: "10px 12px", textAlign: "center", color: C.red, fontWeight: 600, fontSize: 12, whiteSpace: "nowrap" }}>低評価</th>
                  <th style={{ padding: "10px 12px", textAlign: "left", color: C.slate500, fontWeight: 600, fontSize: 12, whiteSpace: "nowrap" }}>ローテーション入力状況</th>
                </tr>
              </thead>
              <tbody>
                {[...data.residents]
                  .sort((a: any, b: any) => (b["2点未満件数"] || 0) - (a["2点未満件数"] || 0))
                  .map((r: any, i: number) => {
                    const name = r["研修医氏名"];
                    const scores = residentScoreMap[name] ?? { sup: {}, res: {} };
                    const rs = rotationSummaryByResident[name];
                    // 研修医自己評価の全体平均
                    const resAvg = scores.res.A != null || scores.res.B != null || scores.res.C != null
                      ? (([scores.res.A, scores.res.B, scores.res.C].filter(v => v != null) as number[])
                          .reduce((a, b) => a + b, 0) /
                         [scores.res.A, scores.res.B, scores.res.C].filter(v => v != null).length)
                      : null;
                    // 指導医評価の全体平均
                    const supAvg = scores.sup.A != null || scores.sup.B != null || scores.sup.C != null
                      ? (([scores.sup.A, scores.sup.B, scores.sup.C].filter(v => v != null) as number[])
                          .reduce((a, b) => a + b, 0) /
                         [scores.sup.A, scores.sup.B, scores.sup.C].filter(v => v != null).length)
                      : null;
                    const hasAlert = r["2点未満件数"] > 0;
                    const hasNone = rs && rs.none > 0;
                    const hasGap = rs && rs.gap > 0;
                    const rowBg = hasAlert ? "#fff9f9" : hasNone ? "#fffcf0" : C.white;
                    return (
                      <tr key={i}
                        style={{ borderBottom: `1px solid ${C.slate100}`, cursor: "pointer", background: rowBg }}
                        onMouseEnter={e => (e.currentTarget.style.background = C.purpleSoft)}
                        onMouseLeave={e => (e.currentTarget.style.background = rowBg)}
                        onClick={() => setSelected(name)}
                      >
                        {/* 研修医名 */}
                        <td style={{ padding: "10px 12px", fontWeight: 700, color: C.purple }}>
                          <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
                            <div style={{
                              width: 28, height: 28, borderRadius: 99, background: C.purple,
                              display: "flex", alignItems: "center", justifyContent: "center",
                              fontSize: 11, fontWeight: 700, color: C.white, flexShrink: 0,
                            }}>{(name || "?")[0]}</div>
                            {name}
                            <ChevronRight size={12} color={C.slate300} />
                          </div>
                        </td>
                        {/* 研修医自己評価 平均 */}
                        <td style={{ padding: "10px 12px", textAlign: "center" }}>
                          {resAvg != null
                            ? <span style={{ fontSize: 13, fontWeight: 700, color: C.blue }}>{resAvg.toFixed(2)}</span>
                            : <span style={{ color: C.slate300, fontSize: 12 }}>—</span>}
                        </td>
                        {/* 指導医評価 平均 */}
                        <td style={{ padding: "10px 12px", textAlign: "center" }}>
                          <ScoreChip score={supAvg} />
                        </td>
                        {/* A / B / C 指導医スコア */}
                        {(["A", "B", "C"] as const).map(g => (
                          <td key={g} style={{ padding: "10px 8px", textAlign: "center" }}>
                            {scores.sup[g] != null
                              ? <ScoreChip score={scores.sup[g]} />
                              : <span style={{ color: C.slate300, fontSize: 12 }}>—</span>}
                          </td>
                        ))}
                        {/* 低評価件数 */}
                        <td style={{ padding: "10px 12px", textAlign: "center" }}>
                          {hasAlert
                            ? <Badge label={`⚠ ${r["2点未満件数"]}`} color={C.red} bg={C.redSoft} />
                            : <span style={{ color: C.green, fontSize: 12, fontWeight: 600 }}>✓</span>}
                        </td>
                        {/* ローテーション入力状況 */}
                        <td style={{ padding: "10px 12px" }}>
                          {rs ? (
                            <div style={{ display: "flex", flexDirection: "column", gap: 4 }}>
                              {/* プログレスバー */}
                              <div style={{ display: "flex", height: 6, borderRadius: 99, overflow: "hidden", width: 120, background: C.slate100 }}>
                                <div style={{ width: `${Math.round((rs.both / rs.total) * 100)}%`, background: C.green }} />
                                <div style={{ width: `${Math.round((rs.gap / rs.total) * 100)}%`, background: C.amber }} />
                                <div style={{ width: `${Math.round((rs.none / rs.total) * 100)}%`, background: C.red }} />
                              </div>
                              <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
                                <span style={{ fontSize: 11, color: C.green, fontWeight: 600 }}>✓{rs.both}</span>
                                {rs.gap > 0 && <span style={{ fontSize: 11, color: C.amber, fontWeight: 600 }}>△{rs.gap}</span>}
                                {rs.none > 0 && <span style={{ fontSize: 11, color: C.red, fontWeight: 600 }}>✗{rs.none}</span>}
                                <span style={{ fontSize: 11, color: C.slate300 }}>/ {rs.total}件</span>
                              </div>
                            </div>
                          ) : <span style={{ color: C.slate300 }}>—</span>}
                        </td>
                      </tr>
                    );
                  })}
              </tbody>
            </table>
          </div>
          <div style={{ marginTop: 12, fontSize: 11, color: C.slate500 }}>
            ※ A/B/C は指導医評価のコンピテンシー別平均点。行をクリックすると個人詳細を表示。
          </div>
        </Card>

        {/* Rotation Month Matrix */}
        {rotationMonthMatrix.months.length > 0 && (() => {
          const downloadExcel = () => {
            const { months, residentNames, matrix } = rotationMonthMatrix;
            const header = ["研修医", ...months.map(m => m.replace("-", "/"))];
            const rows = residentNames.map(name => {
              const row: any[] = [name];
              for (const month of months) {
                const rots: any[] = (matrix[name] && matrix[name][month]) || [];
                if (rots.length === 0) { row.push(""); continue; }
                row.push(rots.map(r => {
                  const res = r["研修医評価あり"] ? "研○" : "研✗";
                  const sup = r["指導医評価あり"] ? "指○" : "指✗";
                  return `${r["診療科名"] || "—"} [${res} ${sup}]`;
                }).join(" / "));
              }
              return row;
            });
            const ws = XLSX.utils.aoa_to_sheet([header, ...rows]);
            // 列幅設定
            ws["!cols"] = [{ wch: 14 }, ...months.map(() => ({ wch: 20 }))];
            const wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, ws, "研修月別ローテーション");
            XLSX.writeFile(wb, "研修月別ローテーション入力状況.xlsx");
          };

          const downloadPDF = () => {
            const { months, residentNames, matrix } = rotationMonthMatrix;
            const thStyle = "padding:6px 8px;border:1px solid #cbd5e1;background:#f1f5f9;font-size:11px;white-space:nowrap;text-align:center;font-weight:600;color:#374151;";
            const tdStyle = "padding:5px 7px;border:1px solid #cbd5e1;font-size:10px;vertical-align:top;";
            const headerCells = [
              `<th style="${thStyle}text-align:left;">研修医</th>`,
              ...months.map(m => `<th style="${thStyle}">${m.replace("-", "/")}</th>`),
            ].join("");
            const bodyRows = residentNames.map(name => {
              const cells = months.map(month => {
                const rots: any[] = (matrix[name] && matrix[name][month]) || [];
                if (rots.length === 0) return `<td style="${tdStyle}color:#cbd5e1;">—</td>`;
                const content = rots.map(r => {
                  const resBg = r["研修医評価あり"] ? "#16a34a" : "#e2e8f0";
                  const supBg = r["指導医評価あり"] ? "#7c3aed" : "#e2e8f0";
                  const resFg = r["研修医評価あり"] ? "#fff" : "#9ca3af";
                  const supFg = r["指導医評価あり"] ? "#fff" : "#9ca3af";
                  return [
                    `<div style="margin-bottom:3px;">`,
                    `<span style="font-size:10px;font-weight:700;">${r["診療科名"] || "—"}</span>`,
                    `<span style="margin-left:4px;display:inline-block;width:14px;height:14px;border-radius:3px;background:${resBg};color:${resFg};font-size:8px;font-weight:700;text-align:center;line-height:14px;">研</span>`,
                    `<span style="margin-left:2px;display:inline-block;width:14px;height:14px;border-radius:3px;background:${supBg};color:${supFg};font-size:8px;font-weight:700;text-align:center;line-height:14px;">指</span>`,
                    `</div>`,
                  ].join("");
                }).join("");
                return `<td style="${tdStyle}">${content}</td>`;
              }).join("");
              return `<tr><td style="${tdStyle}font-weight:700;white-space:nowrap;">${name}</td>${cells}</tr>`;
            }).join("");

            const html = [
              `<!DOCTYPE html><html><head><meta charset="utf-8">`,
              `<title>研修月別ローテーション入力状況</title>`,
              `<style>`,
              `@page{size:A3 landscape;margin:12mm}`,
              `*{-webkit-print-color-adjust:exact;print-color-adjust:exact}`,
              `body{font-family:"Helvetica Neue",Arial,"Hiragino Sans","Meiryo",sans-serif;font-size:11px;margin:0}`,
              `h2{font-size:13px;margin:0 0 8px}`,
              `table{border-collapse:collapse;width:100%}`,
              `p{font-size:9px;color:#64748b;margin-top:6px}`,
              `</style></head><body>`,
              `<h2>研修月別 ローテーション・評価入力状況</h2>`,
              `<table><thead><tr>${headerCells}</tr></thead><tbody>${bodyRows}</tbody></table>`,
              `<p>研=研修医評価、指=指導医評価。色付き=入力済、グレー=未入力</p>`,
              `</body></html>`,
            ].join("");

            // 非表示 iframe に直接書き込んでから print() — ポップアップ不要・スクリプト制限なし
            const iframe = document.createElement("iframe");
            iframe.style.cssText = "position:fixed;top:-9999px;left:-9999px;width:297mm;height:210mm;border:0;";
            document.body.appendChild(iframe);
            const doc = iframe.contentDocument ?? iframe.contentWindow?.document;
            if (!doc) { document.body.removeChild(iframe); return; }
            doc.open();
            doc.write(html);
            doc.close();
            setTimeout(() => {
              iframe.contentWindow?.focus();
              iframe.contentWindow?.print();
              setTimeout(() => { try { document.body.removeChild(iframe); } catch (_) {} }, 2000);
            }, 600);
          };

          return (
          <Card style={{ marginTop: 24 }}>
            <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 16, flexWrap: "wrap", gap: 10 }}>
              <SectionTitle icon={Activity} label="研修月別 ローテーション・評価入力状況" color={C.purple} />
              <div style={{ display: "flex", gap: 8 }}>
                <button onClick={downloadExcel} style={{ display: "flex", alignItems: "center", gap: 6, padding: "7px 14px", borderRadius: 8, border: `1.5px solid ${C.greenBorder}`, background: C.greenSoft, color: C.green, fontWeight: 600, fontSize: 12, cursor: "pointer" }}>
                  <Download size={14} />Excel
                </button>
                <button onClick={downloadPDF} style={{ display: "flex", alignItems: "center", gap: 6, padding: "7px 14px", borderRadius: 8, border: `1.5px solid ${C.purpleBorder}`, background: C.purpleSoft, color: C.purple, fontWeight: 600, fontSize: 12, cursor: "pointer" }}>
                  <Download size={14} />印刷/PDF
                </button>
              </div>
            </div>
            {/* Legend */}
            <div style={{ display: "flex", gap: 12, marginBottom: 14, flexWrap: "wrap" }}>
              {[
                { bg: C.green, label: "研" },
                { bg: C.purple, label: "指" },
              ].map(({ bg, label }) => (
                <div key={label} style={{ display: "flex", alignItems: "center", gap: 5, fontSize: 11, color: C.slate500 }}>
                  <div style={{ width: 14, height: 14, borderRadius: 3, background: bg, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 9, fontWeight: 700, color: "#fff" }}>{label}</div>
                  {label === "研" ? "研修医評価 入力済" : "指導医評価 入力済"}
                </div>
              ))}
              {[
                { label: "研", isRes: true },
                { label: "指", isRes: false },
              ].map(({ label, isRes }, i) => (
                <div key={`empty-${i}`} style={{ display: "flex", alignItems: "center", gap: 5, fontSize: 11, color: C.slate500 }}>
                  <div style={{ width: 14, height: 14, borderRadius: 3, background: C.slate100, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 9, fontWeight: 700, color: C.slate500 }}>{label}</div>
                  {isRes ? "研修医評価 未入力" : "指導医評価 未入力"}
                </div>
              ))}
            </div>
            <div style={{ overflowX: "auto" }}>
              <table style={{ borderCollapse: "collapse", fontSize: 12, minWidth: 600 }}>
                <thead>
                  <tr style={{ borderBottom: `2px solid ${C.slate100}` }}>
                    <th style={{ padding: "8px 12px", textAlign: "left", color: C.slate500, fontWeight: 600, fontSize: 12, whiteSpace: "nowrap", position: "sticky", left: 0, background: C.white, zIndex: 1 }}>
                      研修医
                    </th>
                    {rotationMonthMatrix.months.map(m => (
                      <th key={m} style={{ padding: "8px 10px", textAlign: "center", color: C.slate500, fontWeight: 600, fontSize: 11, whiteSpace: "nowrap", minWidth: 80 }}>
                        {m.replace("-", "/")}
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {rotationMonthMatrix.residentNames.map((name, ri) => (
                    <tr key={ri} style={{ borderBottom: `1px solid ${C.slate100}`, cursor: "pointer" }}
                      onMouseEnter={e => (e.currentTarget.style.background = C.purpleSoft)}
                      onMouseLeave={e => (e.currentTarget.style.background = "")}
                      onClick={() => setSelected(name)}
                    >
                      <td style={{ padding: "8px 12px", fontWeight: 700, color: C.purple, whiteSpace: "nowrap", position: "sticky", left: 0, background: "inherit", zIndex: 1 }}>
                        <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
                          <div style={{ width: 22, height: 22, borderRadius: 99, background: C.purple, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 10, fontWeight: 700, color: C.white, flexShrink: 0 }}>{(name || "?")[0]}</div>
                          {name}
                        </div>
                      </td>
                      {rotationMonthMatrix.months.map(month => {
                        const rotations: any[] = (rotationMonthMatrix.matrix[name] && rotationMonthMatrix.matrix[name][month]) || [];
                        if (rotations.length === 0) {
                          return <td key={month} style={{ padding: "6px 10px", textAlign: "center" }}><span style={{ color: C.slate300, fontSize: 11 }}>—</span></td>;
                        }
                        return (
                          <td key={month} style={{ padding: "6px 8px", verticalAlign: "top" }}>
                            {rotations.map((rot, j) => {
                              const hasRes = rot["研修医評価あり"];
                              const hasSup = rot["指導医評価あり"];
                              const allOk = hasRes && hasSup;
                              const noneOk = !hasRes && !hasSup;
                              const cellBg = allOk ? C.greenSoft : noneOk ? C.redSoft : C.amberSoft;
                              const cellBorder = allOk ? C.greenBorder : noneOk ? C.redBorder : C.amberBorder;
                              return (
                                <div key={j} style={{
                                  marginBottom: j < rotations.length - 1 ? 4 : 0,
                                  padding: "4px 6px",
                                  borderRadius: 6,
                                  background: cellBg,
                                  border: `1px solid ${cellBorder}`,
                                  minWidth: 72,
                                }}>
                                  <div style={{ fontSize: 11, fontWeight: 700, color: C.slate900, whiteSpace: "nowrap", marginBottom: 3 }}>
                                    {rot["診療科名"] || "—"}
                                  </div>
                                  <div style={{ display: "flex", gap: 3 }}>
                                    <div title={hasRes ? "研修医評価入力済" : "研修医評価未入力"} style={{
                                      width: 16, height: 16, borderRadius: 3,
                                      background: hasRes ? C.green : C.slate100,
                                      display: "flex", alignItems: "center", justifyContent: "center",
                                      fontSize: 9, fontWeight: 700,
                                      color: hasRes ? "#fff" : C.slate500,
                                    }}>研</div>
                                    <div title={hasSup ? "指導医評価入力済" : "指導医評価未入力"} style={{
                                      width: 16, height: 16, borderRadius: 3,
                                      background: hasSup ? C.purple : C.slate100,
                                      display: "flex", alignItems: "center", justifyContent: "center",
                                      fontSize: 9, fontWeight: 700,
                                      color: hasSup ? "#fff" : C.slate500,
                                    }}>指</div>
                                  </div>
                                </div>
                              );
                            })}
                          </td>
                        );
                      })}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
            <div style={{ marginTop: 12, fontSize: 11, color: C.slate500 }}>
              ※ 研=研修医評価、指=指導医評価。緑=両方入力済、黄=片方のみ、赤=未入力。行をクリックすると個人詳細を表示。
            </div>
          </Card>
          );
        })()}
        </>
      ) : (
        /* Individual Alert Table */
        filteredAlerts.length > 0 && (
          <Card style={{ border: `1.5px solid ${C.redBorder}` }}>
            <SectionTitle icon={AlertCircle} label={`${selected}先生の低評価アラート（${filteredAlerts.length} 件）`} color={C.red} />
            <div style={{ overflowX: "auto" }}>
              <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
                <thead>
                  <tr style={{ borderBottom: `2px solid ${C.slate100}` }}>
                    {["診療科", "評価項目", "点数", "状態"].map(h => (
                      <th key={h} style={{ padding: "10px 12px", textAlign: "left", color: C.slate500, fontWeight: 600, fontSize: 12 }}>{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {filteredAlerts.slice(0, 50).map((alert: any, i: number) => (
                    <tr key={i} style={{ borderBottom: `1px solid ${C.slate100}`, background: i % 2 === 0 ? C.white : "#fafafa" }}>
                      <td style={{ padding: "10px 12px", color: C.slate700 }}>{alert["診療科名"] || "—"}</td>
                      <td style={{ padding: "10px 12px", color: C.slate700 }}>{alert["評価項目"]}</td>
                      <td style={{ padding: "10px 12px" }}><ScoreChip score={alert["評価点"]} /></td>
                      <td style={{ padding: "10px 12px" }}>
                        <Badge label={`⚠ ${alert["状態"]}`} color={C.red} bg={C.redSoft} />
                      </td>
                    </tr>
                  ))}
                  {filteredAlerts.length > 50 && (
                    <tr>
                      <td colSpan={4} style={{ padding: "10px 12px", textAlign: "center", color: C.slate500, fontSize: 12 }}>...他 {filteredAlerts.length - 50} 件</td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>
          </Card>
        )
      )}

      {selected !== "all" && filteredAlerts.length === 0 && (
        <Card style={{ textAlign: "center", padding: "40px 24px", border: `1.5px solid ${C.greenBorder}` }}>
          <CheckSquare size={36} color={C.green} style={{ marginBottom: 12 }} />
          <div style={{ fontWeight: 700, fontSize: 16, color: C.green }}>低評価アラートはありません</div>
          <div style={{ fontSize: 13, color: C.slate500, marginTop: 4 }}>全評価項目が基準点以上です</div>
        </Card>
      )}
    </div>
  );
}

// ============================================================
// Case Summary Dashboard
// ============================================================
function CaseSummaryDashboard({ data }: { data: any }) {
  const [selected, setSelected] = React.useState("all");
  const [itemStatsSort, setItemStatsSort] = React.useState<"desc" | "asc">("desc"); // desc=ベスト(高い順), asc=ワースト(低い順)

  const names = Array.from(new Set(
    data.residents.map((r: any) => r["研修医氏名"])
  )).filter(Boolean).sort() as string[];

  const sortedResidents = [...data.residents].sort(
    (a: any, b: any) => (b["全体進捗率"] || 0) - (a["全体進捗率"] || 0)
  );

  const filteredResidents = selected === "all"
    ? sortedResidents
    : sortedResidents.filter((r: any) => r["研修医氏名"] === selected);

  const selectedResident = selected !== "all"
    ? data.residents.find((r: any) => r["研修医氏名"] === selected)
    : null;

  const stats = React.useMemo(() => {
    if (selected === "all") {
      return {
        label: `${data.stats.total_residents} 名`,
        rate: (data.stats.avg_overall_progress * 100).toFixed(1),
        items: data.stats.total_items,
      };
    }
    const r = selectedResident;
    if (!r) return { label: selected, rate: "0.0", items: data.stats.total_items };
    return {
      label: selected,
      rate: ((r["全体進捗率"] || 0) * 100).toFixed(1),
      items: r["対象項目数"],
    };
  }, [selected, data, selectedResident]);

  const barData = filteredResidents.slice(0, 20).map((r: any) => ({
    name: r["研修医氏名"] || "不明",
    入力済: r["入力済件数"] || 0,
    未入力: r["未入力件数"] || 0,
    進捗率: Math.round((r["全体進捗率"] || 0) * 100),
  }));

  const rateColor = (rate: number) =>
    rate >= 80 ? C.green : rate >= 50 ? C.amber : C.red;

  return (
    <div>
      <div style={{ marginBottom: 24, display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 12 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          <div style={{ width: 40, height: 40, background: C.greenSoft, borderRadius: 10, display: "flex", alignItems: "center", justifyContent: "center" }}>
            <CheckSquare size={22} color={C.green} />
          </div>
          <div>
            <h2 style={{ margin: 0, fontSize: 18, fontWeight: 700, color: C.slate900 }}>病歴要約等 進捗ダッシュボード</h2>
            <div style={{ fontSize: 12, color: C.slate500 }}>病歴要約の入力・提出状況を一覧管理</div>
          </div>
        </div>
        <ResidentSelector names={names} selected={selected} onChange={setSelected} accentColor={C.green} />
      </div>

      {/* Stat Cards */}
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(200px, 1fr))", gap: 16, marginBottom: 28 }}>
        <StatCard title={selected === "all" ? "対象研修医数" : "選択中研修医"} value={stats.label} icon={Users} color={C.green} bg={C.greenSoft} />
        <StatCard
          title="平均入力進捗率"
          value={`${stats.rate}%`}
          sub={parseFloat(stats.rate) >= 80 ? "順調" : parseFloat(stats.rate) >= 50 ? "進行中" : "遅延"}
          icon={TrendingUp} color={rateColor(parseFloat(stats.rate))} bg={C.greenSoft}
        />
        <StatCard title="対象項目数" value={stats.items} icon={BookOpen} color={C.slate500} bg={C.slate100} />
      </div>

      {/* Individual Resident — 研修医個別ページ */}
      {selectedResident && (
        <>
          <Card style={{ marginBottom: 24, border: `1.5px solid ${C.greenBorder}`, background: C.greenSoft }}>
            <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", flexWrap: "wrap", gap: 16, marginBottom: 20 }}>
              <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
                <div style={{ width: 48, height: 48, borderRadius: 99, background: C.green, display: "flex", alignItems: "center", justifyContent: "center" }}>
                  <span style={{ color: "white", fontWeight: 700, fontSize: 18 }}>
                    {(selectedResident["研修医氏名"] || "?")[0]}
                  </span>
                </div>
                <div>
                  <div style={{ fontWeight: 700, fontSize: 16, color: C.slate900 }}>{selectedResident["研修医氏名"]}</div>
                  <div style={{ fontSize: 12, color: C.slate500 }}>病歴要約等 入力状況</div>
                </div>
              </div>
              <div style={{ display: "flex", alignItems: "center", gap: 16, flexWrap: "wrap" }}>
                <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                  <span style={{ fontSize: 12, color: C.slate500 }}>進捗率</span>
                  <ProgressBar pct={parseFloat(stats.rate)} color={rateColor(parseFloat(stats.rate))} height={8} />
                  <span style={{ fontSize: 14, fontWeight: 700, color: rateColor(parseFloat(stats.rate)), minWidth: 44 }}>{stats.rate}%</span>
                </div>
                <div style={{ display: "flex", gap: 12 }}>
                  <Badge label={`済 ${selectedResident["入力済件数"] || 0}`} color={C.green} bg={C.greenSoft} />
                  <Badge label={`/ 総数 ${selectedResident["対象項目数"]}`} color={C.slate700} bg={C.slate100} />
                  {(selectedResident["未入力件数"] || 0) > 0 && (
                    <Badge label={`未 ${selectedResident["未入力件数"]}`} color={C.amber} bg={C.amberSoft} />
                  )}
                </div>
              </div>
            </div>
          </Card>

          {/* 項目別 済/未 一覧表（症候・疾患でカテゴリ分け） */}
          {(() => {
            const itemDetails = (data.item_details || []) as any[];
            const residentItems = itemDetails.filter((i: any) => i["研修医氏名"] === selected);
            if (residentItems.length === 0) return null;
            const hasKubun = residentItems.some((i: any) => i["区分"]);
            const symptomItems = hasKubun ? residentItems.filter((i: any) => i["区分"] === "症候") : [];
            const diseaseItems = hasKubun ? residentItems.filter((i: any) => i["区分"] === "疾患") : [];
            const uncategorizedItems = hasKubun ? [] : residentItems;
            const renderTable = (title: string, items: any[], color: string, bg: string) => {
              if (items.length === 0) return null;
              const completed = items.filter((i: any) => i["状態"] === "済");
              return (
                <div key={title} style={{ marginBottom: 20 }}>
                  <div style={{ fontSize: 13, fontWeight: 700, color, marginBottom: 8, display: "flex", alignItems: "center", gap: 6 }}>
                    <span style={{ background: bg, padding: "4px 10px", borderRadius: 8 }}>{title}</span>
                    <span style={{ fontSize: 12, fontWeight: 500, color: C.slate500 }}>済 {completed.length} / 未 {items.length - completed.length}</span>
                  </div>
                  <div style={{ overflowX: "auto", border: `1px solid ${C.slate300}`, borderRadius: 10 }}>
                    <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
                      <thead>
                        <tr style={{ background: C.slate100 }}>
                          <th style={{ padding: "10px 12px", textAlign: "left", fontWeight: 600, color: C.slate700, fontSize: 12 }}>項目名</th>
                          <th style={{ padding: "10px 12px", textAlign: "center", fontWeight: 600, color: C.slate700, fontSize: 12, width: 80 }}>状態</th>
                        </tr>
                      </thead>
                      <tbody>
                        {items
                          .sort((a: any, b: any) => (a["状態"] === "済" ? 0 : 1) - (b["状態"] === "済" ? 0 : 1) || (a["項目名"] || "").localeCompare(b["項目名"] || ""))
                          .map((item: any, i: number) => (
                            <tr key={i} style={{ borderBottom: `1px solid ${C.slate100}`, background: item["状態"] === "済" ? C.white : "#fffbeb" }}>
                              <td style={{ padding: "10px 12px", color: C.slate900 }}>{item["項目名"]}</td>
                              <td style={{ padding: "10px 12px", textAlign: "center" }}>
                                {item["状態"] === "済" ? (
                                  <Badge label="済" color={C.green} bg={C.greenSoft} />
                                ) : (
                                  <Badge label="未" color={C.amber} bg={C.amberSoft} />
                                )}
                              </td>
                            </tr>
                          ))}
                      </tbody>
                    </table>
                  </div>
                </div>
              );
            };
            return (
              <Card style={{ marginBottom: 24, border: `1.5px solid ${C.greenBorder}` }}>
                <SectionTitle icon={ClipboardCheck} label={`${selectedResident["研修医氏名"]} — 項目別 入力状況`} color={C.green} />
                <div style={{ fontSize: 11, color: C.slate500, marginBottom: 12, display: "flex", gap: 16 }}>
                  <span><Check size={12} style={{ verticalAlign: "middle", marginRight: 4 }} />緑 = 済</span>
                  <span><X size={12} style={{ verticalAlign: "middle", marginRight: 4 }} />黄 = 未</span>
                </div>
                {hasKubun ? (
                  <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 24 }}>
                    {renderTable("症候", symptomItems, C.blue, C.blueSoft)}
                    {renderTable("疾患", diseaseItems, C.purple, C.purpleSoft)}
                  </div>
                ) : (
                  renderTable("項目一覧", uncategorizedItems, C.slate700, C.slate100)
                )}
              </Card>
            );
          })()}
        </>
      )}

      <div style={{ display: "grid", gridTemplateColumns: selected === "all" ? "1fr 1fr" : "1fr", gap: 24 }}>
        {/* Bar Chart（全員表示時のみ） */}
        {selected === "all" && (
        <Card>
          <SectionTitle icon={BarChart2} label="研修医別 入力状況" color={C.green} />
          <ResponsiveContainer width="100%" height={320}>
            <BarChart data={barData} layout="vertical" margin={{ top: 4, right: 16, left: 8, bottom: 4 }}>
              <CartesianGrid strokeDasharray="3 3" horizontal={false} stroke="#f1f5f9" />
              <XAxis type="number" tick={{ fontSize: 11, fill: C.slate500 }} />
              <YAxis dataKey="name" type="category" width={70} tick={{ fontSize: 11, fill: C.slate700 }} />
              <Tooltip content={<CustomTooltip />} />
              <Legend wrapperStyle={{ fontSize: 12 }} />
              <Bar dataKey="入力済" name="入力済" stackId="a" fill={C.green} radius={[0, 0, 0, 0]} />
              <Bar dataKey="未入力" name="未入力" stackId="a" fill="#e2e8f0" radius={[0, 4, 4, 0]} />
            </BarChart>
          </ResponsiveContainer>
        </Card>
        )}

        {/* Progress Table */}
        <Card>
          <SectionTitle icon={ClipboardCheck} label="研修医別 進捗サマリー" color={C.green} />
          <div style={{ overflowY: "auto", maxHeight: 340 }}>
            <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
              <thead style={{ position: "sticky", top: 0, background: C.white }}>
                <tr style={{ borderBottom: `2px solid ${C.slate100}` }}>
                  {["研修医", "進捗率", "入力済/計"].map(h => (
                    <th key={h} style={{ padding: "10px 12px", textAlign: h === "研修医" ? "left" : "right", color: C.slate500, fontWeight: 600, fontSize: 12 }}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {sortedResidents
                  .filter((r: any) => selected === "all" ? true : r["研修医氏名"] === selected)
                  .map((r: any, i: number) => {
                    const rate = (r["全体進捗率"] || 0) * 100;
                    return (
                      <tr key={i} style={{ borderBottom: `1px solid ${C.slate100}`, cursor: "pointer" }}
                        onMouseEnter={e => (e.currentTarget.style.background = C.greenSoft)}
                        onMouseLeave={e => (e.currentTarget.style.background = "transparent")}
                        onClick={() => setSelected(r["研修医氏名"])}
                      >
                        <td style={{ padding: "10px 12px", fontWeight: 600, color: C.green }}>
                          <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
                            {r["研修医氏名"]}
                            <ChevronRight size={12} color={C.slate300} />
                          </div>
                        </td>
                        <td style={{ padding: "10px 12px" }}>
                          <div style={{ display: "flex", alignItems: "center", gap: 8, justifyContent: "flex-end" }}>
                            <ProgressBar pct={rate} color={rateColor(rate)} height={6} />
                            <span style={{ fontSize: 12, fontWeight: 700, color: rateColor(rate), minWidth: 38, textAlign: "right" }}>{rate.toFixed(0)}%</span>
                          </div>
                        </td>
                        <td style={{ padding: "10px 12px", textAlign: "right", color: C.slate700 }}>
                          <span style={{ color: C.green, fontWeight: 600 }}>{r["入力済件数"] || 0}</span>
                          <span style={{ color: C.slate300 }}> / </span>
                          {r["対象項目数"]}
                        </td>
                      </tr>
                    );
                  })}
              </tbody>
            </table>
          </div>
        </Card>

        {/* Item Stats */}
        {selected === "all" && data.items?.length > 0 && (
          <Card style={{ gridColumn: "1 / -1" }}>
            <SectionTitle icon={BookOpen} label="症候・疾患別 入力状況" color={C.green} />
            <div style={{ fontSize: 12, color: C.slate500, marginBottom: 12 }}>
              完了率ヘッダーをクリックで並べ替え（ベスト順 ⇔ ワースト順）
            </div>
            <div style={{ overflowX: "auto" }}>
              <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
                <thead>
                  <tr style={{ borderBottom: `2px solid ${C.slate100}` }}>
                    {["症候・疾患名", "入力済", "未入力"].map(h => (
                      <th key={h} style={{ padding: "10px 12px", textAlign: h === "症候・疾患名" ? "left" : "right", color: C.slate500, fontWeight: 600, fontSize: 12 }}>{h}</th>
                    ))}
                    <th
                      style={{ padding: "10px 12px", textAlign: "right", color: C.slate500, fontWeight: 600, fontSize: 12, cursor: "pointer", userSelect: "none" }}
                      onClick={() => setItemStatsSort(s => s === "desc" ? "asc" : "desc")}
                      onMouseEnter={e => (e.currentTarget.style.color = C.green)}
                      onMouseLeave={e => (e.currentTarget.style.color = C.slate500)}
                    >
                      <span style={{ display: "inline-flex", alignItems: "center", gap: 4 }}>
                        完了率
                        {itemStatsSort === "desc" ? <ChevronDown size={14} /> : <ChevronUp size={14} />}
                      </span>
                    </th>
                  </tr>
                </thead>
                <tbody>
                  {[...data.items]
                    .map((item: any) => {
                      const total = (item["済_件数"] || 0) + (item["未_件数"] || 0);
                      return { ...item, _rate: total > 0 ? (item["済_件数"] / total) * 100 : 0 };
                    })
                    .sort((a: any, b: any) => itemStatsSort === "desc" ? b._rate - a._rate : a._rate - b._rate)
                    .slice(0, 15)
                    .map((item: any, i: number) => {
                    const rate = item._rate;
                    return (
                      <tr key={i} style={{ borderBottom: `1px solid ${C.slate100}`, background: i % 2 === 0 ? C.white : "#fafafa" }}>
                        <td style={{ padding: "10px 12px", fontWeight: 500, color: C.slate700 }}>{item["項目名"]}</td>
                        <td style={{ padding: "10px 12px", textAlign: "right", color: C.green, fontWeight: 600 }}>{item["済_件数"]}</td>
                        <td style={{ padding: "10px 12px", textAlign: "right", color: C.slate500 }}>{item["未_件数"]}</td>
                        <td style={{ padding: "10px 12px" }}>
                          <div style={{ display: "flex", alignItems: "center", gap: 8, justifyContent: "flex-end" }}>
                            <ProgressBar pct={rate} color={rateColor(rate)} height={6} />
                            <span style={{ fontSize: 12, fontWeight: 700, color: rateColor(rate), minWidth: 38, textAlign: "right" }}>{rate.toFixed(0)}%</span>
                          </div>
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          </Card>
        )}
      </div>
    </div>
  );
}

// ============================================================
// Root Dashboard
// ============================================================
export default function Dashboard({ data }: { data: any }) {
  if (!data) return null;

  return (
    <div style={{
      background: "#f0f4f8",
      borderRadius: 20,
      padding: "28px 0 0",
    }}>
      {data.type === "symptom_disease" ? (
        <SymptomDiseaseDashboard data={data} />
      ) : data.type === "evaluation" ? (
        <EvaluationDashboard data={data} />
      ) : data.type === "case_summary" ? (
        <CaseSummaryDashboard data={data} />
      ) : (
        <Card style={{ border: `1.5px solid ${C.redBorder}` }}>
          <div style={{ display: "flex", alignItems: "center", gap: 8, color: C.red }}>
            <AlertCircle size={18} />
            <span style={{ fontWeight: 600 }}>不明な形式のダッシュボードデータです。</span>
          </div>
        </Card>
      )}
    </div>
  );
}
