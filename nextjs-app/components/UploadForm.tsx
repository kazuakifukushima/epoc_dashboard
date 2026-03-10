"use client";

import { useState, useRef, useCallback } from "react";
import Dashboard from "./Dashboard";
import { parseExcelToDashboard } from "../lib/excelParser";
import { Activity, FileText, CheckSquare, Upload, Download, Loader2, AlertCircle, CheckCircle2 } from "lucide-react";

type FileType = "symptom_disease" | "evaluation" | "case_summary";

const FILE_TYPES: { key: FileType; label: string; sub: string; icon: any; color: string; bg: string; border: string }[] = [
  {
    key: "symptom_disease",
    label: "経験症候・疾患",
    sub: "症例入力状況一覧",
    icon: Activity,
    color: "#1d4ed8",
    bg: "#eff6ff",
    border: "#bfdbfe",
  },
  {
    key: "evaluation",
    label: "研修医評価票",
    sub: "研修医評価 / 指導医評価",
    icon: FileText,
    color: "#7c3aed",
    bg: "#f5f3ff",
    border: "#ddd6fe",
  },
  {
    key: "case_summary",
    label: "病歴要約等",
    sub: "病歴要約入力状況一覧",
    icon: CheckSquare,
    color: "#059669",
    bg: "#f0fdf4",
    border: "#bbf7d0",
  },
];

export default function UploadForm() {
  const [fileType, setFileType] = useState<FileType>("symptom_disease");
  const [file, setFile] = useState<File | null>(null);
  const [status, setStatus] = useState<"idle" | "loading" | "success" | "error">("idle");
  const [message, setMessage] = useState<string>("");
  const [dashboardData, setDashboardData] = useState<any>(null);
  const [excelDownload, setExcelDownload] = useState<{ filename: string; base64: string } | null>(null);
  const [isDragging, setIsDragging] = useState(false);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const jsonInputRef = useRef<HTMLInputElement>(null);

  const selectedType = FILE_TYPES.find(t => t.key === fileType)!;

  const handleFileSelect = useCallback((f: File | null) => {
    if (f && !f.name.endsWith(".xlsx")) {
      setMessage(".xlsx ファイルのみ対応しています。");
      setStatus("error");
      return;
    }
    setFile(f);
    setMessage(f ? `選択中: ${f.name}` : "");
    setStatus(f ? "idle" : "idle");
  }, []);

  const handleJsonLoad = useCallback((e: React.ChangeEvent<HTMLInputElement>) => {
    const f = e.target.files?.[0];
    if (!f || !f.name.endsWith(".json")) {
      setMessage(".json ファイルを選択してください。");
      setStatus("error");
      return;
    }
    setStatus("loading");
    setMessage("JSON を読み込み中...");
    const reader = new FileReader();
    reader.onload = () => {
      try {
        const text = reader.result as string;
        const data = JSON.parse(text);
        if (!data.type || !["symptom_disease", "evaluation", "case_summary"].includes(data.type)) {
          throw new Error("有効なダッシュボードJSONではありません。");
        }
        setDashboardData(data);
        setExcelDownload(null);
        setStatus("success");
        setMessage(`読み込み完了: ${f.name}`);
      } catch (err) {
        setStatus("error");
        setMessage(err instanceof Error ? err.message : "JSONの解析に失敗しました。");
      }
    };
    reader.readAsText(f, "UTF-8");
    e.target.value = "";
  }, []);

  const handleDrop = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
    const dropped = e.dataTransfer.files[0];
    handleFileSelect(dropped || null);
  }, [handleFileSelect]);

  async function handleSubmit(e: React.FormEvent) {
    e.preventDefault();
    if (!file) {
      setMessage("ファイルを選択してください。");
      setStatus("error");
      return;
    }

    setStatus("loading");
    setMessage("Excel を解析中です...");
    setDashboardData(null);
    setExcelDownload(null);

    try {
      const result = await parseExcelToDashboard(file, fileType);
      setDashboardData(result.dashboardData);
      if (result.excelBase64 && result.filename) {
        setExcelDownload({ filename: result.filename, base64: result.excelBase64 });
      }
      setStatus("success");
      setMessage("ダッシュボードの生成が完了しました。");
    } catch (err) {
      setStatus("error");
      setMessage(err instanceof Error ? err.message : "不明なエラーが発生しました。");
    }
  }

  function handleDownload() {
    if (!excelDownload) return;
    const binary = atob(excelDownload.base64);
    const bytes = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
    const blob = new Blob([bytes], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = excelDownload.filename;
    a.click();
    window.URL.revokeObjectURL(url);
  }

  return (
    <div>
      {/* Upload Panel */}
      <div style={{
        background: "white",
        borderRadius: 20,
        padding: 32,
        boxShadow: "0 1px 3px rgba(0,0,0,0.08), 0 8px 24px rgba(0,0,0,0.06)",
        marginBottom: dashboardData ? 28 : 0,
      }}>
        <div style={{ marginBottom: 24 }}>
          <h2 style={{ margin: "0 0 4px", fontSize: 18, fontWeight: 700, color: "#0f172a" }}>
            EPOCファイルをアップロード
          </h2>
          <p style={{ margin: 0, fontSize: 13, color: "#64748b" }}>
            ファイル形式を選択してExcelファイルをアップロードすると、自動でダッシュボードを生成します。
          </p>
          <p style={{ margin: "8px 0 0", fontSize: 12, color: "#94a3b8" }}>
            静的サイトでは
            <button
              type="button"
              onClick={() => jsonInputRef.current?.click()}
              style={{
                background: "none", border: "none", color: "#059669", cursor: "pointer",
                textDecoration: "underline", fontWeight: 600, padding: 0, fontSize: 12,
              }}
            >
              JSON ファイルを読み込む
            </button>
            が利用できます（Python で事前に生成した JSON を選択）。
          </p>
          <input
            ref={jsonInputRef}
            type="file"
            accept=".json"
            onChange={handleJsonLoad}
            style={{ display: "none" }}
          />
        </div>

        {/* File Type Cards */}
        <div style={{ display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: 12, marginBottom: 28 }}>
          {FILE_TYPES.map((t) => {
            const Icon = t.icon;
            const active = fileType === t.key;
            return (
              <button
                key={t.key}
                type="button"
                onClick={() => { setFileType(t.key); setFile(null); setMessage(""); setStatus("idle"); }}
                style={{
                  padding: "16px 20px",
                  borderRadius: 14,
                  border: `2px solid ${active ? t.color : "#e2e8f0"}`,
                  background: active ? t.bg : "white",
                  cursor: "pointer",
                  textAlign: "left",
                  transition: "all 0.15s ease",
                  boxShadow: active ? `0 0 0 3px ${t.color}22` : "none",
                }}
              >
                <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 6 }}>
                  <div style={{
                    width: 32, height: 32, borderRadius: 8,
                    background: active ? t.color : "#f1f5f9",
                    display: "flex", alignItems: "center", justifyContent: "center",
                    flexShrink: 0,
                  }}>
                    <Icon size={16} color={active ? "white" : "#64748b"} />
                  </div>
                  <span style={{ fontSize: 14, fontWeight: 700, color: active ? t.color : "#374151" }}>
                    {t.label}
                  </span>
                </div>
                <div style={{ fontSize: 12, color: active ? t.color : "#94a3b8", paddingLeft: 42 }}>
                  {t.sub}
                </div>
              </button>
            );
          })}
        </div>

        {/* Drop Zone */}
        <form onSubmit={handleSubmit}>
          <div
            onClick={() => fileInputRef.current?.click()}
            onDragOver={(e) => { e.preventDefault(); setIsDragging(true); }}
            onDragLeave={() => setIsDragging(false)}
            onDrop={handleDrop}
            style={{
              border: `2px dashed ${isDragging ? selectedType.color : file ? "#16a34a" : "#cbd5e1"}`,
              borderRadius: 14,
              padding: "32px 24px",
              textAlign: "center",
              cursor: "pointer",
              background: isDragging ? selectedType.bg : file ? "#f0fdf4" : "#f8fafc",
              transition: "all 0.15s ease",
              marginBottom: 20,
            }}
          >
            <input
              ref={fileInputRef}
              type="file"
              accept=".xlsx"
              onChange={(e) => handleFileSelect(e.target.files?.[0] ?? null)}
              style={{ display: "none" }}
            />
            <div style={{
              width: 48, height: 48, borderRadius: 12,
              background: file ? "#dcfce7" : "#e2e8f0",
              display: "flex", alignItems: "center", justifyContent: "center",
              margin: "0 auto 12px",
            }}>
              {file
                ? <CheckCircle2 size={24} color="#16a34a" />
                : <Upload size={24} color="#64748b" />
              }
            </div>
            {file ? (
              <>
                <div style={{ fontWeight: 600, fontSize: 15, color: "#16a34a", marginBottom: 4 }}>{file.name}</div>
                <div style={{ fontSize: 12, color: "#64748b" }}>
                  {(file.size / 1024).toFixed(1)} KB — クリックして変更
                </div>
              </>
            ) : (
              <>
                <div style={{ fontWeight: 600, fontSize: 15, color: "#374151", marginBottom: 4 }}>
                  ここにファイルをドラッグ&ドロップ
                </div>
                <div style={{ fontSize: 13, color: "#94a3b8" }}>または クリックして選択 (.xlsxのみ)</div>
              </>
            )}
          </div>

          <div style={{ display: "flex", gap: 12, alignItems: "center" }}>
            <button
              type="submit"
              disabled={status === "loading" || !file}
              style={{
                flex: 1,
                padding: "14px 24px",
                borderRadius: 12,
                border: "none",
                background: status === "loading" ? "#93c5fd" : !file ? "#e2e8f0" : selectedType.color,
                color: !file ? "#94a3b8" : "white",
                cursor: status === "loading" || !file ? "not-allowed" : "pointer",
                fontWeight: 700,
                fontSize: 15,
                display: "flex", alignItems: "center", justifyContent: "center", gap: 8,
                transition: "all 0.15s ease",
                boxShadow: file && status !== "loading" ? `0 4px 12px ${selectedType.color}44` : "none",
              }}
            >
              {status === "loading" ? (
                <>
                  <span style={{ display: "inline-block", animation: "spin 1s linear infinite", width: 18, height: 18 }}>
                    <Loader2 size={18} />
                  </span>
                  解析中...
                </>
              ) : (
                <>
                  <Activity size={18} />
                  ダッシュボードを生成する
                </>
              )}
            </button>

            {excelDownload && (
              <button
                type="button"
                onClick={handleDownload}
                style={{
                  padding: "14px 20px",
                  borderRadius: 12,
                  border: "1.5px solid #e2e8f0",
                  background: "white",
                  color: "#374151",
                  cursor: "pointer",
                  fontWeight: 600,
                  fontSize: 14,
                  display: "flex", alignItems: "center", gap: 8,
                  whiteSpace: "nowrap",
                }}
              >
                <Download size={16} color="#16a34a" />
                Excelレポート
              </button>
            )}
          </div>

          {/* Status Message */}
          {message && (
            <div style={{
              marginTop: 14,
              padding: "10px 14px",
              borderRadius: 10,
              display: "flex", alignItems: "center", gap: 8,
              fontSize: 13,
              background: status === "error" ? "#fef2f2" : status === "success" ? "#f0fdf4" : "#f0f9ff",
              color: status === "error" ? "#b91c1c" : status === "success" ? "#15803d" : "#0369a1",
              border: `1px solid ${status === "error" ? "#fecaca" : status === "success" ? "#bbf7d0" : "#bae6fd"}`,
            }}>
              {status === "error" ? <AlertCircle size={14} /> : status === "success" ? <CheckCircle2 size={14} /> : <Loader2 size={14} />}
              {message}
            </div>
          )}
        </form>
      </div>

      {/* Dashboard Output */}
      {dashboardData && (
        <div>
          <Dashboard data={dashboardData} />
        </div>
      )}

      <style>{`
        @keyframes spin {
          from { transform: rotate(0deg); }
          to { transform: rotate(360deg); }
        }
      `}</style>
    </div>
  );
}
