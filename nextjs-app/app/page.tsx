import UploadForm from "../components/UploadForm";

export default function Page() {
  return (
    <div style={{ minHeight: "100vh", background: "#f0f4f8" }}>
      {/* Header */}
      <header style={{
        background: "linear-gradient(135deg, #1e3a5f 0%, #1d4ed8 60%, #2563eb 100%)",
        padding: "0 32px",
        boxShadow: "0 4px 20px rgba(30, 58, 95, 0.3)",
        position: "sticky",
        top: 0,
        zIndex: 100,
      }}>
        <div style={{ maxWidth: 1400, margin: "0 auto", display: "flex", alignItems: "center", justifyContent: "space-between", height: 64 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
            <div style={{
              width: 36, height: 36,
              background: "rgba(255,255,255,0.2)",
              borderRadius: 10,
              display: "flex", alignItems: "center", justifyContent: "center",
              fontSize: 20, fontWeight: 700, color: "white"
            }}>E</div>
            <div>
              <div style={{ color: "white", fontWeight: 700, fontSize: 16, lineHeight: 1.2 }}>EPOC ダッシュボード</div>
              <div style={{ color: "rgba(255,255,255,0.65)", fontSize: 11, fontWeight: 400 }}>臨床研修管理システム</div>
            </div>
          </div>
          <div style={{ display: "flex", alignItems: "center", gap: 24 }}>
            <div style={{ color: "rgba(255,255,255,0.8)", fontSize: 13 }}>
              研修医の学習状況をリアルタイムで把握
            </div>
          </div>
        </div>
      </header>

      {/* Main Content */}
      <main style={{ maxWidth: 1400, margin: "0 auto", padding: "32px 24px 64px" }}>
        <UploadForm />
      </main>
    </div>
  );
}
