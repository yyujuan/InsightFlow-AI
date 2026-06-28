import { useState, useEffect, useRef, useCallback } from "react";
import Papa from "papaparse";
import * as XLSX from "xlsx";

const DIFY_API_URL = "https://api.dify.ai/v1/workflows/run";
const DIFY_API_KEY = "app-uXltwnmM7wBNMN3iijKDZqSr";

const COLORS = {
  navy: "#1B3A6B",
  gold: "#C9973A",
  bg: "#F7F8FA",
  white: "#FFFFFF",
  text: "#2C3E50",
  muted: "#6B7280",
  border: "#E5E7EB",
  success: "#27AE60",
  danger: "#E74C3C",
  warning: "#F59E0B",
};

const marketColors = {
  US: "#1B3A6B",
  EU: "#C9973A",
  Asia: "#27AE60",
  Unknown: "#9CA3AF",
};

const sentimentColors = {
  Positive: "#27AE60",
  Neutral: "#F59E0B",
  Negative: "#E74C3C",
};

// ── Print styles (injected once) ────────────────────────────────────────────
const PRINT_STYLE = `
@media print {
  body { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
  .no-print { display: none !important; }
  .print-break { page-break-before: always; }
  nav { position: static !important; }
}
`;

function injectPrintStyle() {
  if (document.getElementById("insightflow-print-style")) return;
  const s = document.createElement("style");
  s.id = "insightflow-print-style";
  s.innerHTML = PRINT_STYLE;
  document.head.appendChild(s);
}

// ── Helpers ──────────────────────────────────────────────────────────────────
function getSavedCompanyContext() {
  try { return localStorage.getItem("insightflow_company_context") || ""; } catch { return ""; }
}
function saveCompanyContext(v) {
  try { localStorage.setItem("insightflow_company_context", v); } catch {}
}

// ── File parser: CSV / Excel / TXT → rows ─────────────────────────────────
function parseFile(file) {
  return new Promise((resolve, reject) => {
    const name = file.name.toLowerCase();
    if (name.endsWith(".csv")) {
      Papa.parse(file, {
        header: true,
        skipEmptyLines: true,
        complete: (r) => resolve({ rows: r.data, fields: r.meta.fields || [] }),
        error: reject,
      });
    } else if (name.endsWith(".xlsx") || name.endsWith(".xls")) {
      const reader = new FileReader();
      reader.onload = (e) => {
        const wb = XLSX.read(e.target.result, { type: "array" });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(ws, { defval: "" });
        const fields = rows.length ? Object.keys(rows[0]) : [];
        resolve({ rows, fields });
      };
      reader.onerror = reject;
      reader.readAsArrayBuffer(file);
    } else {
      // Plain text — each non-empty line is one feedback entry
      const reader = new FileReader();
      reader.onload = (e) => {
        const lines = e.target.result.split("\n").map(l => l.trim()).filter(Boolean);
        const rows = lines.map(l => ({ feedback: l }));
        resolve({ rows, fields: ["feedback"] });
      };
      reader.onerror = reject;
      reader.readAsText(file);
    }
  });
}

// Map parsed rows to Dify fields
function rowsToForm(rows, fields) {
  // Build feedback_list from likely text columns
  const feedbackCols = fields.filter(f =>
    /feedback|comment|review|content|text|内容|反馈|评论/.test(f.toLowerCase())
  );
  const feedbackCol = feedbackCols[0] || fields[0] || "feedback";
  const feedbackLines = rows.map(r => String(r[feedbackCol] || "").trim()).filter(Boolean);

  const pick = (patterns) => {
    const col = fields.find(f => patterns.some(p => f.toLowerCase().includes(p)));
    if (!col) return "";
    const vals = [...new Set(rows.map(r => String(r[col] || "").trim()).filter(Boolean))];
    return vals.join(", ");
  };

  return {
    feedback_list: feedbackLines.join("\n"),
    source_type: pick(["source", "channel", "来源", "渠道"]),
    product_module: pick(["module", "product", "feature", "模块", "产品"]),
    customer_type: pick(["customer", "type", "tier", "客户", "类型"]),
    country: pick(["country", "region", "location", "国家", "地区"]),
  };
}

// ── App root ─────────────────────────────────────────────────────────────────
export default function App() {
  const [page, setPage] = useState("home");
  const [form, setForm] = useState({
    feedback_list: "",
    source_type: "",
    company_context: getSavedCompanyContext(),
    product_module: "",
    customer_type: "",
    country: "",
  });
  const [loading, setLoading] = useState(false);
  const [result, setResult] = useState(null);
  const [error, setError] = useState("");
  const [showSettings, setShowSettings] = useState(false);

  useEffect(() => { injectPrintStyle(); }, []);

  const handleSubmit = async () => {
    if (!form.feedback_list.trim()) { setError("请上传包含反馈内容的文件"); return; }
    setError("");
    setLoading(true);
    try {
      const res = await fetch(DIFY_API_URL, {
        method: "POST",
        headers: { "Content-Type": "application/json", Authorization: `Bearer ${DIFY_API_KEY}` },
        body: JSON.stringify({
          inputs: {
            feedback_list: form.feedback_list,
            source_type: form.source_type,
            company_context: form.company_context,
            product_module: form.product_module,
            customer_type: form.customer_type,
            country: form.country,
          },
          response_mode: "blocking",
          user: "demo-user",
        }),
      });
      const data = await res.json();
      if (data.data && data.data.outputs) {
        setResult(data.data.outputs);
        setPage("result");
      } else {
        setError("分析失败，请检查 API 配置后重试");
      }
    } catch {
      setError("网络错误，请稍后重试");
    }
    setLoading(false);
  };

  // ── Mock result (demo) ────────────────────────────────────────────────────
  const mockResult = {
    aggregated_data: {
      total_count: 7,
      by_market: { US: 2, EU: 3, Asia: 2 },
      by_country: { Germany: 2, France: 1, Singapore: 2, US: 1, Japan: 1 },
      by_sentiment: { Negative: 6, Neutral: 1 },
      all_pain_points: ["导出报告时应用频繁崩溃", "GDPR 合规功能不足", "Salesforce 集成异常", "移动端性能较差", "客服响应时间过慢"],
      all_feature_requests: ["提升应用稳定性", "增加数据驻留选项", "修复 Salesforce 集成", "优化低带宽环境下的性能"],
      high_priority: [
        { summary: "导出超过500行报告时应用崩溃，严重影响日常工作流程", country: "US", market: "US", priority_score: 10, source: "support_ticket" },
        { summary: "Salesforce 集成在最新更新后失效，直接影响销售运营", country: "US", market: "US", priority_score: 9, source: "sales_call" },
        { summary: "新加坡客户对定价透明度存有疑虑，需要本地化支付方式", country: "Singapore", market: "Asia", priority_score: 9, source: "sales_call" },
      ],
    },
    common_themes: ["集成能力不足", "合规与数据隐私", "移动端体验较差", "客户支持响应慢"],
    market_patterns: [
      { market: "US", pattern: "用户高度依赖第三方工具集成，对系统稳定性要求严格", key_countries: "United States" },
      { market: "EU", pattern: "数据隐私和 GDPR 合规是核心关注点，监管合规优先于新功能", key_countries: "Germany, France" },
      { market: "Asia", pattern: "移动端优先，对本地化支付和多语言支持需求强烈", key_countries: "Singapore, Japan" },
    ],
    key_differences: [
      "美国用户更关注功能集成速度，欧洲用户优先考虑数据合规稳定性",
      "亚洲市场对移动端优先方案的需求远高于美国和欧洲市场",
      "欧洲用户对监管合规的敏感度显著高于其他市场",
    ],
    global_opportunities: ["强化第三方工具集成能力，覆盖全球主流平台", "建立完善的数据隐私合规框架，以 EU 标准为基础向全球推广", "开发移动端优化版本，重点面向亚太市场"],
    priority_by_market: [
      { market: "US", urgency: "High", top_issue: "应用稳定性和集成修复" },
      { market: "EU", urgency: "High", top_issue: "GDPR 合规和数据驻留" },
      { market: "Asia", urgency: "Medium", top_issue: "移动端性能和本地化支付" },
    ],
    quick_wins: [
      { action: "紧急修复 Salesforce 集成问题", market: "US", product_module: "CRM", impact: "直接恢复销售运营，防止客户流失", rationale: "多条高优先级反馈指向此问题" },
      { action: "优化报告导出功能，提升行数限制至5000行", market: "US", product_module: "Analytics", impact: "解决最高优先级崩溃问题", rationale: "priority_score 10分，影响核心工作流" },
      { action: "新增 GDPR 数据驻留选项", market: "EU", product_module: "Compliance", impact: "满足法务合规要求，降低法律风险", rationale: "欧洲法务团队明确需求" },
    ],
    roadmap_suggestions: [
      { feature: "本地化支付方式集成", target_market: "Asia", product_module: "Payment", priority: "High", rationale: "亚洲市场销售反馈显示价格透明度和支付方式是关键障碍" },
      { feature: "移动端低带宽优化", target_market: "Asia", product_module: "Mobile", priority: "Medium", rationale: "新加坡和日本用户均反映移动端性能问题" },
      { feature: "多时区客户支持覆盖", target_market: "Asia", product_module: "Support", priority: "Medium", rationale: "亚太区客户受时区差异影响，支持响应时间过长" },
    ],
    risk_flags: [
      { risk: "Salesforce 集成持续失效可能导致美国企业客户流失", affected_market: "US", severity: "High", suggestion: "立即组建专项修复团队，48小时内上线热修复" },
      { risk: "GDPR 合规缺失可能引发欧洲监管处罚", affected_market: "EU", severity: "High", suggestion: "优先完成数据驻留功能，聘请欧洲法律顾问进行合规审查" },
      { risk: "亚太区支持体验不佳可能影响续约率", affected_market: "Asia", severity: "Medium", suggestion: "建立亚太区本地支持团队或引入7×24小时在线支持" },
    ],
  };

  const goDemo = () => { setResult(mockResult); setPage("result"); };

  if (showSettings) {
    return (
      <SettingsPage
        value={form.company_context}
        onChange={(v) => { saveCompanyContext(v); setForm(f => ({ ...f, company_context: v })); }}
        onClose={() => setShowSettings(false)}
      />
    );
  }
  if (page === "home") return <HomePage onStart={() => setPage("analyze")} onDemo={goDemo} onSettings={() => setShowSettings(true)} />;
  if (page === "analyze") return (
    <AnalyzePage
      form={form}
      setForm={setForm}
      onSubmit={handleSubmit}
      onDemo={goDemo}
      loading={loading}
      error={error}
      onBack={() => setPage("home")}
      onSettings={() => setShowSettings(true)}
    />
  );
  if (page === "result") return (
    <ResultPage
      result={result}
      onBack={() => setPage("analyze")}
      onNew={() => { setResult(null); setPage("analyze"); }}
    />
  );
}

// ── Settings Page ────────────────────────────────────────────────────────────
function SettingsPage({ value, onChange, onClose }) {
  const [draft, setDraft] = useState(value);
  const [saved, setSaved] = useState(false);

  const handleSave = () => {
    onChange(draft);
    setSaved(true);
    setTimeout(() => setSaved(false), 2000);
  };

  return (
    <div style={{ minHeight: "100vh", background: COLORS.bg, fontFamily: "'Inter', sans-serif" }}>
      <nav style={{ background: COLORS.white, borderBottom: `1px solid ${COLORS.border}`, padding: "0 2rem", display: "flex", alignItems: "center", justifyContent: "space-between", height: 60 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          <button onClick={onClose} style={{ background: "none", border: "none", cursor: "pointer", color: COLORS.muted, fontSize: 13 }}>← 返回</button>
          <div style={{ width: 1, height: 16, background: COLORS.border }} />
          <span style={{ fontWeight: 600, color: COLORS.navy, fontSize: 15 }}>工作区设置</span>
        </div>
      </nav>
      <div style={{ maxWidth: 640, margin: "0 auto", padding: "48px 2rem" }}>
        <div style={{ marginBottom: 32 }}>
          <h2 style={{ fontSize: 22, fontWeight: 700, color: COLORS.navy, margin: "0 0 8px" }}>公司与产品背景</h2>
          <p style={{ fontSize: 14, color: COLORS.muted, margin: 0, lineHeight: 1.6 }}>
            设置后将自动填入每次分析，帮助 AI 更准确地理解你的产品上下文。保存在本地，不会上传到服务器。
          </p>
        </div>
        <div style={{ background: COLORS.white, border: `1px solid ${COLORS.border}`, borderRadius: 12, padding: "28px 24px" }}>
          <label style={{ fontSize: 13, fontWeight: 600, color: COLORS.text, display: "block", marginBottom: 8 }}>公司背景描述</label>
          <textarea
            value={draft}
            onChange={e => setDraft(e.target.value)}
            placeholder={"例如：\nB2B SaaS 项目管理平台，面向全球 500 强企业客户。\n主要模块包括：CRM、数据分析、移动端、合规管理。\n目前主要市场：美国、欧洲、东南亚。"}
            rows={8}
            style={{ width: "100%", border: `1px solid ${COLORS.border}`, borderRadius: 8, padding: "12px", fontSize: 13, color: COLORS.text, resize: "vertical", fontFamily: "inherit", boxSizing: "border-box", outline: "none", lineHeight: 1.7 }}
          />
          <div style={{ marginTop: 16, display: "flex", gap: 10, alignItems: "center" }}>
            <button
              onClick={handleSave}
              style={{ background: COLORS.navy, color: COLORS.white, border: "none", borderRadius: 8, padding: "11px 24px", fontSize: 13, fontWeight: 600, cursor: "pointer" }}
            >
              {saved ? "✓ 已保存" : "保存设置"}
            </button>
            <span style={{ fontSize: 12, color: COLORS.muted }}>保存后所有新分析将自动使用此背景</span>
          </div>
        </div>
      </div>
    </div>
  );
}

// ── Invite Modal ─────────────────────────────────────────────────────────────
function InviteModal({ onClose }) {
  const [emails, setEmails] = useState("");
  const [copied, setCopied] = useState(false);
  const inviteLink = `${window.location.origin}${window.location.pathname}?invite=team&ref=${Date.now()}`;

  const copyLink = () => {
    navigator.clipboard.writeText(inviteLink).then(() => {
      setCopied(true);
      setTimeout(() => setCopied(false), 2000);
    });
  };

  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.4)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 100 }} onClick={onClose}>
      <div style={{ background: COLORS.white, borderRadius: 16, padding: "32px", width: 480, maxWidth: "90vw", boxShadow: "0 20px 60px rgba(0,0,0,0.15)" }} onClick={e => e.stopPropagation()}>
        <div style={{ position: "relative", marginBottom: 24, textAlign: "center" }}>
          <h3 style={{ fontSize: 18, fontWeight: 700, color: COLORS.navy, margin: "0 0 6px" }}>邀请团队成员</h3>
          <p style={{ fontSize: 13, color: COLORS.muted, margin: 0 }}>分享链接，让团队成员一起查看洞察报告</p>
          <button onClick={onClose} style={{ position: "absolute", top: 0, right: 0, background: "none", border: "none", cursor: "pointer", fontSize: 20, color: COLORS.muted, lineHeight: 1 }}>×</button>
        </div>

        <div style={{ marginBottom: 20 }}>
          <label style={{ fontSize: 12, fontWeight: 600, color: COLORS.muted, display: "block", marginBottom: 8, textTransform: "uppercase", letterSpacing: 0.5 }}>邀请链接</label>
          <div style={{ display: "flex", gap: 8 }}>
            <input
              readOnly
              value={inviteLink}
              style={{ flex: 1, border: `1px solid ${COLORS.border}`, borderRadius: 8, padding: "9px 12px", fontSize: 12, color: COLORS.muted, background: COLORS.bg, outline: "none" }}
            />
            <button
              onClick={copyLink}
              style={{ background: copied ? COLORS.success : COLORS.navy, color: COLORS.white, border: "none", borderRadius: 8, padding: "9px 16px", fontSize: 12, fontWeight: 600, cursor: "pointer", whiteSpace: "nowrap" }}
            >
              {copied ? "✓ 已复制" : "复制链接"}
            </button>
          </div>
        </div>

        <div style={{ marginBottom: 20 }}>
          <label style={{ fontSize: 12, fontWeight: 600, color: COLORS.muted, display: "block", marginBottom: 8, textTransform: "uppercase", letterSpacing: 0.5 }}>通过邮件邀请</label>
          <textarea
            value={emails}
            onChange={e => setEmails(e.target.value)}
            placeholder={"colleague@company.com\nanother@company.com"}
            rows={3}
            style={{ width: "100%", border: `1px solid ${COLORS.border}`, borderRadius: 8, padding: "10px 12px", fontSize: 13, color: COLORS.text, fontFamily: "inherit", boxSizing: "border-box", outline: "none", resize: "none" }}
          />
        </div>

        <div style={{ background: "#EEF2FF", borderRadius: 8, padding: "12px 14px", marginBottom: 20 }}>
          <p style={{ fontSize: 12, color: COLORS.navy, margin: 0, lineHeight: 1.6 }}>
            💡 <strong>提示：</strong>被邀请成员将可以查看报告，但无法修改分析配置。团队协作功能完整版即将上线。
          </p>
        </div>

        <button
          onClick={onClose}
          style={{ width: "100%", background: COLORS.navy, color: COLORS.white, border: "none", borderRadius: 8, padding: "12px", fontSize: 13, fontWeight: 600, cursor: "pointer" }}
        >
          完成
        </button>
      </div>
    </div>
  );
}

// ── Animated counter ──────────────────────────────────────────────────────────
function AnimatedNumber({ target, duration = 1200, suffix = "" }) {
  const [val, setVal] = useState(0);
  const raf = useRef(null);
  useEffect(() => {
    const start = performance.now();
    const tick = (now) => {
      const p = Math.min((now - start) / duration, 1);
      const ease = 1 - Math.pow(1 - p, 3);
      setVal(Math.round(ease * target));
      if (p < 1) raf.current = requestAnimationFrame(tick);
    };
    raf.current = requestAnimationFrame(tick);
    return () => cancelAnimationFrame(raf.current);
  }, [target, duration]);
  return <span>{val}{suffix}</span>;
}

// ── Typewriter ────────────────────────────────────────────────────────────────
function Typewriter({ lines, speed = 45 }) {
  const [displayed, setDisplayed] = useState("");
  const [lineIdx, setLineIdx] = useState(0);
  const [charIdx, setCharIdx] = useState(0);
  const fullText = lines.join("");
  useEffect(() => {
    if (lineIdx >= lines.length) return;
    if (charIdx >= lines[lineIdx].length) {
      const t = setTimeout(() => { setLineIdx(i => i + 1); setCharIdx(0); }, 200);
      return () => clearTimeout(t);
    }
    const t = setTimeout(() => setCharIdx(c => c + 1), speed);
    return () => clearTimeout(t);
  }, [lineIdx, charIdx, lines, speed]);

  let rendered = [];
  let remaining = charIdx;
  for (let i = 0; i <= lineIdx && i < lines.length; i++) {
    const text = i < lineIdx ? lines[i] : lines[i].slice(0, charIdx);
    rendered.push(text);
  }
  return rendered;
}

// ── HomePage ──────────────────────────────────────────────────────────────────
function HomePage({ onStart, onDemo, onSettings }) {
  const [hoveredCard, setHoveredCard] = useState(null);
  const [visible, setVisible] = useState(false);
  const [showInvite, setShowInvite] = useState(false);

  useEffect(() => {
    const t = setTimeout(() => setVisible(true), 100);
    return () => clearTimeout(t);
  }, []);

  const titleLines = ["全球客户反馈\n", "智能洞察平台"];

  const features = [
    { icon: "🌏", title: "跨市场分析", desc: "自动识别 US / EU / Asia 市场差异，对比用户需求优先级", stat: "3", statLabel: "个市场" },
    { icon: "⚡", title: "AI 驱动洞察", desc: "基于真实反馈数据，生成可直接用于产品决策的洞察报告", stat: "60s", statLabel: "出结果" },
    { icon: "🎯", title: "产品路线图建议", desc: "按市场紧迫度排列 Quick Wins 和 Roadmap 功能建议", stat: "100%", statLabel: "AI 生成" },
  ];

  const stats = [
    { label: "分析过的反馈条数", value: 12840, suffix: "+" },
    { label: "覆盖市场", value: 15, suffix: "个" },
    { label: "平均节省时间", value: 6, suffix: "h/周" },
  ];

  return (
    <div style={{ minHeight: "100vh", background: COLORS.bg, fontFamily: "'Inter', sans-serif", overflow: "hidden" }}>
      {showInvite && <InviteModal onClose={() => setShowInvite(false)} />}

      {/* Nav */}
      <nav style={{ background: COLORS.white, borderBottom: `1px solid ${COLORS.border}`, padding: "0 2rem", display: "flex", alignItems: "center", justifyContent: "space-between", height: 60 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          <div style={{ width: 28, height: 28, background: COLORS.navy, borderRadius: 6, display: "flex", alignItems: "center", justifyContent: "center" }}>
            <span style={{ color: COLORS.gold, fontSize: 14, fontWeight: 700 }}>I</span>
          </div>
          <span style={{ fontWeight: 600, color: COLORS.navy, fontSize: 15 }}>InsightFlow AI</span>
        </div>
        <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
          <button onClick={() => setShowInvite(true)} style={{ background: "transparent", color: COLORS.muted, border: `1px solid ${COLORS.border}`, borderRadius: 6, padding: "7px 14px", fontSize: 12, cursor: "pointer" }}>👥 邀请团队</button>
          <button onClick={onSettings} style={{ background: "transparent", color: COLORS.muted, border: `1px solid ${COLORS.border}`, borderRadius: 6, padding: "7px 14px", fontSize: 12, cursor: "pointer" }}>⚙ 设置</button>
          <button onClick={onStart} style={{ background: COLORS.navy, color: COLORS.white, border: "none", borderRadius: 6, padding: "8px 18px", fontSize: 13, fontWeight: 500, cursor: "pointer" }}>开始分析</button>
        </div>
      </nav>

      {/* Hero */}
      <div style={{ maxWidth: 900, margin: "0 auto", padding: "80px 2rem 0", textAlign: "center" }}>
        {/* Badge */}
        <div style={{ display: "inline-flex", alignItems: "center", gap: 8, marginBottom: 24, opacity: visible ? 1 : 0, transform: visible ? "translateY(0)" : "translateY(12px)", transition: "all 0.5s ease" }}>
          <div style={{ width: 3, height: 14, background: COLORS.gold, borderRadius: 2 }} />
          <span style={{ fontSize: 12, color: COLORS.gold, fontWeight: 600, letterSpacing: 1, textTransform: "uppercase" }}>出海企业专属</span>
        </div>

        {/* Title */}
        <h1 style={{
          fontSize: 48, fontWeight: 800, color: COLORS.navy, lineHeight: 1.15, marginBottom: 20, margin: "0 0 20px",
          opacity: visible ? 1 : 0, transform: visible ? "translateY(0)" : "translateY(16px)",
          transition: "all 0.6s ease 0.1s",
        }}>
          全球客户反馈<br />
          <span style={{ color: COLORS.gold }}>智能洞察平台</span>
        </h1>

        {/* Subtitle */}
        <p style={{
          fontSize: 17, color: COLORS.muted, lineHeight: 1.8, maxWidth: 600, margin: "0 auto 48px",
          opacity: visible ? 1 : 0, transform: visible ? "translateY(0)" : "translateY(12px)",
          transition: "all 0.6s ease 0.2s",
        }}>
          将来自销售电话、用户评论、工单等多渠道的客户反馈，结合市场背景自动生成跨市场洞察与产品决策建议，帮助出海团队快速理解全球用户需求。
        </p>

        {/* CTAs */}
        <div style={{
          display: "flex", gap: 12, justifyContent: "center", marginBottom: 64,
          opacity: visible ? 1 : 0, transition: "opacity 0.6s ease 0.35s",
        }}>
          <button
            onClick={onStart}
            style={{ background: COLORS.navy, color: COLORS.white, border: "none", borderRadius: 8, padding: "14px 32px", fontSize: 15, fontWeight: 600, cursor: "pointer", boxShadow: "0 4px 14px rgba(27,58,107,0.25)", transition: "transform 0.15s, box-shadow 0.15s" }}
            onMouseEnter={e => { e.target.style.transform = "translateY(-2px)"; e.target.style.boxShadow = "0 8px 20px rgba(27,58,107,0.3)"; }}
            onMouseLeave={e => { e.target.style.transform = "translateY(0)"; e.target.style.boxShadow = "0 4px 14px rgba(27,58,107,0.25)"; }}
          >开始分析 →</button>
          <button
            onClick={onDemo}
            style={{ background: "transparent", color: COLORS.navy, border: `1.5px solid ${COLORS.navy}`, borderRadius: 8, padding: "14px 32px", fontSize: 15, fontWeight: 500, cursor: "pointer", transition: "background 0.15s" }}
            onMouseEnter={e => { e.target.style.background = "#EEF2FF"; }}
            onMouseLeave={e => { e.target.style.background = "transparent"; }}
          >查看 Demo</button>
        </div>

        {/* Stats bar */}
        <div style={{
          display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: 1, background: COLORS.border, borderRadius: 12, overflow: "hidden", marginBottom: 64,
          opacity: visible ? 1 : 0, transition: "opacity 0.6s ease 0.45s",
        }}>
          {stats.map((s, i) => (
            <div key={i} style={{ background: COLORS.white, padding: "20px 16px", textAlign: "center" }}>
              <div style={{ fontSize: 28, fontWeight: 800, color: COLORS.navy }}>
                {visible ? <AnimatedNumber target={s.value} suffix={s.suffix} /> : `0${s.suffix}`}
              </div>
              <div style={{ fontSize: 12, color: COLORS.muted, marginTop: 4 }}>{s.label}</div>
            </div>
          ))}
        </div>
      </div>

      {/* Feature cards */}
      <div style={{ maxWidth: 900, margin: "0 auto", padding: "0 2rem" }}>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: 16, marginBottom: 60 }}>
          {features.map((f, i) => (
            <div
              key={i}
              style={{
                background: COLORS.white,
                border: `1.5px solid ${hoveredCard === i ? COLORS.navy : COLORS.border}`,
                borderRadius: 14, padding: "28px 22px",
                cursor: "pointer",
                transition: "all 0.2s ease",
                transform: hoveredCard === i ? "translateY(-4px)" : "translateY(0)",
                boxShadow: hoveredCard === i ? "0 12px 32px rgba(27,58,107,0.12)" : "none",
                opacity: visible ? 1 : 0,
                transitionDelay: `${0.5 + i * 0.08}s`,
              }}
              onMouseEnter={() => setHoveredCard(i)}
              onMouseLeave={() => setHoveredCard(null)}
            >
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 16 }}>
                <div style={{ fontSize: 28 }}>{f.icon}</div>
                <div style={{ textAlign: "right" }}>
                  <div style={{ fontSize: 22, fontWeight: 800, color: hoveredCard === i ? COLORS.navy : COLORS.muted }}>{f.stat}</div>
                  <div style={{ fontSize: 10, color: COLORS.muted }}>{f.statLabel}</div>
                </div>
              </div>
              <div style={{ fontWeight: 700, color: COLORS.navy, fontSize: 15, marginBottom: 8 }}>{f.title}</div>
              <div style={{ fontSize: 13, color: COLORS.muted, lineHeight: 1.6 }}>{f.desc}</div>
            </div>
          ))}
        </div>

        {/* Footer strip */}
        <div style={{ borderTop: `1px solid ${COLORS.border}`, paddingTop: 32, paddingBottom: 60 }}>
          <div style={{ fontSize: 12, color: COLORS.muted, marginBottom: 16, textAlign: "center" }}>支持的反馈来源格式</div>
          <div style={{ display: "flex", gap: 10, flexWrap: "wrap", justifyContent: "center" }}>
            {["CSV 文件", "Excel (.xlsx)", "纯文本 (.txt)", "销售电话记录", "用户评论", "支持工单"].map((s, i) => (
              <span key={i} style={{ background: COLORS.bg, border: `1px solid ${COLORS.border}`, borderRadius: 20, padding: "5px 14px", fontSize: 12, color: COLORS.muted }}>{s}</span>
            ))}
          </div>
        </div>
      </div>
    </div>
  );
}

// ── AnalyzePage ───────────────────────────────────────────────────────────────
function AnalyzePage({ form, setForm, onSubmit, onDemo, loading, error, onBack, onSettings }) {
  const [dragOver, setDragOver] = useState(false);
  const [fileInfo, setFileInfo] = useState(null);
  const [parseError, setParseError] = useState("");
  const [showInvite, setShowInvite] = useState(false);
  const fileInputRef = useRef(null);

  const handleFile = useCallback(async (file) => {
    setParseError("");
    setFileInfo(null);
    try {
      const { rows, fields } = await parseFile(file);
      if (rows.length === 0) { setParseError("文件中未找到有效数据，请检查格式"); return; }
      const parsed = rowsToForm(rows, fields);
      setForm(f => ({ ...f, ...parsed }));
      setFileInfo({ name: file.name, rows: rows.length, cols: fields.length });
    } catch (e) {
      setParseError("文件解析失败：" + (e.message || "未知错误"));
    }
  }, [setForm]);

  const onDrop = useCallback(e => {
    e.preventDefault();
    setDragOver(false);
    const file = e.dataTransfer.files[0];
    if (file) handleFile(file);
  }, [handleFile]);

  const onFileChange = useCallback(e => {
    const file = e.target.files[0];
    if (file) handleFile(file);
  }, [handleFile]);

  return (
    <div style={{ minHeight: "100vh", background: COLORS.bg, fontFamily: "'Inter', sans-serif" }}>
      {showInvite && <InviteModal onClose={() => setShowInvite(false)} />}

      <nav style={{ background: COLORS.white, borderBottom: `1px solid ${COLORS.border}`, padding: "0 2rem", display: "flex", alignItems: "center", justifyContent: "space-between", height: 60 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          <button onClick={onBack} style={{ background: "none", border: "none", cursor: "pointer", color: COLORS.muted, fontSize: 13 }}>← 返回</button>
          <div style={{ width: 1, height: 16, background: COLORS.border }} />
          <span style={{ fontWeight: 600, color: COLORS.navy, fontSize: 15 }}>新建分析</span>
        </div>
        <div style={{ display: "flex", gap: 8 }}>
          <button onClick={() => setShowInvite(true)} style={{ background: "transparent", color: COLORS.muted, border: `1px solid ${COLORS.border}`, borderRadius: 6, padding: "6px 12px", fontSize: 12, cursor: "pointer" }}>👥 邀请团队</button>
          <button onClick={onSettings} style={{ background: "transparent", color: COLORS.muted, border: `1px solid ${COLORS.border}`, borderRadius: 6, padding: "6px 12px", fontSize: 12, cursor: "pointer" }}>⚙ 设置</button>
          <button onClick={onDemo} style={{ background: "transparent", color: COLORS.navy, border: `1px solid ${COLORS.border}`, borderRadius: 6, padding: "6px 14px", fontSize: 12, cursor: "pointer" }}>查看 Demo 效果</button>
        </div>
      </nav>

      <div style={{ maxWidth: 680, margin: "0 auto", padding: "40px 2rem" }}>
        <div style={{ marginBottom: 32, textAlign: "center" }}>
          <h2 style={{ fontSize: 24, fontWeight: 700, color: COLORS.navy, margin: "0 0 8px" }}>上传客户反馈</h2>
          <p style={{ fontSize: 14, color: COLORS.muted, margin: 0 }}>上传文件，AI 自动解析并生成跨市场洞察与产品建议</p>
        </div>

        {/* File Upload Zone */}
        <div
          onDragOver={e => { e.preventDefault(); setDragOver(true); }}
          onDragLeave={() => setDragOver(false)}
          onDrop={onDrop}
          onClick={() => fileInputRef.current.click()}
          style={{
            border: `2px dashed ${dragOver ? COLORS.navy : fileInfo ? COLORS.success : COLORS.border}`,
            borderRadius: 14, padding: "40px 24px", textAlign: "center", cursor: "pointer", marginBottom: 20,
            background: dragOver ? "#EEF2FF" : fileInfo ? "#F0FDF4" : COLORS.white,
            transition: "all 0.2s",
          }}
        >
          <input ref={fileInputRef} type="file" accept=".csv,.xlsx,.xls,.txt" onChange={onFileChange} style={{ display: "none" }} />
          {fileInfo ? (
            <div>
              <div style={{ fontSize: 32, marginBottom: 10 }}>✅</div>
              <div style={{ fontSize: 15, fontWeight: 600, color: COLORS.success, marginBottom: 4 }}>{fileInfo.name}</div>
              <div style={{ fontSize: 13, color: COLORS.muted }}>成功解析 <strong>{fileInfo.rows}</strong> 条反馈，共 <strong>{fileInfo.cols}</strong> 个字段</div>
              <div style={{ fontSize: 12, color: COLORS.muted, marginTop: 8 }}>点击重新上传</div>
            </div>
          ) : (
            <div>
              <div style={{ fontSize: 40, marginBottom: 12 }}>📂</div>
              <div style={{ fontSize: 15, fontWeight: 600, color: COLORS.navy, marginBottom: 6 }}>拖拽文件到此处，或点击选择</div>
              <div style={{ fontSize: 13, color: COLORS.muted }}>支持 CSV、Excel (.xlsx/.xls)、纯文本 (.txt)</div>
            </div>
          )}
        </div>

        {parseError && (
          <div style={{ background: "#FEF2F2", border: `1px solid #FECACA`, borderRadius: 8, padding: "10px 14px", fontSize: 13, color: COLORS.danger, marginBottom: 16 }}>
            {parseError}
          </div>
        )}

        {/* Parsed fields (editable) */}
        {fileInfo && (
          <div style={{ background: COLORS.white, border: `1px solid ${COLORS.border}`, borderRadius: 12, padding: "24px", marginBottom: 16 }}>
            <div style={{ fontSize: 13, fontWeight: 600, color: COLORS.navy, marginBottom: 16, display: "flex", alignItems: "center", gap: 6 }}>
              <div style={{ width: 3, height: 13, background: COLORS.gold, borderRadius: 2 }} />
              自动解析结果 — 可以手动修改
            </div>

            {[
              { key: "feedback_list", label: "客户反馈内容", multiline: true, required: true },
              { key: "source_type", label: "反馈来源" },
              { key: "product_module", label: "产品模块" },
              { key: "customer_type", label: "客户类型" },
              { key: "country", label: "来源国家" },
            ].map(f => (
              <div key={f.key} style={{ marginBottom: 16 }}>
                <label style={{ fontSize: 12, fontWeight: 500, color: COLORS.muted, display: "block", marginBottom: 5 }}>
                  {f.label}{f.required && <span style={{ color: COLORS.danger }}> *</span>}
                </label>
                {f.multiline ? (
                  <textarea
                    value={form[f.key]}
                    onChange={e => setForm(p => ({ ...p, [f.key]: e.target.value }))}
                    rows={6}
                    style={{ width: "100%", border: `1px solid ${COLORS.border}`, borderRadius: 8, padding: "10px 12px", fontSize: 13, color: COLORS.text, resize: "vertical", fontFamily: "inherit", boxSizing: "border-box", outline: "none", lineHeight: 1.6 }}
                  />
                ) : (
                  <input
                    value={form[f.key]}
                    onChange={e => setForm(p => ({ ...p, [f.key]: e.target.value }))}
                    style={{ width: "100%", border: `1px solid ${COLORS.border}`, borderRadius: 8, padding: "10px 12px", fontSize: 13, color: COLORS.text, fontFamily: "inherit", boxSizing: "border-box", outline: "none" }}
                  />
                )}
              </div>
            ))}

            {/* Company context (from settings) */}
            <div style={{ borderTop: `1px solid ${COLORS.border}`, paddingTop: 16, marginTop: 4 }}>
              <label style={{ fontSize: 12, fontWeight: 500, color: COLORS.muted, display: "block", marginBottom: 5 }}>
                公司背景（来自设置）
              </label>
              <textarea
                value={form.company_context}
                onChange={e => setForm(p => ({ ...p, company_context: e.target.value }))}
                rows={3}
                placeholder="公司和产品背景，可在「设置」中统一配置"
                style={{ width: "100%", border: `1px solid ${COLORS.border}`, borderRadius: 8, padding: "10px 12px", fontSize: 13, color: COLORS.text, resize: "vertical", fontFamily: "inherit", boxSizing: "border-box", outline: "none", lineHeight: 1.6, background: "#FAFAFA" }}
              />
            </div>
          </div>
        )}

        {error && (
          <div style={{ background: "#FEF2F2", border: `1px solid #FECACA`, borderRadius: 8, padding: "10px 14px", fontSize: 13, color: COLORS.danger, marginBottom: 16 }}>
            {error}
          </div>
        )}

        <button
          onClick={onSubmit}
          disabled={loading || !fileInfo}
          style={{
            width: "100%", background: loading || !fileInfo ? COLORS.muted : COLORS.navy,
            color: COLORS.white, border: "none", borderRadius: 8, padding: "14px", fontSize: 14,
            fontWeight: 600, cursor: loading || !fileInfo ? "not-allowed" : "pointer", transition: "background 0.2s",
          }}
        >
          {loading ? "🔍 分析中，请稍候..." : !fileInfo ? "请先上传文件" : "开始 AI 分析 →"}
        </button>

        <p style={{ fontSize: 12, color: COLORS.muted, textAlign: "center", marginTop: 12 }}>分析通常需要 30–60 秒</p>
      </div>
    </div>
  );
}

// ── ResultPage ────────────────────────────────────────────────────────────────
function ResultPage({ result, onBack, onNew }) {
  const [showInvite, setShowInvite] = useState(false);
  const printRef = useRef(null);

  const agg = result?.aggregated_data || {};
  const total = agg.total_count || 0;
  const byMarket = agg.by_market || {};
  const bySentiment = agg.by_sentiment || {};
  const highPriority = agg.high_priority || [];
  const commonThemes = result?.common_themes || [];
  const marketPatterns = result?.market_patterns || [];
  const keyDifferences = result?.key_differences || [];
  const priorityByMarket = result?.priority_by_market || [];
  const quickWins = result?.quick_wins || [];
  const roadmap = result?.roadmap_suggestions || [];
  const risks = result?.risk_flags || [];

  const handleExportPDF = () => {
    window.print();
  };

  return (
    <div ref={printRef} style={{ minHeight: "100vh", background: COLORS.bg, fontFamily: "'Inter', sans-serif" }}>
      {showInvite && <InviteModal onClose={() => setShowInvite(false)} />}

      <nav className="no-print" style={{ background: COLORS.white, borderBottom: `1px solid ${COLORS.border}`, padding: "0 2rem", display: "flex", alignItems: "center", justifyContent: "space-between", height: 60, position: "sticky", top: 0, zIndex: 10 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          <div style={{ width: 28, height: 28, background: COLORS.navy, borderRadius: 6, display: "flex", alignItems: "center", justifyContent: "center" }}>
            <span style={{ color: COLORS.gold, fontSize: 14, fontWeight: 700 }}>I</span>
          </div>
          <span style={{ fontWeight: 600, color: COLORS.navy, fontSize: 15 }}>InsightFlow AI</span>
          <div style={{ width: 1, height: 16, background: COLORS.border, margin: "0 4px" }} />
          <span style={{ fontSize: 13, color: COLORS.muted }}>分析报告</span>
        </div>
        <div style={{ display: "flex", gap: 8 }}>
          <button onClick={() => setShowInvite(true)} style={{ background: "transparent", color: COLORS.muted, border: `1px solid ${COLORS.border}`, borderRadius: 6, padding: "6px 12px", fontSize: 12, cursor: "pointer" }}>👥 邀请团队</button>
          <button
            onClick={handleExportPDF}
            style={{ background: COLORS.gold, color: COLORS.white, border: "none", borderRadius: 6, padding: "6px 14px", fontSize: 12, fontWeight: 600, cursor: "pointer" }}
          >⬇ 导出 PDF</button>
          <button onClick={onBack} style={{ background: "transparent", color: COLORS.muted, border: `1px solid ${COLORS.border}`, borderRadius: 6, padding: "6px 14px", fontSize: 12, cursor: "pointer" }}>← 返回</button>
          <button onClick={onNew} style={{ background: COLORS.navy, color: COLORS.white, border: "none", borderRadius: 6, padding: "6px 14px", fontSize: 12, cursor: "pointer" }}>新建分析</button>
        </div>
      </nav>

      <div style={{ maxWidth: 900, margin: "0 auto", padding: "32px 2rem" }}>
        {/* 概览 */}
        <div style={{ marginBottom: 10 }}><SectionLabel>概览</SectionLabel></div>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 12, marginBottom: 24 }}>
          <MetricCard label="反馈总数" value={total} />
          <MetricCard label="覆盖市场" value={Object.keys(byMarket).filter(m => m !== "Unknown").length} suffix="个" />
          <MetricCard label="高优先级" value={highPriority.length} suffix="条" color={COLORS.danger} />
          <MetricCard label="负面情感" value={bySentiment["Negative"] || 0} suffix="条" color={COLORS.warning} />
        </div>

        {/* 市场 + 情感分布 */}
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16, marginBottom: 24 }}>
          <Card title="市场分布">
            {Object.entries(byMarket).map(([market, count]) => (
              <div key={market} style={{ marginBottom: 12 }}>
                <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 4 }}>
                  <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                    <div style={{ width: 8, height: 8, borderRadius: "50%", background: marketColors[market] || COLORS.muted }} />
                    <span style={{ fontSize: 13, color: COLORS.text }}>{market}</span>
                  </div>
                  <span style={{ fontSize: 13, fontWeight: 500 }}>{count} 条</span>
                </div>
                <div style={{ height: 4, background: COLORS.bg, borderRadius: 2 }}>
                  <div style={{ height: 4, background: marketColors[market] || COLORS.muted, borderRadius: 2, width: `${(count / total) * 100}%` }} />
                </div>
              </div>
            ))}
          </Card>
          <Card title="情感分布">
            {Object.entries(bySentiment).map(([s, count]) => (
              <div key={s} style={{ marginBottom: 12 }}>
                <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 4 }}>
                  <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                    <div style={{ width: 8, height: 8, borderRadius: "50%", background: sentimentColors[s] || COLORS.muted }} />
                    <span style={{ fontSize: 13, color: COLORS.text }}>{s}</span>
                  </div>
                  <span style={{ fontSize: 13, fontWeight: 500 }}>{count} 条</span>
                </div>
                <div style={{ height: 4, background: COLORS.bg, borderRadius: 2 }}>
                  <div style={{ height: 4, background: sentimentColors[s] || COLORS.muted, borderRadius: 2, width: `${(count / total) * 100}%` }} />
                </div>
              </div>
            ))}
          </Card>
        </div>

        {/* 高优先级 */}
        {highPriority.length > 0 && (
          <div style={{ marginBottom: 24 }}>
            <SectionLabel>高优先级反馈</SectionLabel>
            <div style={{ display: "flex", flexDirection: "column", gap: 10, marginTop: 10 }}>
              {highPriority.map((item, i) => (
                <div key={i} style={{ background: COLORS.white, border: `1px solid ${COLORS.border}`, borderLeft: `3px solid ${item.priority_score >= 9 ? COLORS.danger : COLORS.warning}`, borderRadius: 8, padding: "14px 16px", display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
                  <div style={{ flex: 1 }}>
                    <p style={{ fontSize: 13, color: COLORS.text, margin: "0 0 8px", lineHeight: 1.5 }}>{item.summary}</p>
                    <div style={{ display: "flex", gap: 8 }}>
                      <Tag color={marketColors[item.market]}>{item.market}</Tag>
                      {item.source && <Tag color={COLORS.muted}>{item.source}</Tag>}
                      {item.country && item.country !== "Unknown" && <Tag color={COLORS.navy}>{item.country}</Tag>}
                    </div>
                  </div>
                  <div style={{ marginLeft: 16, textAlign: "center" }}>
                    <div style={{ fontSize: 22, fontWeight: 700, color: item.priority_score >= 9 ? COLORS.danger : COLORS.warning }}>{item.priority_score}</div>
                    <div style={{ fontSize: 10, color: COLORS.muted }}>优先级</div>
                  </div>
                </div>
              ))}
            </div>
          </div>
        )}

        {/* 市场机会 */}
        <div style={{ marginBottom: 24 }}>
          <SectionLabel>市场机会分析</SectionLabel>
          <div style={{ display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: 12, marginTop: 10 }}>
            {(Array.isArray(marketPatterns) ? marketPatterns : Object.entries(marketPatterns).map(([market, pattern]) => ({ market, pattern }))).map((mp, i) => (
              <div key={i} style={{ background: COLORS.white, border: `1px solid ${COLORS.border}`, borderRadius: 10, padding: "16px" }}>
                <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 10 }}>
                  <div style={{ width: 10, height: 10, borderRadius: "50%", background: marketColors[mp.market] || COLORS.muted }} />
                  <span style={{ fontSize: 13, fontWeight: 600, color: COLORS.navy }}>{mp.market}</span>
                  {Array.isArray(priorityByMarket) && priorityByMarket.find(p => p.market === mp.market) && (
                    <span style={{ marginLeft: "auto", fontSize: 10, fontWeight: 600, color: priorityByMarket.find(p => p.market === mp.market)?.urgency === "High" ? COLORS.danger : COLORS.warning, background: priorityByMarket.find(p => p.market === mp.market)?.urgency === "High" ? "#FEF2F2" : "#FFFBEB", padding: "2px 8px", borderRadius: 10 }}>
                      {priorityByMarket.find(p => p.market === mp.market)?.urgency}
                    </span>
                  )}
                </div>
                <p style={{ fontSize: 12, color: COLORS.muted, lineHeight: 1.6, margin: "0 0 8px" }}>{mp.pattern}</p>
                {Array.isArray(priorityByMarket) && priorityByMarket.find(p => p.market === mp.market) && (
                  <p style={{ fontSize: 11, color: COLORS.navy, margin: 0, fontWeight: 500 }}>
                    核心问题：{priorityByMarket.find(p => p.market === mp.market)?.top_issue}
                  </p>
                )}
              </div>
            ))}
          </div>
        </div>

        {/* 共同主题 & 差异 */}
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16, marginBottom: 24 }}>
          <Card title="全球共同主题">
            {commonThemes.map((t, i) => (
              <div key={i} style={{ display: "flex", alignItems: "flex-start", gap: 8, marginBottom: 8 }}>
                <div style={{ width: 16, height: 16, borderRadius: "50%", background: "#EEF2FF", color: COLORS.navy, fontSize: 10, fontWeight: 600, display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0, marginTop: 1 }}>{i + 1}</div>
                <span style={{ fontSize: 13, color: COLORS.text, lineHeight: 1.5 }}>{t}</span>
              </div>
            ))}
          </Card>
          <Card title="跨市场关键差异">
            {keyDifferences.map((d, i) => (
              <div key={i} style={{ borderLeft: `2px solid ${COLORS.gold}`, paddingLeft: 10, marginBottom: 10 }}>
                <span style={{ fontSize: 12, color: COLORS.muted, lineHeight: 1.6 }}>{d}</span>
              </div>
            ))}
          </Card>
        </div>

        {/* Quick Wins */}
        <div style={{ marginBottom: 24 }}>
          <SectionLabel>Quick Wins — 立即可做</SectionLabel>
          <div style={{ display: "flex", flexDirection: "column", gap: 10, marginTop: 10 }}>
            {quickWins.map((w, i) => (
              <div key={i} style={{ background: COLORS.white, border: `1px solid ${COLORS.border}`, borderRadius: 8, padding: "14px 16px", display: "flex", gap: 14 }}>
                <div style={{ width: 28, height: 28, borderRadius: 6, background: "#EEF2FF", color: COLORS.navy, fontSize: 12, fontWeight: 700, display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>Q{i + 1}</div>
                <div>
                  <div style={{ fontSize: 13, fontWeight: 600, color: COLORS.text, marginBottom: 4 }}>{w.action}</div>
                  <div style={{ fontSize: 12, color: COLORS.muted, marginBottom: 6 }}>{w.rationale}</div>
                  <div style={{ display: "flex", gap: 6 }}>
                    {w.market && <Tag color={marketColors[w.market] || COLORS.muted}>{w.market}</Tag>}
                    {w.product_module && <Tag color={COLORS.navy}>{w.product_module}</Tag>}
                  </div>
                </div>
                <div style={{ marginLeft: "auto", fontSize: 11, color: COLORS.success, fontWeight: 500, whiteSpace: "nowrap", flexShrink: 0 }}>{w.impact}</div>
              </div>
            ))}
          </div>
        </div>

        {/* Roadmap */}
        <div style={{ marginBottom: 24 }}>
          <SectionLabel>产品路线图建议</SectionLabel>
          <div style={{ display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: 12, marginTop: 10 }}>
            {roadmap.map((r, i) => (
              <div key={i} style={{ background: COLORS.white, border: `1px solid ${COLORS.border}`, borderRadius: 10, padding: "16px" }}>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 8 }}>
                  <span style={{ fontSize: 13, fontWeight: 600, color: COLORS.text, flex: 1 }}>{r.feature}</span>
                  <span style={{ fontSize: 10, fontWeight: 600, padding: "2px 8px", borderRadius: 10, marginLeft: 8, flexShrink: 0, background: r.priority === "High" ? "#FEF2F2" : r.priority === "Medium" ? "#FFFBEB" : "#F0FDF4", color: r.priority === "High" ? COLORS.danger : r.priority === "Medium" ? COLORS.warning : COLORS.success }}>{r.priority}</span>
                </div>
                <p style={{ fontSize: 12, color: COLORS.muted, lineHeight: 1.5, margin: "0 0 10px" }}>{r.rationale}</p>
                <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
                  {r.target_market && <Tag color={COLORS.navy}>{r.target_market}</Tag>}
                  {r.product_module && <Tag color={COLORS.muted}>{r.product_module}</Tag>}
                </div>
              </div>
            ))}
          </div>
        </div>

        {/* 风险预警 */}
        <div style={{ marginBottom: 40 }}>
          <SectionLabel>风险预警</SectionLabel>
          <div style={{ display: "flex", flexDirection: "column", gap: 10, marginTop: 10 }}>
            {risks.map((r, i) => (
              <div key={i} style={{ background: COLORS.white, border: `1px solid ${COLORS.border}`, borderLeft: `3px solid ${r.severity === "High" ? COLORS.danger : COLORS.warning}`, borderRadius: 8, padding: "14px 16px" }}>
                <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 6 }}>
                  <span style={{ fontSize: 13, fontWeight: 600, color: COLORS.text }}>{r.risk}</span>
                  <span style={{ fontSize: 10, fontWeight: 600, padding: "2px 8px", borderRadius: 10, background: r.severity === "High" ? "#FEF2F2" : "#FFFBEB", color: r.severity === "High" ? COLORS.danger : COLORS.warning }}>{r.severity}</span>
                </div>
                <p style={{ fontSize: 12, color: COLORS.muted, margin: "0 0 8px" }}>建议：{r.suggestion}</p>
                {r.affected_market && <Tag color={marketColors[r.affected_market] || COLORS.muted}>{r.affected_market}</Tag>}
              </div>
            ))}
          </div>
        </div>

        <div className="no-print" style={{ borderTop: `1px solid ${COLORS.border}`, paddingTop: 24, display: "flex", justifyContent: "center", gap: 12 }}>
          <button onClick={onNew} style={{ background: COLORS.navy, color: COLORS.white, border: "none", borderRadius: 8, padding: "11px 24px", fontSize: 13, fontWeight: 600, cursor: "pointer" }}>新建分析</button>
          <button onClick={handleExportPDF} style={{ background: COLORS.gold, color: COLORS.white, border: "none", borderRadius: 8, padding: "11px 24px", fontSize: 13, fontWeight: 600, cursor: "pointer" }}>导出 PDF</button>
          <button onClick={onBack} style={{ background: "transparent", color: COLORS.muted, border: `1px solid ${COLORS.border}`, borderRadius: 8, padding: "11px 24px", fontSize: 13, cursor: "pointer" }}>返回</button>
        </div>
      </div>
    </div>
  );
}

// ── Shared components ────────────────────────────────────────────────────────
function MetricCard({ label, value, suffix = "", color }) {
  return (
    <div style={{ background: COLORS.white, border: `1px solid ${COLORS.border}`, borderRadius: 10, padding: "16px" }}>
      <div style={{ fontSize: 12, color: COLORS.muted, marginBottom: 6 }}>{label}</div>
      <div style={{ fontSize: 26, fontWeight: 700, color: color || COLORS.navy }}>{value}<span style={{ fontSize: 13, fontWeight: 500, color: COLORS.muted, marginLeft: 3 }}>{suffix}</span></div>
    </div>
  );
}

function Card({ title, children }) {
  return (
    <div style={{ background: COLORS.white, border: `1px solid ${COLORS.border}`, borderRadius: 10, padding: "18px 16px" }}>
      <div style={{ fontSize: 13, fontWeight: 600, color: COLORS.navy, marginBottom: 14, display: "flex", alignItems: "center", gap: 6 }}>
        <div style={{ width: 3, height: 13, background: COLORS.gold, borderRadius: 2 }} />
        {title}
      </div>
      {children}
    </div>
  );
}

function SectionLabel({ children }) {
  return (
    <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 4 }}>
      <div style={{ width: 3, height: 14, background: COLORS.navy, borderRadius: 2 }} />
      <span style={{ fontSize: 13, fontWeight: 600, color: COLORS.navy }}>{children}</span>
    </div>
  );
}

function Tag({ color, children }) {
  return (
    <span style={{ fontSize: 10, fontWeight: 500, padding: "2px 8px", borderRadius: 10, background: color + "18", color: color, border: `1px solid ${color}30` }}>
      {children}
    </span>
  );
}
