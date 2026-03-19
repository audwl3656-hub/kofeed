"""
연도별 히스토리 데이터 → 자체완결 HTML 대시보드 생성.
CDN(React + Babel)을 사용하므로 수신자가 브라우저에서 열면 인터랙티브하게 동작합니다.
"""
import json


_JSX_BODY = r"""
function calcZScores(data) {
  const result = data.map(r => ({ ...r }));
  ITEMS.forEach(item => {
    const vals = data.map(r => r[item]).filter(v => v !== null && v !== undefined && !isNaN(v));
    if (vals.length < 2) return;
    const mean = vals.reduce((a, b) => a + b, 0) / vals.length;
    const std = Math.sqrt(vals.reduce((a, b) => a + (b - mean) ** 2, 0) / (vals.length - 1));
    result.forEach(r => {
      const v = r[item];
      r[`z_${item}`] = (v === null || v === undefined || isNaN(v) || std === 0)
        ? null : +(((v - mean) / std).toFixed(3));
    });
  });
  return result;
}

function zBg(z) {
  if (z === null) return "transparent";
  const c = Math.max(-3, Math.min(3, z));
  if (c < 0) return `rgba(37,99,235,${(-c / 3) * 0.18})`;
  return `rgba(220,38,38,${(c / 3) * 0.18})`;
}
function zText(z) {
  if (z === null) return "#94a3b8";
  if (z > 2) return "#b91c1c";
  if (z < -2) return "#1d4ed8";
  if (z > 1) return "#dc2626";
  if (z < -1) return "#2563eb";
  return "#374151";
}

function Sparkline({ values, color, w = 88, h = 32 }) {
  const clean = values.filter(v => v !== null && !isNaN(v));
  if (!clean || clean.length < 2) return null;
  const min = Math.min(...clean), max = Math.max(...clean);
  const range = max - min || 0.001;
  const pad = 4;
  const xs = values.map((_, i) => pad + (i / (values.length - 1)) * (w - pad * 2));
  const ys = values.map(v => v === null ? h/2 : h - pad - ((v - min) / range) * (h - pad * 2));
  const linePath = xs.map((x, i) => `${i === 0 ? "M" : "L"}${x.toFixed(1)},${ys[i].toFixed(1)}`).join(" ");
  const areaPath = `${linePath} L${xs[xs.length-1]},${h-pad} L${xs[0]},${h-pad} Z`;
  const last = clean[clean.length-1], first = clean[0];
  const trend = last - first;
  const endColor = Math.abs(trend) < 0.05 ? "#94a3b8" : trend > 0 ? "#ef4444" : "#3b82f6";
  const uid = `${color.replace("#","")}${w}${h}${values.join("")}`;
  return (
    <svg width={w} height={h} style={{ display:"block", overflow:"visible" }}>
      <defs>
        <linearGradient id={`g${uid}`} x1="0" y1="0" x2="0" y2="1">
          <stop offset="0%" stopColor={color} stopOpacity="0.28"/>
          <stop offset="100%" stopColor={color} stopOpacity="0.01"/>
        </linearGradient>
      </defs>
      <line x1={pad} x2={w-pad} y1={h/2} y2={h/2} stroke="#e5e7eb" strokeWidth="0.7" strokeDasharray="2,2"/>
      <path d={areaPath} fill={`url(#g${uid})`}/>
      <path d={linePath} fill="none" stroke={color} strokeWidth="1.8" strokeLinejoin="round" strokeLinecap="round"/>
      {xs.map((x, i) => (
        <circle key={i} cx={x} cy={ys[i]} r={i === values.length-1 ? 3.2 : 2.2}
          fill={i === values.length-1 ? endColor : color} stroke="#fff" strokeWidth="1.2"/>
      ))}
    </svg>
  );
}

function TrendBadge({ values }) {
  const clean = values.filter(v => v !== null && !isNaN(v));
  if (!clean || clean.length < 2) return null;
  const diff = clean[clean.length-1] - clean[0];
  if (Math.abs(diff) < 0.05) return <span style={{ fontSize:11, color:"#94a3b8" }}>→</span>;
  const up = diff > 0;
  return (
    <span style={{ fontSize:10, fontWeight:700, padding:"2px 6px", borderRadius:10,
      background: up ? "#fef2f2" : "#eff6ff", color: up ? "#dc2626" : "#2563eb" }}>
      {up ? "▲" : "▼"} {Math.abs(diff).toFixed(2)}
    </span>
  );
}

function buildGlobalBriefing(zData) {
  const allAlerts = [];
  zData.forEach(row => {
    ITEMS.forEach(item => {
      const z = row[`z_${item}`];
      if (z !== null && Math.abs(z) > 2) allAlerts.push({ feed: row.feed, year: row.year, item, z });
    });
  });
  const alertByFeed = {};
  allAlerts.forEach(a => { alertByFeed[a.feed] = (alertByFeed[a.feed] || 0) + 1; });
  const worstFeed = Object.entries(alertByFeed).sort((a,b) => b[1]-a[1])[0];
  const alertByItem = {};
  allAlerts.forEach(a => { alertByItem[a.item] = (alertByItem[a.item] || 0) + 1; });
  const worstItem = Object.entries(alertByItem).sort((a,b) => b[1]-a[1])[0];

  const firstYear = Math.min(...YEARS), lastYear = Math.max(...YEARS);
  const itemTrends = ITEMS.map(item => {
    const yFirst = zData.filter(r => r.year === firstYear).map(r => r[`z_${item}`]).filter(v => v !== null);
    const yLast  = zData.filter(r => r.year === lastYear).map(r => r[`z_${item}`]).filter(v => v !== null);
    if (!yFirst.length || !yLast.length) return null;
    const avg0 = yFirst.reduce((a,b) => a+b,0)/yFirst.length;
    const avg1 = yLast.reduce((a,b) => a+b,0)/yLast.length;
    return { item, diff: +(avg1 - avg0).toFixed(2) };
  }).filter(t => t && Math.abs(t.diff) > 0.15).sort((a,b) => Math.abs(b.diff) - Math.abs(a.diff));
  const risingItems  = itemTrends.filter(t => t.diff > 0).slice(0,2);
  const fallingItems = itemTrends.filter(t => t.diff < 0).slice(0,2);

  const itemVariance = ITEMS.map(item => {
    const zAvgPerFeed = FEEDS.map(feed => {
      const vals = zData.filter(r => r.feed === feed).map(r => r[`z_${item}`]).filter(v => v !== null);
      return vals.length ? vals.reduce((a,b) => a+b,0)/vals.length : 0;
    });
    const mean = zAvgPerFeed.reduce((a,b) => a+b,0)/zAvgPerFeed.length;
    return { item, variance: +(zAvgPerFeed.reduce((a,b) => a+(b-mean)**2,0)/zAvgPerFeed.length).toFixed(2) };
  }).sort((a,b) => b.variance - a.variance);
  const mostVaried = itemVariance[0];

  const line1 = allAlerts.length === 0
    ? { icon:"✅", color:"#15803d", bg:"#f0fdf4", border:"#bbf7d0",
        text: `분석 기간(${firstYear}–${lastYear}) 동안 전체 ${FEEDS.length}종 사료에서 이상치(|Z|>2)가 감지되지 않았습니다. 전반적으로 안정적인 품질 수준을 유지하고 있습니다.` }
    : { icon:"⚠️", color:"#92400e", bg:"#fffbeb", border:"#fde68a",
        text: `전체 ${FEEDS.length}종 사료 중 총 ${allAlerts.length}건의 이상치(|Z|>2)가 감지되었습니다.${worstFeed ? ` 특히 ${worstFeed[0]}에서 ${worstFeed[1]}건으로 가장 많이 발생하였으며,` : ""}${worstItem ? ` ${worstItem[0]} 항목이 ${worstItem[1]}건으로 가장 빈번하게 벗어났습니다.` : ""}` };

  const line2Parts = [];
  if (risingItems.length > 0) line2Parts.push(`${risingItems.map(t=>`${t.item}(+${t.diff}σ)`).join(", ")}은 전 사료에 걸쳐 상승 추세`);
  if (fallingItems.length > 0) line2Parts.push(`${fallingItems.map(t=>`${t.item}(${t.diff}σ)`).join(", ")}은 하락 추세`);
  const line2 = {
    icon:"📈", color:"#1e40af", bg:"#eff6ff", border:"#bfdbfe",
    text: line2Parts.length > 0
      ? `${firstYear}년 대비 ${lastYear}년 기준, ${line2Parts.join(" / ")}를 보이고 있어 원료 배합 변화 또는 계절적 요인 검토가 권장됩니다.`
      : `${firstYear}–${lastYear}년 동안 전체 항목의 Z-Score 평균 변화가 ±0.15σ 이내로 유지되어 배합 안정성이 양호한 것으로 판단됩니다.`
  };
  const line3 = {
    icon:"🔍", color:"#5b21b6", bg:"#f5f3ff", border:"#ddd6fe",
    text: `사료 간 성분 편차가 가장 큰 항목은 ${mostVaried.item}(분산 ${mostVaried.variance})으로, 제조사별 원료 품질 차이가 두드러집니다.`
  };
  return [line1, line2, line3];
}

function GlobalBriefing({ zData }) {
  const lines = React.useMemo(() => buildGlobalBriefing(zData), [zData]);
  const [open, setOpen] = React.useState(true);
  return (
    <div style={{ marginBottom:20, background:"#fff", borderRadius:12,
      border:"1px solid #e2e8f0", boxShadow:"0 1px 6px rgba(0,0,0,.05)", overflow:"hidden" }}>
      <div onClick={() => setOpen(o => !o)}
        style={{ display:"flex", alignItems:"center", justifyContent:"space-between",
          padding:"14px 20px", cursor:"pointer", borderBottom: open ? "1px solid #f1f5f9" : "none",
          background:"#fafafa", userSelect:"none" }}>
        <div style={{ display:"flex", alignItems:"center", gap:10 }}>
          <div style={{ width:4, height:18, background:"linear-gradient(180deg,#6366f1,#a855f7)", borderRadius:3 }}/>
          <span style={{ fontSize:13, fontWeight:800, color:"#0f172a" }}>전체 사료 종합 브리핑</span>
          <span style={{ fontSize:11, color:"#94a3b8" }}>— {FEEDS.length}종 · {ITEMS.length}항목 · {YEARS.length}개년 기준</span>
        </div>
        <span style={{ fontSize:12, color:"#94a3b8", display:"inline-block", transform: open ? "rotate(180deg)" : "rotate(0deg)" }}>▼</span>
      </div>
      {open && (
        <div style={{ padding:"16px 20px", display:"flex", flexDirection:"column", gap:10 }}>
          {lines.map((line, i) => (
            <div key={i} style={{ display:"flex", alignItems:"flex-start", gap:12,
              background:line.bg, border:`1px solid ${line.border}`, borderRadius:9, padding:"12px 16px" }}>
              <div style={{ display:"flex", flexDirection:"column", alignItems:"center", gap:3, flexShrink:0 }}>
                <span style={{ fontSize:15, lineHeight:1 }}>{line.icon}</span>
                <span style={{ fontSize:9, fontWeight:800, color:line.color, background:line.border, borderRadius:20, padding:"1px 5px" }}>0{i+1}</span>
              </div>
              <p style={{ margin:0, fontSize:12.5, color:"#1e293b", lineHeight:1.8, wordBreak:"keep-all" }}>{line.text}</p>
            </div>
          ))}
          <div style={{ display:"flex", justifyContent:"flex-end", marginTop:2 }}>
            <span style={{ fontSize:10.5, color:"#94a3b8" }}>기준: |Z| &gt; 2 → 주의 &nbsp;|&nbsp; |Z| &gt; 3 → 이상치</span>
          </div>
        </div>
      )}
    </div>
  );
}

function NarrativeBox({ feed, zData }) {
  const color = FEED_COLORS[feed] || "#64748b";
  const { alerts, trends } = React.useMemo(() => {
    const rows = zData.filter(r => r.feed === feed);
    const alerts = [];
    rows.forEach(row => {
      ITEMS.forEach(item => {
        const z = row[`z_${item}`];
        if (z !== null && Math.abs(z) > 2) alerts.push({ year:row.year, item, z, raw:row[item] });
      });
    });
    const trends = ITEMS.map(item => {
      const vals = YEARS.map(y => { const r = rows.find(d => d.year===y); return r ? r[`z_${item}`] : null; });
      const clean = vals.filter(v => v !== null);
      if (clean.length < 2) return null;
      return { item, diff: clean[clean.length-1]-clean[0], vals };
    }).filter(t => t && Math.abs(t.diff) > 0.3);
    return { alerts, trends };
  }, [feed, zData]);

  const sorted = [...alerts].sort((a,b) => Math.abs(b.z)-Math.abs(a.z));
  const highAlerts = sorted.filter(a => Math.abs(a.z) > 3);
  if (alerts.length === 0 && trends.length === 0) {
    return (
      <div style={{ background:"#f0fdf4", border:"1px solid #bbf7d0", borderRadius:9,
        padding:"12px 16px", marginBottom:16, display:"flex", alignItems:"center", gap:8 }}>
        <span style={{ fontSize:15 }}>✅</span>
        <p style={{ margin:0, fontSize:12.5, color:"#166534", lineHeight:1.7 }}>
          <strong>{feed}</strong> — 분석 기간 전체에서 이상치가 감지되지 않았습니다.
        </p>
      </div>
    );
  }
  return (
    <div style={{ background:"#fff", border:`1.5px solid ${color}28`,
      borderLeft:`4px solid ${color}`, borderRadius:9, padding:"14px 18px", marginBottom:16 }}>
      <div style={{ display:"flex", alignItems:"center", gap:8, marginBottom:10 }}>
        <div style={{ width:8, height:8, borderRadius:"50%", background:color }}/>
        <span style={{ fontSize:13, fontWeight:800, color:"#0f172a" }}>{feed} 분석 요약</span>
        {alerts.length > 0 && (
          <span style={{ fontSize:11, padding:"2px 8px", borderRadius:20, fontWeight:700,
            background: highAlerts.length>0 ? "#fef2f2" : "#fffbeb",
            color: highAlerts.length>0 ? "#dc2626" : "#d97706",
            border:`1px solid ${highAlerts.length>0 ? "#fca5a5" : "#fde68a"}` }}>
            {highAlerts.length>0 ? "🔴 이상치" : "🟡 주의"} {alerts.length}건
          </span>
        )}
      </div>
      {sorted.map((a, i) => {
        const isHigh = Math.abs(a.z) > 3;
        const tc = a.z > 0 ? "#b91c1c" : "#1d4ed8";
        return (
          <div key={i} style={{ display:"flex", alignItems:"flex-start", gap:8, marginBottom:7,
            padding:"8px 12px", borderRadius:7,
            background: isHigh ? "#fef2f2" : "#fffbeb",
            border:`1px solid ${isHigh ? "#fecaca" : "#fde68a"}` }}>
            <span style={{ fontSize:14, lineHeight:1.4, flexShrink:0 }}>{isHigh?"🔴":"🟡"}</span>
            <p style={{ margin:0, fontSize:12.5, color:"#1e293b", lineHeight:1.75 }}>
              <strong style={{ color:tc }}>{a.year}년</strong>의{" "}
              <strong style={{ color:"#334155" }}>{a.item}</strong>에서 이상치 발생.{" "}
              측정값 <strong style={{ color:tc }}>{typeof a.raw === "number" ? a.raw.toFixed(2) : a.raw}</strong>은{" "}
              평균 대비 <strong style={{ color:tc }}>{Math.abs(a.z).toFixed(2)}σ</strong>만큼{" "}
              <strong style={{ color:tc }}>{a.z > 0 ? "높게" : "낮게"}</strong> 측정.
              {isHigh && <span style={{ marginLeft:4, fontSize:11, color:"#dc2626", fontWeight:700 }}>(재검사 권장)</span>}
            </p>
          </div>
        );
      })}
      {trends.length > 0 && (
        <div style={{ borderTop: alerts.length>0 ? "1px solid #f1f5f9" : "none",
          paddingTop: alerts.length>0 ? 10 : 0, display:"flex", flexWrap:"wrap", gap:6, marginTop:4 }}>
          <span style={{ fontSize:11, fontWeight:700, color:"#64748b", width:"100%", marginBottom:2 }}>📈 추세</span>
          {trends.filter(t=>t.diff>0).map((t,i) => (
            <span key={i} style={{ fontSize:11.5, background:"#fef2f2", color:"#991b1b", border:"1px solid #fecaca", padding:"4px 10px", borderRadius:7 }}>
              <strong>{t.item}</strong> +{t.diff.toFixed(2)}σ 상승
            </span>
          ))}
          {trends.filter(t=>t.diff<0).map((t,i) => (
            <span key={i} style={{ fontSize:11.5, background:"#eff6ff", color:"#1e40af", border:"1px solid #bfdbfe", padding:"4px 10px", borderRadius:7 }}>
              <strong>{t.item}</strong> {t.diff.toFixed(2)}σ 하락
            </span>
          ))}
        </div>
      )}
    </div>
  );
}

function App() {
  const [selectedFeed, setSelectedFeed] = React.useState(FEEDS[0]);
  const [sortItem, setSortItem] = React.useState(null);
  const [sortDir, setSortDir] = React.useState(-1);
  const [viewMode, setViewMode] = React.useState("detail");

  const zData = React.useMemo(() => calcZScores(RAW_DATA), []);

  const feedRows = React.useMemo(() =>
    YEARS.map(year => zData.find(r => r.feed === selectedFeed && r.year === year)),
    [zData, selectedFeed]
  );

  const sparklines = React.useMemo(() => {
    const obj = {};
    ITEMS.forEach(item => {
      obj[item] = YEARS.map(y => {
        const r = zData.find(d => d.feed === selectedFeed && d.year === y);
        return r ? r[`z_${item}`] : null;
      });
    });
    return obj;
  }, [zData, selectedFeed]);

  const compareRows = React.useMemo(() => {
    const rows = FEEDS.map(feed => {
      const row = { feed };
      ITEMS.forEach(item => {
        const vals = YEARS.map(y => { const r = zData.find(d => d.feed===feed && d.year===y); return r ? r[`z_${item}`] : null; });
        row[`spark_${item}`] = vals;
        const clean = vals.filter(v => v !== null);
        row[`avg_${item}`] = clean.length ? +(clean.reduce((a,b)=>a+b,0)/clean.length).toFixed(2) : 0;
      });
      return row;
    });
    if (!sortItem) return rows;
    return [...rows].sort((a,b) => (a[`avg_${sortItem}`]-b[`avg_${sortItem}`])*sortDir);
  }, [zData, sortItem, sortDir]);

  const color = FEED_COLORS[selectedFeed] || "#64748b";

  return (
    <div style={{ fontFamily:"'Noto Sans KR','Malgun Gothic',sans-serif", background:"#f1f5f9", minHeight:"100vh", padding:"26px 28px" }}>
      <div style={{ display:"flex", alignItems:"flex-start", justifyContent:"space-between", marginBottom:20 }}>
        <div>
          <div style={{ display:"flex", alignItems:"center", gap:9, marginBottom:3 }}>
            <div style={{ width:4, height:22, background:"linear-gradient(180deg,#2563eb,#dc2626)", borderRadius:3 }}/>
            <h1 style={{ margin:0, fontSize:17, fontWeight:800, color:"#0f172a" }}>사료 Z-Score 연도별 분석 리포트</h1>
          </div>
          <p style={{ margin:"0 0 0 13px", fontSize:11, color:"#94a3b8" }}>
            {FEEDS.length}종 × {ITEMS.length}항목 × {YEARS.length}개년 ({Math.min(...YEARS)}–{Math.max(...YEARS)}) | 항목별 전체 표준화
          </p>
        </div>
        <div style={{ display:"flex", background:"#e2e8f0", borderRadius:8, padding:3, gap:2 }}>
          {[["detail","사료별 상세"],["compare","전체 비교"]].map(([v,l]) => (
            <button key={v} onClick={() => setViewMode(v)} style={{
              padding:"6px 14px", border:"none", borderRadius:6, cursor:"pointer", fontSize:12, fontWeight:600,
              background: viewMode===v ? "#fff" : "transparent",
              color: viewMode===v ? "#0f172a" : "#94a3b8",
              boxShadow: viewMode===v ? "0 1px 3px rgba(0,0,0,.1)" : "none",
            }}>{l}</button>
          ))}
        </div>
      </div>

      <GlobalBriefing zData={zData} />

      {viewMode === "detail" && (
        <div style={{ background:"#fff", borderRadius:12, border:"1px solid #e2e8f0", overflow:"hidden" }}>
          <div style={{ display:"flex", padding:"0 8px", background:"#f8fafc", borderBottom:"1px solid #e2e8f0" }}>
            {FEEDS.map(f => {
              const fc = FEED_COLORS[f] || "#64748b";
              const active = f === selectedFeed;
              return (
                <button key={f} onClick={() => setSelectedFeed(f)} style={{
                  padding:"12px 18px", border:"none", cursor:"pointer", fontSize:13, fontWeight:700,
                  background:"transparent", color: active ? fc : "#94a3b8",
                  borderBottom: active ? `2.5px solid ${fc}` : "2.5px solid transparent", marginBottom:-1,
                }}>
                  <span style={{ width:8, height:8, borderRadius:"50%", background:fc,
                    display:"inline-block", marginRight:6, opacity:active?1:0.35 }}/>
                  {f}
                </button>
              );
            })}
          </div>
          <div style={{ padding:"20px 24px" }}>
            <NarrativeBox feed={selectedFeed} zData={zData} />
            <div style={{ overflowX:"auto" }}>
              <table style={{ width:"100%", borderCollapse:"collapse", fontSize:12.5 }}>
                <thead>
                  <tr style={{ borderBottom:"2px solid #e2e8f0" }}>
                    <th style={{ padding:"8px 12px 8px 4px", textAlign:"left", color:"#64748b", fontSize:11, fontWeight:700, minWidth:68 }}>항목</th>
                    {YEARS.map(y => (
                      <th key={y} style={{ padding:"8px 14px", textAlign:"center", color:"#64748b", fontSize:11, fontWeight:700, minWidth:100 }}>
                        {y}년
                        <div style={{ fontSize:9.5, color:"#cbd5e1", fontWeight:400, marginTop:1 }}>Z-Score (원값)</div>
                      </th>
                    ))}
                    <th style={{ padding:"8px 10px", textAlign:"center", color:"#64748b", fontSize:11, fontWeight:700, minWidth:100 }}>추이</th>
                    <th style={{ padding:"8px 10px", textAlign:"center", color:"#64748b", fontSize:11, fontWeight:700, minWidth:72 }}>변화폭</th>
                    <th style={{ padding:"8px 10px", textAlign:"center", color:"#64748b", fontSize:11, fontWeight:700, minWidth:64 }}>평균Z</th>
                  </tr>
                </thead>
                <tbody>
                  {ITEMS.map((item, idx) => {
                    const zVals = sparklines[item];
                    const clean = zVals.filter(v => v !== null);
                    const avg = clean.length ? +(clean.reduce((a,b)=>a+b,0)/clean.length).toFixed(2) : 0;
                    const hasAlert = zVals.some(v => v !== null && Math.abs(v) > 2);
                    return (
                      <tr key={item} style={{ background:idx%2===0?"#f8fafc":"#fff", borderBottom:"1px solid #f1f5f9" }}>
                        <td style={{ padding:"10px 12px 10px 4px", fontWeight:700, color:"#1e293b", fontSize:12.5 }}>
                          {hasAlert && <span style={{ fontSize:10, marginRight:3 }}>⚠️</span>}
                          {item}
                        </td>
                        {YEARS.map((y, yi) => {
                          const row = feedRows[yi];
                          const z = row ? row[`z_${item}`] : null;
                          const raw = row ? row[item] : null;
                          const alert = z !== null && Math.abs(z) > 2;
                          return (
                            <td key={y} style={{ padding:"6px 8px", textAlign:"center", verticalAlign:"middle" }}>
                              <div style={{ display:"inline-flex", flexDirection:"column", alignItems:"center",
                                background: z !== null ? zBg(z) : "#f8fafc",
                                border: alert ? `1.5px solid ${zText(z)}55` : "1.5px solid transparent",
                                borderRadius:7, padding:"5px 12px", minWidth:80 }}>
                                <span style={{ fontSize:14, fontWeight:800, color: z !== null ? zText(z) : "#94a3b8", lineHeight:1.1 }}>
                                  {z !== null ? `${z > 0 ? "+" : ""}${z.toFixed(2)}` : "–"}
                                </span>
                                <span style={{ fontSize:10, color:"#94a3b8", marginTop:2, lineHeight:1 }}>
                                  ({raw !== null && raw !== undefined ? (typeof raw === "number" ? raw.toFixed(2) : raw) : "–"})
                                </span>
                              </div>
                            </td>
                          );
                        })}
                        <td style={{ padding:"6px 10px", textAlign:"center", verticalAlign:"middle" }}>
                          <div style={{ display:"flex", justifyContent:"center" }}>
                            <Sparkline values={zVals} color={color} w={88} h={34}/>
                          </div>
                          <div style={{ display:"flex", justifyContent:"center", gap:6, marginTop:2 }}>
                            {YEARS.map(y => <span key={y} style={{ fontSize:9, color:"#94a3b8" }}>{String(y).slice(2)}′</span>)}
                          </div>
                        </td>
                        <td style={{ padding:"6px 10px", textAlign:"center", verticalAlign:"middle" }}>
                          <TrendBadge values={zVals}/>
                        </td>
                        <td style={{ padding:"6px 10px", textAlign:"center", verticalAlign:"middle" }}>
                          <div style={{ display:"inline-block", background:zBg(avg), borderRadius:6,
                            padding:"3px 10px", fontSize:12, fontWeight:700, color:zText(avg) }}>
                            {avg > 0 ? "+" : ""}{avg}
                          </div>
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
            <div style={{ marginTop:10, display:"flex", gap:20, fontSize:10.5, color:"#94a3b8", flexWrap:"wrap" }}>
              <span>괄호 = 원래 측정값(중앙값)</span>
              <span style={{ color:"#b91c1c" }}>빨강 = Z &gt; 1</span>
              <span style={{ color:"#1d4ed8" }}>파랑 = Z &lt; -1</span>
              <span>⚠️ = |Z| &gt; 2 존재</span>
            </div>
          </div>
        </div>
      )}

      {viewMode === "compare" && (
        <div style={{ background:"#fff", borderRadius:12, border:"1px solid #e2e8f0", overflow:"hidden" }}>
          <div style={{ padding:"14px 24px 10px", borderBottom:"1px solid #f1f5f9", display:"flex", alignItems:"center", gap:10 }}>
            <h2 style={{ margin:0, fontSize:13, fontWeight:800, color:"#0f172a" }}>전체 사료 비교</h2>
            <span style={{ fontSize:11, color:"#94a3b8" }}>항목 헤더 클릭 → 평균Z 정렬</span>
          </div>
          <div style={{ overflowX:"auto" }}>
            <table style={{ width:"100%", borderCollapse:"collapse", fontSize:12 }}>
              <thead>
                <tr style={{ borderBottom:"2px solid #e2e8f0", background:"#f8fafc" }}>
                  <th style={{ padding:"10px 20px", textAlign:"left", color:"#64748b", fontSize:11, fontWeight:700, width:80 }}>사료</th>
                  {ITEMS.map(item => (
                    <th key={item} onClick={() => { if(sortItem===item) setSortDir(d=>-d); else { setSortItem(item); setSortDir(-1); } }}
                      style={{ padding:"10px 8px", textAlign:"center", fontSize:11, fontWeight:700,
                        color: sortItem===item ? "#1d4ed8" : "#64748b",
                        background: sortItem===item ? "#eff6ff" : "transparent",
                        cursor:"pointer", userSelect:"none", minWidth:120 }}>
                      {item}
                      <span style={{ marginLeft:3, fontSize:10 }}>{sortItem===item?(sortDir===-1?"▼":"▲"):"⇅"}</span>
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {compareRows.map((row, ri) => {
                  const fc = FEED_COLORS[row.feed] || "#64748b";
                  return (
                    <tr key={row.feed} style={{ borderBottom:"1px solid #f1f5f9", background:ri%2===0?"#fff":"#fafafa" }}>
                      <td style={{ padding:"10px 20px" }}>
                        <div style={{ display:"flex", alignItems:"center", gap:7 }}>
                          <div style={{ width:9, height:9, borderRadius:"50%", background:fc }}/>
                          <span style={{ fontWeight:800, color:fc, fontSize:13 }}>{row.feed}</span>
                        </div>
                      </td>
                      {ITEMS.map(item => {
                        const vals = row[`spark_${item}`];
                        const avg  = row[`avg_${item}`];
                        const hasAlert = vals.some(v => v !== null && Math.abs(v) > 2);
                        return (
                          <td key={item} style={{ padding:"8px 6px", textAlign:"center", verticalAlign:"middle",
                            background: sortItem===item?(ri%2===0?"#f5f9ff":"#eff6ff"):"transparent" }}>
                            <div style={{ display:"flex", flexDirection:"column", alignItems:"center", gap:3 }}>
                              <Sparkline values={vals} color={fc} w={88} h={28}/>
                              <div style={{ display:"flex", gap:3 }}>
                                {vals.map((v,i) => (
                                  <span key={i} style={{ fontSize:9.5, fontWeight:600, color: v !== null ? zText(v) : "#94a3b8",
                                    background: v !== null ? zBg(v) : "#f1f5f9", padding:"1px 5px", borderRadius:3,
                                    border: v !== null && Math.abs(v)>2 ? `1px solid ${zText(v)}55` : "none" }}>
                                    {v !== null ? `${v>0?"+":""}${v.toFixed(1)}` : "–"}
                                  </span>
                                ))}
                              </div>
                              <div style={{ display:"flex", alignItems:"center", gap:3 }}>
                                <span style={{ fontSize:10, color:"#94a3b8" }}>μ</span>
                                <span style={{ fontSize:11.5, fontWeight:800, color:zText(avg) }}>
                                  {avg>0?"+":""}{avg}
                                </span>
                                {hasAlert && <span style={{ fontSize:9 }}>⚠️</span>}
                              </div>
                            </div>
                          </td>
                        );
                      })}
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
          <div style={{ padding:"10px 24px 14px", fontSize:10.5, color:"#94a3b8" }}>
            스파크라인 끝점: 빨강=상승 / 파랑=하락 / 회색=변화없음
          </div>
        </div>
      )}

      <div style={{ marginTop:12, fontSize:10.5, color:"#94a3b8", textAlign:"right" }}>
        Z-Score = (측정값 − 항목 전체 평균) ÷ 표준편차 (ddof=1) | 측정값 = 해당 연도 참가기관 중앙값
      </div>
    </div>
  );
}

const root = ReactDOM.createRoot(document.getElementById('root'));
root.render(<App />);
"""

# 자동으로 색상 배정 (사료 수가 달라져도 대응)
_PALETTE = [
    "#2563eb", "#d97706", "#16a34a", "#dc2626", "#7c3aed",
    "#0891b2", "#db2777", "#65a30d", "#ea580c", "#6366f1",
]


def generate_history_html(raw_data: list[dict]) -> str:
    """
    raw_data: [{"feed": "축우사료", "year": 2022, "조단백": 18.2, ...}, ...]
    ITEMS, FEEDS, YEARS, FEED_COLORS는 raw_data에서 자동 추출.
    반환: 브라우저에서 바로 열 수 있는 자체완결 HTML 문자열.
    """
    feeds = sorted({r["feed"] for r in raw_data}, key=lambda f: str(f))
    years = sorted({r["year"] for r in raw_data})
    items = [k for k in raw_data[0].keys() if k not in ("feed", "year")]
    feed_colors = {f: _PALETTE[i % len(_PALETTE)] for i, f in enumerate(feeds)}

    data_json   = json.dumps(raw_data,   ensure_ascii=False)
    feeds_json  = json.dumps(feeds,      ensure_ascii=False)
    years_json  = json.dumps(years,      ensure_ascii=False)
    items_json  = json.dumps(items,      ensure_ascii=False)
    colors_json = json.dumps(feed_colors, ensure_ascii=False)

    html = f"""<!DOCTYPE html>
<html lang="ko">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>사료 Z-Score 연도별 분석 리포트</title>
  <script crossorigin src="https://unpkg.com/react@18/umd/react.production.min.js"></script>
  <script crossorigin src="https://unpkg.com/react-dom@18/umd/react-dom.production.min.js"></script>
  <script src="https://unpkg.com/@babel/standalone/babel.min.js"></script>
  <style>* {{ box-sizing: border-box; margin: 0; padding: 0; }} body {{ background:#f1f5f9; }}</style>
</head>
<body>
  <div id="root"></div>
  <script>
    /* 데이터 주입 */
    const RAW_DATA   = {data_json};
    const FEEDS      = {feeds_json};
    const YEARS      = {years_json};
    const ITEMS      = {items_json};
    const FEED_COLORS = {colors_json};
  </script>
  <script type="text/babel" data-presets="react">
{_JSX_BODY}
  </script>
</body>
</html>"""
    return html


def generate_history_html_bytes(raw_data: list[dict]) -> bytes:
    return generate_history_html(raw_data).encode("utf-8")
