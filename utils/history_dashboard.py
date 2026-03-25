"""
연도별 히스토리 데이터 → 기관별 자체완결 HTML 대시보드 생성.
CDN(React + Babel)을 사용하므로 수신자가 브라우저에서 열면 인터랙티브하게 동작합니다.

history 시트 컬럼: year | feed | institution | 성분1 | 성분2 | ...
"""
import json
import numpy as np
import pandas as pd


_PALETTE = [
    "#2563eb", "#d97706", "#16a34a", "#dc2626", "#7c3aed",
    "#0891b2", "#db2777", "#65a30d", "#ea580c", "#6366f1",
]

_JSX_BODY = r"""
function zBg(z) {
  if (z === null || z === undefined || Math.abs(z) < 3) return "transparent";
  if (z < 0) return "rgba(37,99,235,0.12)";
  return "rgba(220,38,38,0.12)";
}
function zText(z) {
  if (z === null || z === undefined) return "#94a3b8";
  if (z >= 3) return "#b91c1c";
  if (z <= -3) return "#1d4ed8";
  return "#374151";
}

function Sparkline({ values, color, w = 80, h = 30 }) {
  const clean = values.filter(v => v !== null && v !== undefined && !isNaN(v));
  if (clean.length < 2) return <span style={{fontSize:10, color:"#94a3b8"}}>–</span>;
  // z=0을 항상 포함한 범위로 스케일 설정
  const dataMin = Math.min(...clean), dataMax = Math.max(...clean);
  const min = Math.min(dataMin, 0), max = Math.max(dataMax, 0);
  const range = max - min || 0.001;
  const pad = 4;
  const xs = values.map((_, i) => pad + (i / (values.length - 1)) * (w - pad * 2));
  // null은 null 유지 (h/2로 대체하지 않음)
  const ys = values.map(v => (v === null || v === undefined) ? null : h - pad - ((v - min) / range) * (h - pad * 2));

  // null을 건너뛰며 경로 생성 (null 구간에서 M으로 이동)
  let linePath = "";
  let gapped = true;
  xs.forEach((x, i) => {
    if (ys[i] === null) { gapped = true; return; }
    linePath += `${gapped ? "M" : "L"}${x.toFixed(1)},${ys[i].toFixed(1)} `;
    gapped = false;
  });

  // 연속 구간별 area path
  let areaPath = "";
  let segStart = -1;
  for (let i = 0; i <= values.length; i++) {
    const valid = i < values.length && ys[i] !== null;
    if (valid && segStart === -1) { segStart = i; }
    else if (!valid && segStart !== -1) {
      if (i - segStart >= 2) {
        const sx = xs.slice(segStart, i), sy = ys.slice(segStart, i);
        const seg = sx.map((x, j) => `${j===0?"M":"L"}${x.toFixed(1)},${sy[j].toFixed(1)}`).join(" ");
        areaPath += `${seg} L${sx[sx.length-1]},${h-pad} L${sx[0]},${h-pad} Z `;
      }
      segStart = -1;
    }
  }

  const diff = clean[clean.length-1] - clean[0];
  const endColor = Math.abs(diff) < 0.05 ? "#94a3b8" : diff > 0 ? "#ef4444" : "#3b82f6";
  const uid = `sp${color.replace("#","")}${w}${h}${values.join("")}`;
  const y0 = h - pad - ((0 - min) / range) * (h - pad * 2);
  // 마지막 non-null 인덱스
  const lastIdx = values.reduce((acc, v, i) => (v !== null && v !== undefined) ? i : acc, -1);
  return (
    <svg width={w} height={h} style={{ display:"block", overflow:"visible" }}>
      <defs>
        <linearGradient id={`g${uid}`} x1="0" y1="0" x2="0" y2="1">
          <stop offset="0%" stopColor={color} stopOpacity="0.25"/>
          <stop offset="100%" stopColor={color} stopOpacity="0.02"/>
        </linearGradient>
      </defs>
      {/* z=0 기준선 */}
      <line x1={pad} x2={w-pad} y1={y0} y2={y0} stroke="#94a3b8" strokeWidth="1.2" strokeDasharray="3,2"/>
      <text x={w-pad+2} y={y0+3.5} fontSize="7" fill="#94a3b8">0</text>
      <path d={areaPath} fill={`url(#g${uid})`}/>
      <path d={linePath} fill="none" stroke={color} strokeWidth="1.8" strokeLinejoin="round" strokeLinecap="round"/>
      {xs.map((x, i) => ys[i] !== null && (
        <circle key={i} cx={x} cy={ys[i]} r={i === lastIdx ? 3.2 : 2.2}
          fill={i === lastIdx ? endColor : color} stroke="#fff" strokeWidth="1.2"/>
      ))}
    </svg>
  );
}

// ── 피드별 탭 뷰 ──────────────────────────────────────────────
function FeedView({ feed, items, years, myData, groupData, feedColor }) {
  return (
    <div style={{ overflowX:"auto" }}>
      <table style={{ width:"100%", borderCollapse:"collapse", fontSize:12.5 }}>
        <thead>
          <tr style={{ borderBottom:"2px solid #e2e8f0" }}>
            <th style={{ padding:"8px 12px 8px 4px", textAlign:"left", color:"#64748b", fontSize:11, fontWeight:700, minWidth:68 }}>항목</th>
            {years.map(y => (
              <th key={y} style={{ padding:"8px 10px", textAlign:"center", color:"#64748b", fontSize:11, fontWeight:700, minWidth:110 }}>
                {y}년
                <div style={{ fontSize:9.5, color:"#cbd5e1", fontWeight:400, marginTop:1 }}>Z-Score (원값)</div>
              </th>
            ))}
            <th style={{ padding:"8px 10px", textAlign:"center", color:"#64748b", fontSize:11, fontWeight:700, minWidth:90 }}>추이</th>
            <th style={{ padding:"8px 10px", textAlign:"center", color:"#64748b", fontSize:11, fontWeight:700, minWidth:60 }}>평균Z</th>
          </tr>
        </thead>
        <tbody>
          {items.map((item, idx) => {
            const zVals = years.map(y => {
              const m = myData[`${y}_${feed}_${item}`];
              return (m !== null && m !== undefined) ? m.z : null;
            });
            const rawVals = years.map(y => {
              const m = myData[`${y}_${feed}_${item}`];
              return (m !== null && m !== undefined) ? m.raw : null;
            });
            const medVals = years.map(y => {
              const g = groupData[`${y}_${feed}_${item}`];
              return (g !== null && g !== undefined) ? g.median : null;
            });
            const cleanZ = zVals.filter(v => v !== null);
            const cleanRaw = rawVals.filter(v => v !== null);
            if (cleanRaw.length === 0) return null;
            const avgZ = cleanZ.length ? +(cleanZ.reduce((a,b)=>a+b,0)/cleanZ.length).toFixed(2) : null;
            const hasAlert = zVals.some(v => v !== null && Math.abs(v) >= 3);
            return (
              <tr key={item} style={{ background:idx%2===0?"#f8fafc":"#fff", borderBottom:"1px solid #f1f5f9" }}>
                <td style={{ padding:"10px 12px 10px 4px", fontWeight:700, color:"#1e293b", fontSize:12.5 }}>
                  {hasAlert && <span style={{ fontSize:10, marginRight:3 }}>⚠️</span>}
                  {item}
                </td>
                {years.map((y, yi) => {
                  const z   = zVals[yi];
                  const raw = rawVals[yi];
                  const med = medVals[yi];
                  const alert = z !== null && Math.abs(z) >= 3;
                  return (
                    <td key={y} style={{ padding:"6px 8px", textAlign:"center", verticalAlign:"middle" }}>
                      {z !== null ? (
                        <div style={{ display:"inline-flex", flexDirection:"column", alignItems:"center",
                          background:zBg(z),
                          border: alert ? `1.5px solid ${zText(z)}55` : "1.5px solid transparent",
                          borderRadius:7, padding:"5px 10px", minWidth:90 }}>
                          <span style={{ fontSize:14, fontWeight:800, color:zText(z), lineHeight:1.1 }}>
                            {z.toFixed(2)}
                          </span>
                          <span style={{ fontSize:10, color:"#94a3b8", marginTop:1 }}>
                            {raw !== null ? raw.toFixed(2) : "–"}
                          </span>
                          {med !== null && (
                            <span style={{ fontSize:9, color:"#c4b5fd", marginTop:1 }}>
                              그룹중앙값 {med.toFixed(2)}
                            </span>
                          )}
                        </div>
                      ) : (
                        <span style={{ fontSize:12, color:"#cbd5e1" }}>–</span>
                      )}
                    </td>
                  );
                })}
                <td style={{ padding:"6px 10px", textAlign:"center", verticalAlign:"middle" }}>
                  <div style={{ display:"flex", justifyContent:"center" }}>
                    <Sparkline values={zVals} color={feedColor} w={80} h={30}/>
                  </div>
                  <div style={{ display:"flex", justifyContent:"center", gap:5, marginTop:2 }}>
                    {years.map(y => <span key={y} style={{ fontSize:9, color:"#94a3b8" }}>{String(y).slice(2)}′</span>)}
                  </div>
                </td>
                <td style={{ padding:"6px 10px", textAlign:"center", verticalAlign:"middle" }}>
                  {avgZ !== null ? (
                    <div style={{ display:"inline-block", background:zBg(avgZ), borderRadius:6,
                      padding:"3px 10px", fontSize:12, fontWeight:700, color:zText(avgZ) }}>
                      {avgZ > 0 ? "+" : ""}{avgZ}
                    </div>
                  ) : <span style={{ color:"#cbd5e1" }}>–</span>}
                </td>
              </tr>
            );
          })}
        </tbody>
      </table>
      <div style={{ marginTop:8, fontSize:10.5, color:"#94a3b8" }}>
        괄호 = 기관 제출 원값 &nbsp;|&nbsp; 그룹중앙값 = 해당 연도 전체 참가기관 중앙값
      </div>
    </div>
  );
}

// ── 요약 브리핑 ───────────────────────────────────────────────
function Briefing({ myData, feeds, items, years }) {
  const latestYear = Math.max(...years);
  const prevYears  = years.filter(y => y !== latestYear);

  const issues   = [];  // 최근 연도 |Z| > 2
  const improved = [];  // 이전에 이상 → 최근에 |Z| ≤ 1로 개선

  feeds.forEach(feed => {
    items.forEach(item => {
      const latestM = myData[`${latestYear}_${feed}_${item}`];
      const latestZ = latestM?.z ?? null;
      const prevZs  = prevYears.map(y => myData[`${y}_${feed}_${item}`]?.z ?? null).filter(v => v !== null);

      if (latestZ !== null && Math.abs(latestZ) >= 3) {
        issues.push({ feed, item, z: latestZ });
      } else if (prevZs.some(z => Math.abs(z) >= 3) && latestZ !== null && Math.abs(latestZ) <= 1) {
        const worstPrev = prevZs.reduce((a, b) => Math.abs(a) > Math.abs(b) ? a : b);
        improved.push({ feed, item, prevZ: worstPrev, latestZ });
      }
    });
  });

  if (issues.length === 0 && improved.length === 0) {
    return (
      <div style={{ background:"#f0fdf4", border:"1px solid #bbf7d0", borderRadius:10,
        padding:"14px 18px", marginBottom:20, display:"flex", alignItems:"center", gap:10 }}>
        <span style={{ fontSize:18 }}>✅</span>
        <p style={{ margin:0, fontSize:13, color:"#166534", lineHeight:1.7 }}>
          {latestYear}년 기준 모든 항목이 안정적인 범위 내에 있습니다.
        </p>
      </div>
    );
  }

  return (
    <div style={{ display:"flex", flexDirection:"column", gap:10, marginBottom:20 }}>
      {issues.length > 0 && (
        <div style={{ background:"#fef2f2", border:"1px solid #fecaca", borderRadius:10, padding:"14px 18px" }}>
          <div style={{ fontWeight:800, fontSize:13, color:"#991b1b", marginBottom:8 }}>
            ⚠️ 주의 필요 — {latestYear}년 기준 {issues.length}건 (|Z| ≥ 3)
          </div>
          <div style={{ display:"flex", flexWrap:"wrap", gap:6 }}>
            {issues.sort((a,b) => Math.abs(b.z)-Math.abs(a.z)).map((a, i) => {
              const tc = a.z > 0 ? "#b91c1c" : "#1d4ed8";
              const isHigh = Math.abs(a.z) > 3;
              return (
                <div key={i} style={{ fontSize:11.5, padding:"5px 10px", borderRadius:7,
                  background: isHigh ? "#fff1f2" : "#fff",
                  border:`1px solid ${isHigh ? "#fca5a5" : "#fecaca"}`, color:"#374151" }}>
                  <strong style={{ color:tc }}>{a.feed} · {a.item}</strong>
                  {" "}Z={a.z.toFixed(2)}
                  {isHigh && <span style={{ color:"#dc2626", fontWeight:700, marginLeft:4 }}>🔴 재검사</span>}
                </div>
              );
            })}
          </div>
        </div>
      )}

      {improved.length > 0 && (
        <div style={{ background:"#f0fdf4", border:"1px solid #bbf7d0", borderRadius:10, padding:"14px 18px" }}>
          <div style={{ fontWeight:800, fontSize:13, color:"#166534", marginBottom:8 }}>
            ✅ 개선된 항목 — 이전 이상치 → {latestYear}년 정상 범위 ({improved.length}건)
          </div>
          <div style={{ display:"flex", flexWrap:"wrap", gap:6 }}>
            {improved.map((a, i) => (
              <div key={i} style={{ fontSize:11.5, padding:"5px 10px", borderRadius:7,
                background:"#fff", border:"1px solid #bbf7d0", color:"#374151" }}>
                <strong style={{ color:"#166534" }}>{a.feed} · {a.item}</strong>
                {" "}Z: <span style={{ color:"#6b7280" }}>{a.prevZ.toFixed(2)}</span>
                {" → "}<span style={{ color:"#166534", fontWeight:700 }}>{a.latestZ.toFixed(2)}</span>
              </div>
            ))}
          </div>
        </div>
      )}
    </div>
  );
}

// ── 메인 ──────────────────────────────────────────────────────
function App() {
  const feeds = [...new Set(Object.keys(MY_DATA).map(k => k.split("_")[1]))];
  const items = ITEMS;
  const years = YEARS;
  const [selectedFeed, setSelectedFeed] = React.useState(feeds[0] || "");

  const feedColor = FEED_COLORS[selectedFeed] || "#2563eb";

  return (
    <div style={{ fontFamily:"'Noto Sans KR','Malgun Gothic',sans-serif", background:"#f1f5f9", minHeight:"100vh", padding:"26px 28px" }}>
      {/* 헤더 */}
      <div style={{ marginBottom:20 }}>
        <div style={{ display:"flex", alignItems:"center", gap:9, marginBottom:4 }}>
          <div style={{ width:4, height:24, background:"linear-gradient(180deg,#2563eb,#dc2626)", borderRadius:3 }}/>
          <h1 style={{ margin:0, fontSize:18, fontWeight:800, color:"#0f172a" }}>
            {INSTITUTION} — 연도별 Z-Score 추이 리포트
          </h1>
        </div>
        <p style={{ margin:"0 0 0 13px", fontSize:11, color:"#94a3b8" }}>
          {feeds.length}종 사료 · {items.length}항목 · {years.length}개년 ({Math.min(...years)}–{Math.max(...years)}) | 각 연도 그룹 전체 기준 표준화
        </p>
      </div>

      {/* 브리핑 */}
      <Briefing myData={MY_DATA} feeds={feeds} items={items} years={years} />

      {/* 사료별 탭 */}
      <div style={{ background:"#fff", borderRadius:12, border:"1px solid #e2e8f0", overflow:"hidden" }}>
        <div style={{ display:"flex", padding:"0 8px", background:"#f8fafc", borderBottom:"1px solid #e2e8f0" }}>
          {feeds.map(f => {
            const fc = FEED_COLORS[f] || "#64748b";
            const active = f === selectedFeed;
            return (
              <button key={f} onClick={() => setSelectedFeed(f)} style={{
                padding:"11px 18px", border:"none", cursor:"pointer", fontSize:13, fontWeight:700,
                background:"transparent", color: active ? fc : "#94a3b8",
                borderBottom: active ? `2.5px solid ${fc}` : "2.5px solid transparent",
                marginBottom:-1,
              }}>
                <span style={{ width:8, height:8, borderRadius:"50%", background:fc,
                  display:"inline-block", marginRight:6, opacity:active?1:0.4 }}/>
                {f}
              </button>
            );
          })}
        </div>
        <div style={{ padding:"20px 24px" }}>
          <FeedView
            feed={selectedFeed} items={items} years={years}
            myData={MY_DATA} groupData={GROUP_DATA} feedColor={feedColor}
          />
        </div>
      </div>

      <div style={{ marginTop:12, fontSize:10.5, color:"#94a3b8", textAlign:"right" }}>
        Z-Score = (기관값 − 그룹중앙값) ÷ (1.4826 × MAD) | 5개 이하 참가 시 N/A
      </div>
    </div>
  );
}

const root = ReactDOM.createRoot(document.getElementById('root'));
root.render(<App />);
"""


def _robust_z(value: float, values: list[float]) -> float | None:
    arr = np.array([v for v in values if v is not None and not np.isnan(v)], dtype=float)
    if len(arr) < 2:
        return None
    median = float(np.median(arr))
    q1, q3 = float(np.percentile(arr, 25)), float(np.percentile(arr, 75))
    niqr = (q3 - q1) * 0.7413
    if niqr == 0:
        std = float(np.std(arr))
        if std == 0:
            return 0.0
        return round((value - float(np.mean(arr))) / std, 3)
    return round((value - median) / niqr, 3)


def _ordered_item_cols(history_df: pd.DataFrame) -> list[str]:
    """
    config 성분 순서에 맞게 item_cols 정렬.
    그룹 order → 성분 order 2단계 정렬. config에 없는 성분은 뒤에 추가.
    """
    raw_cols = [c for c in history_df.columns if c not in ("year", "feed", "institution")]
    try:
        from utils.config import get_config
        cfg = get_config()

        # 그룹별 order 매핑
        group_order_map = {
            row["name"]: int(row["order"])
            for _, row in cfg[cfg["type"] == "group"].iterrows()
        }

        comp_rows = cfg[cfg["type"] == "component"].copy()
        comp_rows["group_order"] = comp_rows["group"].map(group_order_map).fillna(999).astype(int)
        comp_rows["comp_order"] = comp_rows["order"].fillna(999).astype(int)
        comp_rows = comp_rows.sort_values(["group_order", "comp_order"])

        cfg_order = comp_rows["name"].tolist()
        ordered   = [c for c in cfg_order if c in raw_cols]
        remaining = [c for c in raw_cols if c not in ordered]
        return ordered + remaining
    except Exception:
        return raw_cols


def generate_institution_html(
    history_df: pd.DataFrame,
    institution: str,
) -> str:
    """
    history_df 컬럼: year | feed | institution | 성분1 | 성분2 | ...
    institution: 이 HTML을 받는 기관명
    반환: 자체완결 HTML 문자열
    성분 순서는 config 기준, 데이터 없는 연도는 '–' 표시.
    """
    if history_df.empty:
        return ""

    item_cols = _ordered_item_cols(history_df)
    feeds = sorted(history_df["feed"].dropna().unique().tolist())
    years = sorted([int(y) for y in history_df["year"].dropna().unique()])

    # MY_DATA: {"{year}_{feed}_{item}": {"raw": float, "z": float}}
    my_data: dict = {}
    # GROUP_DATA: {"{year}_{feed}_{item}": {"median": float, "n": int}}
    group_data: dict = {}

    for feed in feeds:
        for year in years:
            year_feed = history_df[
                (history_df["feed"] == feed) &
                (history_df["year"] == year)
            ]
            if year_feed.empty:
                continue
            for item in item_cols:
                all_vals = pd.to_numeric(year_feed[item], errors="coerce").dropna().tolist()
                key = f"{year}_{feed}_{item}"
                group_data[key] = {
                    "median": round(float(np.median(all_vals)), 4) if all_vals else None,
                    "n": len(all_vals),
                }
                my_row = year_feed[year_feed["institution"] == institution]
                if my_row.empty:
                    continue
                raw_val = pd.to_numeric(my_row.iloc[0][item], errors="coerce")
                if pd.isna(raw_val):
                    continue
                raw_val = float(raw_val)
                z = _robust_z(raw_val, all_vals)
                my_data[key] = {"raw": raw_val, "z": z}

    feed_colors = {f: _PALETTE[i % len(_PALETTE)] for i, f in enumerate(feeds)}

    html = f"""<!DOCTYPE html>
<html lang="ko">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>{institution} — 연도별 Z-Score 추이</title>
  <script crossorigin src="https://unpkg.com/react@18/umd/react.production.min.js"></script>
  <script crossorigin src="https://unpkg.com/react-dom@18/umd/react-dom.production.min.js"></script>
  <script src="https://unpkg.com/@babel/standalone/babel.min.js"></script>
  <style>* {{ box-sizing: border-box; margin: 0; padding: 0; }} body {{ background:#f1f5f9; }}</style>
</head>
<body>
  <div id="root"></div>
  <script>
    const INSTITUTION = {json.dumps(institution, ensure_ascii=False)};
    const ITEMS       = {json.dumps(item_cols, ensure_ascii=False)};
    const YEARS       = {json.dumps(years, ensure_ascii=False)};
    const FEED_COLORS = {json.dumps(feed_colors, ensure_ascii=False)};
    const MY_DATA     = {json.dumps(my_data, ensure_ascii=False)};
    const GROUP_DATA  = {json.dumps(group_data, ensure_ascii=False)};
  </script>
  <script type="text/babel" data-presets="react">
{_JSX_BODY}
  </script>
</body>
</html>"""
    return html


def generate_institution_html_bytes(history_df: pd.DataFrame, institution: str) -> bytes:
    return generate_institution_html(history_df, institution).encode("utf-8")


def _z_td(z: float | None, raw: float | None, med: float | None) -> str:
    """이메일용 정적 Z-score 셀 HTML."""
    if z is None:
        return '<td style="padding:6px 10px;text-align:center;color:#94a3b8;">–</td>'
    abs_z = abs(z)
    if abs_z > 3:
        bg, color, border = "#fef2f2", "#b91c1c", "1px solid #fca5a5"
    elif abs_z > 2:
        bg, color, border = "#fff7ed", "#c2410c", "1px solid #fed7aa"
    elif abs_z > 1:
        bg, color, border = "#fef9f0", "#d97706", "none"
    elif z < -1:
        bg, color, border = "#eff6ff", "#1d4ed8", "none"
    else:
        bg, color, border = "#f8fafc", "#374151", "none"

    z_str = f"{z:.2f}"
    raw_str = f"{raw:.2f}" if raw is not None else "–"
    med_str = f"그룹중앙 {med:.2f}" if med is not None else ""
    warn = " ⚠" if abs_z > 2 else ""

    return f"""<td style="padding:6px 8px;text-align:center;vertical-align:middle;">
  <div style="display:inline-block;background:{bg};border:{border};border-radius:6px;padding:4px 10px;min-width:80px;">
    <div style="font-size:13px;font-weight:700;color:{color};">{z_str}{warn}</div>
    <div style="font-size:10px;color:#94a3b8;">{raw_str}</div>
    <div style="font-size:9px;color:#c4b5fd;">{med_str}</div>
  </div>
</td>"""


def generate_institution_email_html(
    history_df: pd.DataFrame,
    institution: str,
) -> str:
    """
    이메일 본문 삽입용 정적 HTML (JavaScript 없음).
    Z-score 표를 사료별로 생성하며 색상 강조 포함.
    """
    if history_df.empty:
        return ""

    item_cols = _ordered_item_cols(history_df)
    feeds = sorted(history_df["feed"].dropna().unique().tolist())
    years = sorted([int(y) for y in history_df["year"].dropna().unique()])

    # 데이터 계산 (generate_institution_html와 동일 로직)
    my_data: dict = {}
    group_data: dict = {}
    for feed in feeds:
        for year in years:
            year_feed = history_df[(history_df["feed"] == feed) & (history_df["year"] == year)]
            if year_feed.empty:
                continue
            for item in item_cols:
                all_vals = pd.to_numeric(year_feed[item], errors="coerce").dropna().tolist()
                key = f"{year}_{feed}_{item}"
                group_data[key] = float(np.median(all_vals)) if all_vals else None
                my_row = year_feed[year_feed["institution"] == institution]
                if my_row.empty:
                    continue
                raw_val = pd.to_numeric(my_row.iloc[0][item], errors="coerce")
                if pd.isna(raw_val):
                    continue
                raw_val = float(raw_val)
                z = _robust_z(raw_val, all_vals)
                my_data[key] = {"raw": raw_val, "z": z}

    year_headers = "".join(
        f'<th style="padding:8px 10px;text-align:center;color:#475569;font-size:11px;min-width:100px;">'
        f'{y}년<br><span style="font-weight:400;color:#94a3b8;font-size:9px;">Z (원값)</span></th>'
        for y in years
    )

    sections = []
    for feed in feeds:
        rows_html = []
        for i, item in enumerate(item_cols):
            bg_row = "#f8fafc" if i % 2 == 0 else "#ffffff"
            tds = []
            for year in years:
                key = f"{year}_{feed}_{item}"
                md = my_data.get(key)
                z   = md["z"]   if md else None
                raw = md["raw"] if md else None
                med = group_data.get(key)
                tds.append(_z_td(z, raw, med))
            rows_html.append(
                f'<tr style="background:{bg_row};">'
                f'<td style="padding:8px 10px;font-weight:600;color:#1e293b;font-size:12px;white-space:nowrap;">{item}</td>'
                + "".join(tds) +
                "</tr>"
            )

        sections.append(f"""
<h3 style="margin:20px 0 8px;font-size:13px;color:#1e293b;border-left:3px solid #2563eb;padding-left:8px;">{feed}</h3>
<table style="border-collapse:collapse;width:100%;font-size:12px;background:#fff;border-radius:8px;overflow:hidden;border:1px solid #e2e8f0;">
  <thead>
    <tr style="background:#f1f5f9;border-bottom:2px solid #e2e8f0;">
      <th style="padding:8px 10px;text-align:left;color:#475569;font-size:11px;min-width:70px;">항목</th>
      {year_headers}
    </tr>
  </thead>
  <tbody>{"".join(rows_html)}</tbody>
</table>""")

    return f"""
<div style="font-family:'Malgun Gothic','Apple SD Gothic Neo',sans-serif;max-width:700px;margin:0 auto;color:#1e293b;">
  <div style="background:linear-gradient(135deg,#1e40af,#7c3aed);padding:20px 24px;border-radius:10px 10px 0 0;">
    <h2 style="margin:0;color:#fff;font-size:16px;">📊 {institution} — 연도별 Z-Score 추이</h2>
    <p style="margin:4px 0 0;color:#bfdbfe;font-size:11px;">{min(years)}–{max(years)}년 · {len(feeds)}종 사료 · {len(item_cols)}항목 | 각 연도 전체 참가기관 기준 표준화</p>
  </div>
  <div style="background:#fff;padding:16px 20px;border:1px solid #e2e8f0;border-top:none;border-radius:0 0 10px 10px;">
    {"".join(sections)}
    <p style="margin-top:14px;font-size:10px;color:#94a3b8;text-align:right;">
      ⚠ = |Z|&gt;2 주의 &nbsp;|&nbsp; 빨강 = |Z|&gt;3 이상치 &nbsp;|&nbsp; 괄호=기관 원값 &nbsp;|&nbsp; 그룹중앙=해당 연도 참가기관 중앙값<br>
      첨부 HTML 파일을 브라우저에서 열면 인터랙티브 대시보드를 확인할 수 있습니다.
    </p>
  </div>
</div>"""
