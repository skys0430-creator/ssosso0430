// 협의체 심의 발표자료(PPT) 생성: node build.js <out.pptx>
// 필요: pptxgenjs, pptx 스킬의 scripts/apply_theme.js (환경변수 APPLY_THEME 로 경로 지정)
const path = require("path");
const pptxgen = require("pptxgenjs");
const { applyTheme } = require(process.env.APPLY_THEME);

const OUT = process.argv[2] || "deck.pptx";
const FIG = path.join(__dirname, "..", "figures");
const FONT = "맑은 고딕";
const THEME = {
  name: "Uiwang Care 2027", headFontFace: FONT, bodyFontFace: FONT,
  colors: {
    dk1: "1D2B3A", lt1: "FFFFFF", dk2: "1D3A5C", lt2: "EEF5FB",
    accent1: "24507D", accent2: "E58A1F", accent3: "7FA1C4",
    accent4: "D9E8F5", accent5: "4D6680", accent6: "9FB6CC",
    hlink: "24507D", folHlink: "4D6680",
  },
};
const H = THEME.colors;

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.33 x 7.5
pres.theme = { headFontFace: FONT, bodyFontFace: FONT };
pres.title = "2027년 의왕시 통합지원 지역계획(안) 심의";
pres.author = "의왕시 복지정책과 통합돌봄팀";
const C = pres.SchemeColor;
const NAVY = C.text2, INK = C.text1, WHITE = C.background1, TINT = C.background2;
const BLUE = C.accent1, ORANGE = C.accent2, MID = C.accent3, PALE = C.accent4, MUTED = C.accent5;

const W = 13.33, MX = 0.6, CW = W - 2 * MX;
const FOOT = "2027년 의왕시 통합지원 지역계획(안) · 통합지원협의체 심의";

pres.defineSlideMaster({
  title: "COVER", background: { color: H.dk2 },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: MX, y: 2.55, w: 9.6, h: 1.9, fontSize: 40, bold: true, color: WHITE, valign: "top", align: "left", margin: 0 }, text: "" } },
    { placeholder: { options: { name: "body", type: "body", x: MX, y: 4.75, w: 9.6, h: 0.9, fontSize: 18, color: PALE, valign: "top", margin: 0 }, text: "" } },
  ],
});
pres.defineSlideMaster({
  title: "CONTENT", background: { color: H.lt1 },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: MX, y: 0.4, w: CW, h: 0.9, fontSize: 30, bold: true, color: NAVY, valign: "middle", align: "left", margin: 0 }, text: "" } },
    { text: { text: FOOT, options: { x: MX, y: 6.98, w: 8, h: 0.3, fontSize: 10, color: MUTED, margin: 0 } } },
  ],
  slideNumber: { x: W - MX - 0.8, y: 6.98, w: 0.8, h: 0.3, fontSize: 10, color: MUTED, align: "right" },
});
pres.defineSlideMaster({
  title: "CLOSING", background: { color: H.dk2 },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: MX, y: 0.55, w: CW, h: 0.9, fontSize: 30, bold: true, color: WHITE, valign: "middle", align: "left", margin: 0 }, text: "" } },
  ],
});

// ---------- helpers ----------
let n = 0;
const nm = (p) => `${p}-${++n}`;
function box(s, x, y, w, h, fill, o = {}) {
  s.addShape(pres.ShapeType.roundRect, { x, y, w, h, fill: { color: fill }, line: o.line ? { color: o.line, width: 1 } : { type: "none" }, rectRadius: o.r ?? 0.12, objectName: nm("card") });
}
function T(s, text, x, y, w, h, o = {}) {
  s.addText(text, { x, y, w, h, isTextBox: true, margin: 0, fontSize: 14, color: INK, valign: "top", objectName: nm("text"), ...o });
}
function num(s, label, x, y, d = 0.5, fill = NAVY) {
  s.addShape(pres.ShapeType.ellipse, { x, y, w: d, h: d, fill: { color: fill }, line: { type: "none" }, objectName: nm("dot") });
  T(s, label, x, y, d, d, { fontSize: 14, bold: true, color: WHITE, align: "center", valign: "middle" });
}
function badge(s, label, x, y, w = 0.7, fill = ORANGE) {
  box(s, x, y, w, 0.3, fill, { r: 0.06 });
  T(s, label, x, y, w, 0.3, { fontSize: 11, bold: true, color: WHITE, align: "center", valign: "middle" });
}
function slide(section, title, notes) {
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: section });
  s.addText(title, { placeholder: "title" });
  if (notes) s.addNotes(notes);
  return s;
}
function fig(s, file, x, y, w, ratio, alt) {
  s.addImage({ path: path.join(FIG, file), x, y, w, h: w / ratio, altText: alt, objectName: nm("figure") });
}
const chartText = { catAxisLabelFontFace: "+mn-lt", valAxisLabelFontFace: "+mn-lt", dataLabelFontFace: "+mn-lt", legendFontFace: "+mn-lt",
  catAxisLabelColor: H.dk1, valAxisLabelColor: H.accent5, dataLabelColor: H.dk1, catAxisLabelFontSize: 12, valAxisLabelFontSize: 11, dataLabelFontSize: 12 };

// ================= 1. 표지 =================
const S1 = "개요", S2 = "2026 추진 분석", S3 = "2027 계획", S4 = "예산·심의";
pres.addSection({ title: S1 });
{
  const s = pres.addSlide({ masterName: "COVER", sectionTitle: S1 });
  badge(s, "심의 안건", MX, 1.9, 1.1);
  s.addText("2027년 의왕시\n통합지원 지역계획(안)", { placeholder: "title" });
  s.addText("살던 곳에서 건강하게, 함께 살피는 의왕", { placeholder: "body" });
  T(s, "의왕시 통합지원협의체  |  2026. 11. 26.", MX, 6.3, 8, 0.35, { fontSize: 14, color: PALE });
  T(s, "복지정책과 통합돌봄팀", MX, 6.68, 8, 0.35, { fontSize: 14, color: PALE });
  s.addShape(pres.ShapeType.ellipse, { x: 10.3, y: 1.2, w: 5.2, h: 5.2, fill: { color: BLUE }, line: { type: "none" }, objectName: nm("deco") });
  s.addShape(pres.ShapeType.ellipse, { x: 11.5, y: 4.3, w: 2.6, h: 2.6, fill: { color: ORANGE }, line: { type: "none" }, objectName: nm("deco") });
  s.addNotes("2027년 의왕시 통합지원 지역계획(안)을 보고드리고 심의를 요청드립니다.");
}

// ================= 2. 심의 개요 =================
{
  const s = slide(S1, "심의 안건 개요", "법에 따라 매년 지역계획을 수립하고 협의체 심의를 거칩니다. 오늘은 2026년 분석, 2027년 목표·과제, 예산 순으로 말씀드립니다.");
  const rows = [
    ["안건", "2027년 의왕시 통합지원 지역계획(안)"],
    ["근거", "「의료·요양 등 지역 돌봄의 통합지원에 관한 법률」 제6조(지역계획), 제20조(통합지원협의체)"],
    ["계획기간", "2027. 1. ~ 12.  (2026년 추진 분석 및 건보공단 실태조사 결과 반영)"],
    ["요청사항", "지역계획(안)의 목표·성과지표·추진과제·예산에 대한 심의"],
  ];
  rows.forEach((r, i) => {
    const y = 1.6 + i * 0.78;
    box(s, MX, y, 1.7, 0.6, NAVY, { r: 0.08 });
    T(s, r[0], MX, y, 1.7, 0.6, { bold: true, color: WHITE, align: "center", valign: "middle", fontSize: 15 });
    T(s, r[1], MX + 1.95, y, CW - 1.95, 0.6, { fontSize: 15, valign: "middle" });
  });
  T(s, "보고 순서", MX, 4.85, 3, 0.35, { fontSize: 16, bold: true, color: NAVY });
  const steps = [["2026년 추진 분석", "성과 · 지역 여건 · 진단"], ["2027년 목표", "비전 · 대상자 · 성과지표"], ["추진과제", "체계 · 발굴 · 퇴원 · 서비스"], ["예산 및 심의 요청", "사업비 689,636천원"]];
  const sw = (CW - 3 * 0.3) / 4;
  steps.forEach((st, i) => {
    const x = MX + i * (sw + 0.3);
    box(s, x, 5.3, sw, 1.3, TINT);
    num(s, String(i + 1), x + 0.2, 5.5, 0.45);
    T(s, st[0], x + 0.8, 5.48, sw - 0.95, 0.5, { fontSize: 15, bold: true, color: NAVY, valign: "middle" });
    T(s, st[1], x + 0.8, 6.02, sw - 0.95, 0.4, { fontSize: 12, color: MUTED });
  });
}

// ================= 3. 2026 성과 =================
pres.addSection({ title: S2 });
{
  const s = slide(S2, "2026년 추진 성과: 신청 393명, 대상자 276명", "10월 말 기준 신청 393명, 대상자 276명입니다. 10월 한 달에 147명이 신청해 노인 1만명당 신청이 119.4명으로 올랐습니다. 연말 450명 내외를 예상합니다.");
  const st = [
    ["393명", "통합돌봄 신청", "’26.10.31. 기준 · 연말 450명 내외 예상"],
    ["276명", "통합지원 대상자", "우선발굴필요군 222명 포함"],
    ["80.4%", "우선발굴필요군 비중", "대상자 276명 중 222명"],
    ["119.4명", "노인 1만명당 신청", "’26.9. 74.8명 → 10월 신청 147명"],
  ];
  const cw = (CW - 3 * 0.3) / 4;
  st.forEach((v, i) => {
    const x = MX + i * (cw + 0.3);
    box(s, x, 1.6, cw, 2.55, TINT);
    T(s, v[0], x + 0.25, 1.85, cw - 0.5, 0.95, { fontSize: 44, bold: true, color: i === 3 ? ORANGE : NAVY, valign: "middle" });
    T(s, v[1], x + 0.25, 2.9, cw - 0.5, 0.4, { fontSize: 14, bold: true, color: INK });
    T(s, v[2], x + 0.25, 3.35, cw - 0.5, 0.6, { fontSize: 12, color: MUTED });
  });
  T(s, "2026년 운영 기반", MX, 4.5, 4, 0.35, { fontSize: 16, bold: true, color: NAVY });
  const base = [
    ["동 전담인력", "전담인력(복지) 동별 1명, ’26.11.30. 배치"],
    ["재택의료 자원", "방문진료 참여의원 14개소(의과 2 · 한의 12), 재택의료센터 1개소"],
    ["퇴원환자 연계", "협약병원 3개소(’26.7.)"],
    ["의왕살핌 정산", "80,361천원(’26.10.) · 집행률 55.4%"],
  ];
  const bw = (CW - 0.3) / 2;
  base.forEach((b, i) => {
    const x = MX + (i % 2) * (bw + 0.3), y = 4.98 + Math.floor(i / 2) * 0.88;
    box(s, x, y, bw, 0.72, WHITE, { line: PALE, r: 0.08 });
    T(s, b[0], x + 0.2, y, 1.7, 0.72, { fontSize: 14, bold: true, color: NAVY, valign: "middle" });
    T(s, b[1], x + 1.9, y, bw - 2.05, 0.72, { fontSize: 13, valign: "middle" });
  });
}

// ================= 4. 지역 여건 =================
{
  const s = slide(S2, "지역 여건: 노인의 64%가 청계·오전·부곡에 거주", "고령화율 20.3%로 초고령사회에 들어섰습니다. 노인의 64.2%가 3개 동에 살고, 부곡동은 기초연금·기초생활 비중이 가장 높습니다.");
  const st = [["33,930명", "노인인구(’26.10.31.)"], ["20.3%", "고령화율"], ["6,607세대", "65세 이상 1인세대(85세 이상 1,046세대)"]];
  st.forEach((v, i) => {
    const y = 1.6 + i * 1.72;
    box(s, MX, y, 4.3, 1.5, i === 0 ? NAVY : TINT);
    T(s, v[0], MX + 0.3, y + 0.18, 3.7, 0.75, { fontSize: 36, bold: true, color: i === 0 ? WHITE : NAVY, valign: "middle" });
    T(s, v[1], MX + 0.3, y + 0.95, 3.7, 0.4, { fontSize: 13, color: i === 0 ? PALE : MUTED });
  });
  T(s, "동별 취약 지표(%)", 5.3, 1.6, 5, 0.35, { fontSize: 16, bold: true, color: NAVY });
  s.addChart(pres.charts.BAR, [
    { name: "기초연금 수급", labels: ["부곡동", "오전동", "청계동", "의왕시"], values: [70.3, 55.9, 40.5, 50.3] },
    { name: "노인 기초생활수급", labels: ["부곡동", "오전동", "청계동", "의왕시"], values: [12.0, 3.4, 8.3, 6.9] },
  ], { x: 5.2, y: 2.0, w: 7.5, h: 3.55, barDir: "col", barGapWidthPct: 60, chartColors: [H.accent1, H.accent2],
    showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: "0.0", showLegend: true, legendPos: "t", legendFontSize: 12, legendColor: H.dk1,
    valAxisMaxVal: 80, valAxisMinVal: 0, valAxisMajorUnit: 20, valGridLine: { color: H.accent4, size: 0.75 }, catGridLine: { style: "none" }, ...chartText, objectName: nm("chart") });
  box(s, 5.3, 5.7, 7.43, 0.95, TINT);
  badge(s, "부곡동", 5.5, 6.03, 0.85);
  T(s, "단독·연립·다세대 32.8%, 독거 26.2%로 저소득 고령가구 밀집", 6.55, 5.7, 6.05, 0.95, { fontSize: 14, valign: "middle" });
}

// ================= 5. 진단 → 과제 =================
{
  const s = slide(S2, "진단: 신청은 늘었지만 ‘먼저 찾는’ 발굴이 약함", "신청 대부분이 본인·가족 신청이고 직권·자체발굴은 4.3%입니다. 2027년은 먼저 찾고, 끝까지 살피는 체계로 바꾸는 것이 핵심입니다.");
  T(s, "2026 진단", MX, 1.5, 5.4, 0.35, { fontSize: 14, bold: true, color: MUTED });
  T(s, "2027 대응", 7.3, 1.5, 5.4, 0.35, { fontSize: 14, bold: true, color: MUTED });
  const rows = [
    ["4.3%", "직권·자체발굴 비중(17건 ÷ 393명)", "초고령 위험군 선제발굴, 건보 빅데이터 활용 → 발굴률 20% 이상"],
    ["3.6%", "우선발굴필요군 중 대상자 비중", "분모 6,506명 기준 6% 이상(우선 421명)"],
    ["55.4%", "의왕살핌 집행률(’26.10.)", "치료식 맞춤형 영양식 개편, 이동·회복살핌 등 신규서비스"],
    ["지연·분절", "서비스 개시 확인 지연, 복합욕구 분절 지원", "의뢰·회신 표준체계, 동 전담인력 모니터링 일원화"],
  ];
  rows.forEach((r, i) => {
    const y = 1.92 + i * 1.22;
    box(s, MX, y, 5.9, 1.05, TINT);
    T(s, r[0], MX + 0.25, y, 1.95, 1.05, { fontSize: i === 3 ? 22 : 30, bold: true, color: NAVY, valign: "middle" });
    T(s, r[1], MX + 2.25, y, 3.5, 1.05, { fontSize: 14, valign: "middle" });
    s.addShape(pres.ShapeType.rightArrow, { x: 6.67, y: y + 0.33, w: 0.45, h: 0.4, fill: { color: ORANGE }, line: { type: "none" }, objectName: nm("arrow") });
    box(s, 7.3, y, CW - 6.7, 1.05, WHITE, { line: MID });
    T(s, r[2], 7.55, y, CW - 7.2, 1.05, { fontSize: 14, bold: true, color: NAVY, valign: "middle" });
  });
}

// ================= 6. 비전 =================
pres.addSection({ title: S3 });
{
  const s = slide(S3, "2027 비전과 추진과제", "비전은 ‘살던 곳에서 건강하게, 함께 살피는 의왕’입니다. 3대 목표 아래 4개 추진과제, 12개 핵심사업으로 구성했습니다.");
  fig(s, "의왕시_비전체계도.png", 0.97, 1.42, 11.4, 2280 / 1096, "2027 의왕시 통합지원 비전 체계도: 비전, 3대 목표, 4개 추진과제와 핵심사업");
}

// ================= 7. 목표 대상자 =================
{
  const s = slide(S3, "2027 목표: 통합지원 대상자 541명", "2027년 목표 대상자는 541명으로 2026년 276명의 약 2배입니다. 우선발굴필요군 6,506명 중 421명, 6.5%를 대상자로 지원합니다.");
  box(s, MX, 1.6, 4.3, 2.3, NAVY);
  T(s, "541명", MX + 0.3, 1.8, 3.7, 1.1, { fontSize: 54, bold: true, color: WHITE, valign: "middle" });
  T(s, "우선발굴필요군 421명 + 기타 120명", MX + 0.3, 2.95, 3.7, 0.35, { fontSize: 14, color: WHITE });
  T(s, "2026년 276명 대비 약 2배", MX + 0.3, 3.33, 3.7, 0.35, { fontSize: 13, color: PALE });
  box(s, MX, 4.1, 4.3, 2.55, TINT);
  T(s, "우선발굴필요군(분모)", MX + 0.3, 4.28, 3.7, 0.35, { fontSize: 14, bold: true, color: NAVY });
  T(s, "6,506명", MX + 0.3, 4.65, 3.7, 0.7, { fontSize: 32, bold: true, color: NAVY, valign: "middle" });
  T(s, "건보공단 중복제거 6,208명\n+ 65세 미만 중증장애인 298명", MX + 0.3, 5.42, 3.7, 0.6, { fontSize: 12, color: MUTED });
  T(s, "421 ÷ 6,506 = 6.5% (목표 6% 이상)", MX + 0.3, 6.1, 3.7, 0.35, { fontSize: 13, bold: true, color: ORANGE });
  T(s, "유형별 목표 인원(명)", 5.3, 1.6, 5, 0.35, { fontSize: 16, bold: true, color: NAVY });
  const lab = ["재가급여 수급", "고령 장애인", "퇴원환자", "등급외·노인맞춤돌봄", "장기요양 기각", "중증장애인(65세 미만)", "시설 퇴소", "기타(우선발굴필요군 외)"];
  const val = [145, 107, 60, 50, 40, 17, 2, 120];
  s.addChart(pres.charts.BAR, [{ name: "목표 인원", labels: lab.slice().reverse(), values: val.slice().reverse() }],
    { x: 5.2, y: 1.95, w: 7.5, h: 4.7, barDir: "bar", barGapWidthPct: 45, chartColors: [H.accent1], showValue: true, dataLabelPosition: "outEnd",
      showLegend: false, valAxisHidden: true, valAxisMaxVal: 170, valGridLine: { style: "none" }, catGridLine: { style: "none" }, ...chartText, objectName: nm("chart") });
}

// ================= 8. 성과지표 =================
{
  const s = slide(S3, "성과지표: 합동평가 80점, 발굴률 20% 이상", "중앙 지표는 행안부 합동평가와 복지부 지역복지사업 평가 기준을 따랐습니다. 합동평가 80점 달성을 위해 발굴률 20% 이상이 필요합니다.");
  const hd = (t, w) => ({ text: t, options: { bold: true, color: WHITE, fill: { color: NAVY }, align: "center", valign: "middle" } });
  const rows = [
    ["통합돌봄 대상자 발굴률(직권·자체발굴 ÷ 신청)", "4.3%", "20% 이상"],
    ["우선발굴필요군 중 통합지원 대상자 비중", "3.6%", "6% 이상"],
    ["통합지원 대상자 중 우선발굴필요군 비중", "80.4%", "70% 이상"],
    ["개인별 지원계획 수립률", "-", "6% 이상"],
    ["모니터링 수행률", "-", "80%"],
    ["전담인력 실제 배치율 / 본청 다직렬 배치", "-", "100% / 2개 직렬"],
    ["보건소 통합지원회의 참석률", "-", "90% 이상"],
    ["보건소 통합돌봄 서비스 이용률(건수 기준)", "-", "25% 이상"],
  ];
  const body = rows.map((r, i) => r.map((c, j) => ({ text: c, options: { align: j ? "center" : "left", valign: "middle", bold: j === 2, color: j === 2 ? NAVY : INK, fill: { color: i % 2 ? TINT : WHITE } } })));
  s.addTable([[hd("중앙정부 기반 지표"), hd("’26년"), hd("’27 목표")], ...body],
    { x: MX, y: 1.55, w: 8.1, colW: [4.9, 1.1, 2.1], rowH: 0.47, fontSize: 14, color: INK, border: { type: "solid", pt: 0.75, color: H.accent4 }, margin: [0, 0.12, 0, 0.12], objectName: nm("table") });
  T(s, "’26년 값은 10.31. 기준, ‘-’는 신규 지표", MX, 5.9, 8.1, 0.3, { fontSize: 11, color: MUTED });
  const x = 9.0, w = W - MX - x;
  box(s, x, 1.55, w, 2.35, NAVY);
  T(s, "합동평가 산식(100점)", x + 0.25, 1.72, w - 0.5, 0.35, { fontSize: 14, bold: true, color: WHITE });
  T(s, [
    { text: "발굴률 70점", options: { bold: true, breakLine: true } },
    { text: "배치율 100%면 1.1 보정", options: { fontSize: 12, breakLine: true } },
    { text: "보건소 참석률 10점", options: { bold: true, breakLine: true } },
    { text: "보건소 서비스 이용률 20점", options: { bold: true } },
  ], x + 0.25, 2.15, w - 0.5, 1.6, { fontSize: 14, color: WHITE, paraSpaceAfter: 5 });
  box(s, x, 4.1, w, 2.55, TINT);
  T(s, "자체 지표", x + 0.25, 4.27, w - 0.5, 0.35, { fontSize: 14, bold: true, color: NAVY });
  T(s, [
    { text: "모니터링 결과 반영 90%", options: { bullet: true, breakLine: true } },
    { text: "자체교육·설명회 각 연 6회", options: { bullet: true, breakLine: true } },
    { text: "만족도 조사 연 300명", options: { bullet: true, breakLine: true } },
    { text: "통합지원회의 월 2회 이상", options: { bullet: true } },
  ], x + 0.25, 4.7, w - 0.5, 1.8, { fontSize: 13, paraSpaceAfter: 5 });
}

// ================= 9. 추진체계 =================
{
  const s = slide(S3, "추진체계: 동별 2인 체계로 발굴부터 모니터링까지", "복지정책과 통합돌봄팀이 총괄하고, 6개 동에 복지 전담 1명과 간호 겸직 1명을 둡니다. 복지 전담인력이 발굴부터 연계, 모니터링까지 맡고, 간호직은 의료 필요가 큰 대상자의 사전조사에 동행해 의견을 냅니다.");
  fig(s, "의왕시_통합돌봄_추진체계도.png", MX, 1.5, 8.3, 3280 / 1914, "의왕시 통합돌봄 추진체계도: 복지정책과 통합돌봄팀, 6개 동 주민센터, 전문기관, 서비스 제공기관, 협업 부서");
  const x = 9.15, w = W - MX - x;
  const pts = [
    ["전담인력(복지) 1명", "발굴 · 조사 · 지원계획 · 연계 · 모니터링 총괄"],
    ["겸직인력(간호) 1명", "의료 필요 대상자 사전조사 동행 · 의견 제시"],
    ["변경 절차", "전담인력 모니터링 → 변경(안) 작성 → 통합지원회의 재상정"],
    ["인력 16명", "전담 10명(배치율 100%) + 겸직 6명"],
  ];
  pts.forEach((p, i) => {
    const y = 1.5 + i * 1.3;
    box(s, x, y, w, 1.13, i === 3 ? NAVY : TINT);
    T(s, p[0], x + 0.2, y + 0.13, w - 0.4, 0.35, { fontSize: 14, bold: true, color: i === 3 ? WHITE : NAVY });
    T(s, p[1], x + 0.2, y + 0.5, w - 0.4, 0.55, { fontSize: 12, color: i === 3 ? PALE : INK });
  });
}

// ================= 10. 발굴 =================
{
  const s = slide(S3, "대상자 발굴: 먼저 찾아가는 선제발굴로 전환", "신청을 기다리지 않고 먼저 찾아갑니다. 초고령 위험군을 동별로 연중 조사하고, 건보 빅데이터 명단을 반기마다 동에 배분하며, 동별 설명회를 엽니다.");
  const cs = [
    ["초고령 위험군 선제발굴", "기초연금 수급자 중 85세 이상 1인가구 명단 추출(1~2월) 후 동별 연중 조사", "확대", BLUE],
    ["건보 빅데이터 활용", "건보공단 대상자 명단을 반기 1회 받아 동별 배분, 매월 행복e음 위기정보 확인", "확대", BLUE],
    ["찾아가는 통합돌봄 설명회", "동별 연 1회 이상, 총 6회. 주민 생활 밀착 홍보와 병행(홍보자료 연 1,000건)", "신규", ORANGE],
  ];
  const cw = (CW - 0.6) / 3;
  cs.forEach((c, i) => {
    const x = MX + i * (cw + 0.3);
    box(s, x, 1.55, cw, 2.2, TINT);
    num(s, String(i + 1), x + 0.25, 1.75, 0.45);
    T(s, c[0], x + 0.85, 1.75, cw - 1.85, 0.45, { fontSize: 15, bold: true, color: NAVY, valign: "middle" });
    badge(s, c[2], x + cw - 0.9, 1.83, 0.65, c[3]);
    T(s, c[1], x + 0.25, 2.4, cw - 0.5, 1.2, { fontSize: 13 });
  });
  fig(s, "의왕시_연간발굴로드맵.png", MX, 4.0, CW, 2480 / 540, "연간 발굴 로드맵: 1~2월 명단 정리, 3~11월 동별 연중 조사, 건보 빅데이터 상·하반기 수신, 매월 실적 점검");
}

// ================= 11. 퇴원환자 =================
{
  const s = slide(S3, "퇴원환자 연계: 병원 유형에 따라 의뢰 경로 이원화", "상급종합병원은 퇴원일이 전날 확정되어 사전조사가 어려우므로 안내문 배부 후 동 신청으로, 종합병원은 병원이 시에 직접 의뢰하는 방식으로 나눕니다.");
  fig(s, "의왕시_퇴원환자_의뢰경로이원화.png", MX, 1.5, CW, 2480 / 618, "퇴원환자 의뢰 경로 이원화: 상급종합병원은 안내문 배부 후 동 신청, 종합병원은 시 전담조직에 직접 의뢰");
  const cs = [
    ["목표 60명", "퇴원환자 통합지원 대상자(’27년)"],
    ["협약병원 확대", "’26. 3개소 → ’27. 확대, 연계수당 건당 5만원 이내"],
    ["퇴원 직후 단기돌봄", "회복살핌 꾸러미 + 노인맞춤돌봄 퇴원환자 단기집중, 미해당 시 일상살핌"],
  ];
  const cw = (CW - 0.6) / 3;
  cs.forEach((c, i) => {
    const x = MX + i * (cw + 0.3);
    box(s, x, 4.85, cw, 1.75, i === 0 ? NAVY : TINT);
    T(s, c[0], x + 0.25, 5.03, cw - 0.5, 0.55, { fontSize: i === 0 ? 28 : 16, bold: true, color: i === 0 ? WHITE : NAVY, valign: "middle" });
    T(s, c[1], x + 0.25, 5.68, cw - 0.5, 0.8, { fontSize: 13, color: i === 0 ? PALE : INK });
  });
}

// ================= 12. 서비스 =================
{
  const s = slide(S3, "서비스: 의왕살핌 고도화와 돌봄공백 신규서비스", "의왕살핌은 식사지원을 퇴원·급성기 치료식 맞춤형 영양식으로 개편하고, 이동살핌·회복살핌 패키지·AI 돌봄로봇 3종을 새로 도입합니다. 방문진료는 기관 확충보다 이용 활성화에 집중합니다.");
  const lw = 5.6;
  box(s, MX, 1.55, lw, 5.1, TINT);
  T(s, "의왕살핌 고도화", MX + 0.3, 1.75, 3.5, 0.4, { fontSize: 17, bold: true, color: NAVY, valign: "middle" });
  badge(s, "확대", MX + lw - 1.0, 1.8, 0.7, BLUE);
  const sv = [["일상살핌", "125명"], ["내집살핌", "95명"], ["건강살핌", "110명"]];
  const tw = (lw - 0.6 - 0.4) / 3;
  sv.forEach((v, i) => {
    const x = MX + 0.3 + i * (tw + 0.2);
    box(s, x, 2.35, tw, 1.35, WHITE);
    T(s, v[1], x, 2.45, tw, 0.7, { fontSize: 28, bold: true, color: NAVY, align: "center", valign: "middle" });
    T(s, v[0], x, 3.15, tw, 0.4, { fontSize: 13, color: MUTED, align: "center" });
  });
  T(s, [
    { text: "식사지원 개편: 일반식은 기존 서비스로 연계, 퇴원·급성기 치료식 맞춤형 영양식 제공", options: { bullet: true, breakLine: true } },
    { text: "예산 514,136천원(’27년)", options: { bullet: true, breakLine: true } },
    { text: "방문진료 참여의원 14개소·재택의료센터는 대상자 발굴로 이용 활성화", options: { bullet: true } },
  ], MX + 0.3, 3.95, lw - 0.6, 2.5, { fontSize: 14, paraSpaceAfter: 8 });
  const x = MX + lw + 0.3, w = CW - lw - 0.3;
  const nw = [
    ["이동살핌", "병원 이동이 어려운 대상자 왕복 이동 지원(34명, 1인 4회)", "20,000천원"],
    ["회복살핌 패키지", "퇴원·급성기 회복기 단기돌봄 꾸러미", "10,000천원"],
    ["AI 돌봄로봇", "독거·고위험 대상자 정서·안전 돌봄", "80,000천원"],
  ];
  nw.forEach((v, i) => {
    const y = 1.55 + i * 1.75;
    box(s, x, y, w, 1.6, WHITE, { line: MID });
    badge(s, "신규", x + 0.25, y + 0.25, 0.65);
    T(s, v[0], x + 1.05, y + 0.18, w - 3.3, 0.45, { fontSize: 17, bold: true, color: NAVY, valign: "middle" });
    T(s, v[2], x + w - 2.25, y + 0.18, 2.0, 0.45, { fontSize: 15, bold: true, color: ORANGE, align: "right", valign: "middle" });
    T(s, v[1], x + 0.25, y + 0.8, w - 0.5, 0.65, { fontSize: 14 });
  });
}

// ================= 13. 예산 =================
pres.addSection({ title: S4 });
{
  const s = slide(S4, "2027 예산: 사업비 689,636천원", "사업비는 6억 8,964만원으로 국비 50, 도비 5, 시비 45퍼센트입니다. 의왕살핌이 75퍼센트를 차지하고, 신규서비스 3종에 1억 1천만원을 편성했습니다.");
  T(s, "재원별(천원)", MX, 1.55, 4, 0.35, { fontSize: 16, bold: true, color: NAVY });
  s.addChart(pres.charts.DOUGHNUT, [{ name: "재원", labels: ["국비 50%", "시비 45%", "도비 5%"], values: [344818, 310336, 34482] }],
    { x: MX - 0.1, y: 1.9, w: 4.6, h: 3.5, holeSize: 58, chartColors: [H.dk2, H.accent3, H.accent2], showPercent: false, showValue: false, showLegend: true, legendPos: "b",
      legendFontSize: 12, legendColor: H.dk1, legendFontFace: "+mn-lt", dataBorder: { pt: 1, color: H.lt1 }, objectName: nm("chart") });
  box(s, MX, 5.55, 4.4, 1.1, TINT);
  T(s, "국 344,818 · 도 34,482 · 시 310,336", MX + 0.2, 5.62, 4.0, 0.45, { fontSize: 13, bold: true, color: NAVY, valign: "middle" });
  T(s, "운영비 17,000천원 별도(회의·홍보·행사)", MX + 0.2, 6.07, 4.0, 0.45, { fontSize: 12, color: MUTED, valign: "middle" });
  const hd = (t) => ({ text: t, options: { bold: true, color: WHITE, fill: { color: NAVY }, align: "center", valign: "middle" } });
  const rows = [
    ["의왕살핌(일상·건강·내집)", "514,136", ""], ["AI 돌봄로봇", "80,000", "신규"], ["의왕형 통합돌봄(보조금사업)", "60,000", ""],
    ["이동살핌", "20,000", "신규"], ["회복살핌 패키지", "10,000", "신규"], ["재택의료센터 운영지원", "5,000", ""], ["퇴원환자 연계수당", "500", "신규"],
  ];
  const body = rows.map((r, i) => r.map((c, j) => ({ text: c, options: { align: j === 0 ? "left" : j === 1 ? "right" : "center", valign: "middle", bold: j === 2, color: j === 2 ? ORANGE : INK, fill: { color: i % 2 ? TINT : WHITE } } })));
  const tot = ["사업비 계", "689,636", ""].map((c, j) => ({ text: c, options: { bold: true, color: NAVY, fill: { color: PALE }, align: j === 0 ? "left" : "right", valign: "middle" } }));
  s.addTable([[hd("사업"), hd("예산(천원)"), hd("비고")], ...body, tot],
    { x: 5.4, y: 1.55, w: 7.33, colW: [4.2, 1.9, 1.23], rowH: 0.52, fontSize: 14, color: INK, border: { type: "solid", pt: 0.75, color: H.accent4 }, margin: [0, 0.15, 0, 0.15], objectName: nm("table") });
}

// ================= 14. 심의 요청 =================
{
  const s = pres.addSlide({ masterName: "CLOSING", sectionTitle: S4 });
  s.addText("심의 요청 사항", { placeholder: "title" });
  s.addNotes("이상 네 가지에 대해 위원님들의 심의와 의견을 부탁드립니다. 주신 의견은 계획에 반영해 확정하겠습니다.");
  const it = [
    ["목표 대상자", "통합지원 대상자 541명(우선발굴필요군 421명 + 기타 120명)"],
    ["성과지표", "발굴률 20% 이상, 우선발굴필요군 중 대상자 비중 6% 이상 등"],
    ["추진과제", "4개 추진과제, 12개 핵심사업(선제발굴 · 퇴원환자 이원화 · 신규서비스)"],
    ["예산", "사업비 689,636천원(국 50% · 도 5% · 시 45%)"],
  ];
  it.forEach((v, i) => {
    const y = 1.75 + i * 1.02;
    box(s, MX, y, CW, 0.85, BLUE);
    num(s, String(i + 1), MX + 0.25, y + 0.2, 0.45, ORANGE);
    T(s, v[0], MX + 0.95, y, 2.0, 0.85, { fontSize: 17, bold: true, color: WHITE, valign: "middle" });
    T(s, v[1], MX + 3.0, y, CW - 3.2, 0.85, { fontSize: 15, color: WHITE, valign: "middle" });
  });
  T(s, "향후 절차", MX, 6.0, 1.5, 0.4, { fontSize: 14, bold: true, color: ORANGE, valign: "middle" });
  T(s, "협의체 심의(11.26.) → 심의 의견 반영 → 지역계획 확정·제출 → 2027. 1. 시행", MX + 1.5, 6.0, CW - 1.5, 0.4, { fontSize: 14, color: PALE, valign: "middle" });
  T(s, "감사합니다", MX, 6.6, CW, 0.4, { fontSize: 14, color: PALE });
}

(async () => {
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT);
})();
