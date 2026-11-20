// 통합지원협의체 심의 발표자료(PPT) 생성: node build.js <out.pptx>
// 필요: pptxgenjs, pptx 스킬의 scripts/apply_theme.js (환경변수 APPLY_THEME 로 경로 지정)
// 내용 기준: 2027 지역계획 최종본('26.11.20. 업로드본)
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
pres.title = "2027년 의왕시 통합지원 지역계획(안)";
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
    { placeholder: { options: { name: "body", type: "body", x: MX, y: 4.75, w: 9.6, h: 0.9, fontSize: 20, color: PALE, valign: "top", margin: 0 }, text: "" } },
  ],
});
pres.defineSlideMaster({
  title: "DIVIDER", background: { color: H.dk2 },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: MX + 1.6, y: 2.9, w: 10, h: 1.2, fontSize: 36, bold: true, color: WHITE, valign: "middle", align: "left", margin: 0 }, text: "" } },
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
  s.addShape(pres.ShapeType.roundRect, { x, y, w, h, fill: { color: fill }, line: o.line ? { color: o.line, width: o.lw || 1 } : { type: "none" }, rectRadius: o.r ?? 0.12, objectName: nm("card") });
}
function T(s, text, x, y, w, h, o = {}) {
  s.addText(text, { x, y, w, h, isTextBox: true, margin: 0, fontSize: 14, color: INK, valign: "top", objectName: nm("text"), ...o });
}
function dot(s, label, x, y, d = 0.5, fill = NAVY, fs = 14) {
  s.addShape(pres.ShapeType.ellipse, { x, y, w: d, h: d, fill: { color: fill }, line: { type: "none" }, objectName: nm("dot") });
  T(s, label, x, y, d, d, { fontSize: fs, bold: true, color: WHITE, align: "center", valign: "middle" });
}
function badge(s, label, x, y, w = 0.7, fill = ORANGE, fs = 11) {
  box(s, x, y, w, 0.3, fill, { r: 0.06 });
  T(s, label, x, y, w, 0.3, { fontSize: fs, bold: true, color: WHITE, align: "center", valign: "middle" });
}
function arrowR(s, x, y, w = 0.4, h = 0.36, fill = ORANGE) {
  s.addShape(pres.ShapeType.rightArrow, { x, y, w, h, fill: { color: fill }, line: { type: "none" }, objectName: nm("arrow") });
}
function arrowD(s, x, y, w = 0.36, h = 0.3, fill = MID) {
  s.addShape(pres.ShapeType.downArrow, { x, y, w, h, fill: { color: fill }, line: { type: "none" }, objectName: nm("arrow") });
}
function slide(section, title, notes) {
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: section });
  s.addText(title, { placeholder: "title" });
  if (notes) s.addNotes(notes);
  return s;
}
function divider(section, num, title) {
  const s = pres.addSlide({ masterName: "DIVIDER", sectionTitle: section });
  T(s, num, MX, 2.9, 1.3, 1.2, { fontSize: 54, bold: true, color: ORANGE, valign: "middle" });
  s.addText(title, { placeholder: "title" });
  return s;
}
function fig(s, file, x, y, w, ratio, alt) {
  s.addImage({ path: path.join(FIG, file), x, y, w, h: w / ratio, altText: alt, objectName: nm("figure") });
}
const chartText = { catAxisLabelFontFace: "+mn-lt", valAxisLabelFontFace: "+mn-lt", dataLabelFontFace: "+mn-lt", legendFontFace: "+mn-lt",
  catAxisLabelColor: H.dk1, valAxisLabelColor: H.accent5, dataLabelColor: H.dk1, catAxisLabelFontSize: 13, valAxisLabelFontSize: 11, dataLabelFontSize: 13 };
function stat(s, x, y, w, h, big, label, o = {}) {
  box(s, x, y, w, h, o.fill || TINT);
  T(s, big, x + 0.25, y + 0.15, w - 0.5, h * 0.52, { fontSize: o.fs || 40, bold: true, color: o.color || NAVY, valign: "middle" });
  T(s, label, x + 0.25, y + h * 0.62, w - 0.5, h * 0.32, { fontSize: 14, color: o.lc || INK, bold: !!o.lb });
}

const S1 = "개요", S2 = "2026 추진 성과", S3 = "지역 진단", S4 = "2027 목표", S5 = "추진과제", S6 = "예산·심의";

// ================= 1. 표지 =================
pres.addSection({ title: S1 });
{
  const s = pres.addSlide({ masterName: "COVER", sectionTitle: S1 });
  badge(s, "심의 안건", MX, 1.9, 1.1);
  s.addText("2027년 의왕시\n통합지원 지역계획(안)", { placeholder: "title" });
  s.addText("살던 곳에서 건강하게, 함께 살피는 의왕", { placeholder: "body" });
  T(s, "제2차 의왕시 통합지원협의체  |  2026. 11. 26.", MX, 6.3, 8, 0.35, { fontSize: 14, color: PALE });
  T(s, "복지정책과 통합돌봄팀", MX, 6.68, 8, 0.35, { fontSize: 14, color: PALE });
  s.addShape(pres.ShapeType.ellipse, { x: 10.3, y: 1.2, w: 5.2, h: 5.2, fill: { color: BLUE }, line: { type: "none" }, objectName: nm("deco") });
  s.addShape(pres.ShapeType.ellipse, { x: 11.5, y: 4.3, w: 2.6, h: 2.6, fill: { color: ORANGE }, line: { type: "none" }, objectName: nm("deco") });
  s.addNotes("2027년 의왕시 통합지원 지역계획(안)을 보고드리고 심의를 요청드립니다.");
}

// ================= 2. 순서 =================
{
  const s = slide(S1, "보고 순서", "2026년 성과와 지역 진단을 먼저 말씀드리고, 2027년 목표와 추진과제, 예산 순으로 보고드립니다.");
  const steps = [["2026 추진 성과", "무엇을 했고, 무엇이 달라졌나"], ["지역 진단", "우리 지역 돌봄 수요와 과제"], ["2027 목표", "비전 · 대상자 · 성과지표"], ["추진과제", "발굴 · 퇴원 · 모니터링 · 서비스"], ["예산 · 심의 요청", "사업비 689,636천원"]];
  steps.forEach((st, i) => {
    const y = 1.55 + i * 1.03;
    dot(s, String(i + 1), MX + 0.2, y + 0.12, 0.62, i % 2 ? BLUE : NAVY, 18);
    box(s, MX + 1.1, y, CW - 1.1, 0.86, TINT);
    T(s, st[0], MX + 1.4, y, 4.2, 0.86, { fontSize: 20, bold: true, color: NAVY, valign: "middle" });
    T(s, st[1], MX + 5.6, y, CW - 5.8, 0.86, { fontSize: 15, color: MUTED, valign: "middle" });
  });
}

// ================= 3. 2026 한눈에 =================
pres.addSection({ title: S2 });
divider(S2, "01", "2026년 추진 성과");
{
  const s = slide(S2, "2026년 한눈에 보기", "법 시행 첫해, 390명이 신청해 276명을 통합지원 대상자로 지원했습니다. 1인당 2.7건, 총 740건의 서비스를 연계했고 의왕살핌은 340명이 이용했습니다.");
  const cw = (CW - 3 * 0.3) / 4;
  const st = [["390명", "통합돌봄 신청"], ["276명", "통합지원 대상자"], ["740건", "서비스 연계 건수"], ["340명", "의왕살핌 이용"]];
  st.forEach((v, i) => stat(s, MX + i * (cw + 0.3), 1.6, cw, 1.9, v[0], v[1], { fs: 44, color: i === 3 ? ORANGE : NAVY, lb: true }));
  // 흐름: 신청 → 대상자 → 연계
  T(s, "운영 기반", MX, 3.9, 4, 0.4, { fontSize: 16, bold: true, color: NAVY });
  const base = [["통합돌봄팀 신설", "’26. 2."], ["조례 제정", "’26. 3."], ["통합지원회의", "15회 · 300건 심의"], ["동 전담인력", "’26.11.30. 6명 배치"]];
  base.forEach((b, i) => {
    const x = MX + i * (cw + 0.3);
    box(s, x, 4.4, cw, 1.5, WHITE, { line: PALE, lw: 1.5 });
    T(s, b[0], x + 0.2, 4.6, cw - 0.4, 0.5, { fontSize: 17, bold: true, color: NAVY, align: "center" });
    T(s, b[1], x + 0.2, 5.15, cw - 0.4, 0.45, { fontSize: 14, color: MUTED, align: "center" });
    if (i < 3) arrowR(s, x + cw + 0.02, 4.98, 0.26, 0.32, MID);
  });
  T(s, "※ ’26. 10. 31. 기준", MX, 6.25, 5, 0.3, { fontSize: 11, color: MUTED });
}

// ================= 4. 발굴 변화 =================
{
  const s = slide(S2, "발굴 경로를 넓히자 신청이 늘었습니다", "본인 신청에 의존하던 초기와 달리, 건보 빅데이터·장기요양기관·사회단체·병원 안내로 경로를 넓히고 동별 목표를 관리했습니다. 10월 한 달 147명이 신청해 월평균의 5배가 넘었고, 실적이 낮던 동이 크게 올라왔습니다.");
  // 왼쪽: 경로 다이어그램
  const paths = ["건보 빅데이터", "장기요양기관", "사회단체 설명", "병원 안내"];
  paths.forEach((p, i) => {
    const y = 1.6 + i * 0.85;
    box(s, MX, y, 2.6, 0.65, TINT);
    T(s, p, MX, y, 2.6, 0.65, { fontSize: 15, bold: true, color: NAVY, align: "center", valign: "middle" });
    arrowR(s, MX + 2.7, y + 0.15, 0.4, 0.35, MID);
  });
  box(s, MX + 3.25, 1.6, 2.3, 3.2, NAVY);
  T(s, "10월\n신청", MX + 3.25, 1.85, 2.3, 0.9, { fontSize: 16, color: PALE, align: "center", bold: true });
  T(s, "147명", MX + 3.25, 2.75, 2.3, 0.9, { fontSize: 38, bold: true, color: WHITE, align: "center", valign: "middle" });
  T(s, "1~9월 월평균\n27명의 5배+", MX + 3.25, 3.75, 2.3, 0.8, { fontSize: 13, color: PALE, align: "center" });
  // 오른쪽: 1만명당 신청 차트
  T(s, "노인 1만 명당 신청(명)", 6.6, 1.5, 6, 0.4, { fontSize: 16, bold: true, color: NAVY });
  s.addChart(pres.charts.BAR, [
    { name: "’26. 9.", labels: ["내손2동", "청계동", "오전동", "의왕시"], values: [57.9, 64.4, 64.5, 74.8] },
    { name: "’26. 10.", labels: ["내손2동", "청계동", "오전동", "의왕시"], values: [207.2, 91.7, 88.3, 119.4] },
  ], { x: 6.5, y: 1.9, w: 6.25, h: 4.0, barDir: "col", barGapWidthPct: 55, chartColors: [H.accent3, H.accent2],
    showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: "0.0", showLegend: true, legendPos: "t", legendFontSize: 13, legendColor: H.dk1,
    valAxisHidden: true, valAxisMinVal: 0, valAxisMaxVal: 240, valGridLine: { style: "none" }, catGridLine: { style: "none" }, ...chartText, objectName: nm("chart") });
  box(s, MX, 5.1, 4.95, 0.85, PALE);
  T(s, "퇴원환자: 병원 안내 전환 후 43명 선정", MX + 0.2, 5.1, 4.6, 0.85, { fontSize: 14, bold: true, color: NAVY, valign: "middle" });
}

// ================= 5. 서비스 인프라 =================
{
  const s = slide(S2, "없던 재가의료 서비스가 생겼습니다", "실태조사에서 방문진료·재택의료·방문간호는 관내 공급이 없었습니다. 의왕살핌 건강살핌으로 등록만 돼 있던 의원 12개소가 실제 방문진료를 시작했고, 재택의료센터도 관내·관외 각 1개소를 확보했습니다.");
  const rows = [["방문진료", "무공급", "참여의원 12개소", "의과 3 · 한의 9"], ["재택의료센터", "무공급", "관내 1 · 관외 1", "한센복지협회 부설의원 등"], ["방문간호(가정간호)", "무공급", "관내 1개소", "건강살핌 연계"]];
  T(s, "’25. 실태조사", MX + 2.6, 1.5, 2.2, 0.35, { fontSize: 13, bold: true, color: MUTED, align: "center" });
  T(s, "’26. 현재", MX + 5.4, 1.5, 2.2, 0.35, { fontSize: 13, bold: true, color: ORANGE, align: "center" });
  rows.forEach((r, i) => {
    const y = 1.95 + i * 1.05;
    T(s, r[0], MX, y, 2.5, 0.8, { fontSize: 16, bold: true, color: NAVY, valign: "middle" });
    box(s, MX + 2.6, y, 2.2, 0.8, "F2F2F2");
    T(s, r[1], MX + 2.6, y, 2.2, 0.8, { fontSize: 16, color: MUTED, align: "center", valign: "middle" });
    arrowR(s, MX + 4.88, y + 0.22, 0.42, 0.36);
    box(s, MX + 5.4, y, 2.2, 0.8, NAVY);
    T(s, r[2], MX + 5.4, y + 0.06, 2.2, 0.42, { fontSize: 15, bold: true, color: WHITE, align: "center", valign: "middle" });
    T(s, r[3], MX + 5.4, y + 0.44, 2.2, 0.3, { fontSize: 10.5, color: PALE, align: "center", valign: "middle" });
  });
  // 오른쪽 의왕살핌
  const x = 8.9, w = W - MX - x;
  box(s, x, 1.5, w, 4.45, TINT);
  T(s, "의왕살핌 3대 패키지", x + 0.25, 1.65, w - 0.5, 0.4, { fontSize: 16, bold: true, color: NAVY });
  T(s, "340명 이용", x + 0.25, 2.05, w - 0.5, 0.55, { fontSize: 28, bold: true, color: ORANGE });
  s.addChart(pres.charts.BAR, [{ name: "이용(명)", labels: ["일상살핌", "건강살핌", "내집살핌"], values: [138, 103, 99] }],
    { x: x + 0.1, y: 2.7, w: w - 0.2, h: 3.1, barDir: "bar", barGapWidthPct: 45, chartColors: [H.accent1], showValue: true, dataLabelPosition: "outEnd",
      showLegend: false, valAxisHidden: true, valAxisMaxVal: 175, valGridLine: { style: "none" }, catGridLine: { style: "none" }, ...chartText, objectName: nm("chart") });
  box(s, MX, 5.25, 7.6, 0.7, PALE);
  T(s, "특화사업 3종 신규 개시: 영양식 · 건강돌봄가 · 동고동락(127명)", MX + 0.2, 5.25, 7.3, 0.7, { fontSize: 13.5, bold: true, color: NAVY, valign: "middle" });
}

// ================= 6. 지역 여건 =================
pres.addSection({ title: S3 });
divider(S3, "02", "지역 진단");
{
  const s = slide(S3, "초고령 1인 가구가 부곡동에 몰려 있습니다", "노인인구는 5년 새 43.5% 늘어 초고령사회에 들어섰습니다. 85세 이상 1인세대 1,046세대 가운데 부곡동이 292세대로 가장 많고, 노인 수 대비로는 부곡동과 내손1동이 시 평균보다 높습니다.");
  const st = [["33,930명", "노인인구 · 고령화율 20.3%", NAVY, WHITE, PALE], ["+43.5%", "’21.12. 대비 노인 증가", TINT, NAVY, MUTED], ["1,046세대", "85세 이상 1인세대", TINT, ORANGE, MUTED]];
  st.forEach((v, i) => {
    const y = 1.55 + i * 1.5;
    box(s, MX, y, 4.2, 1.32, v[2]);
    T(s, v[0], MX + 0.3, y + 0.12, 3.6, 0.7, { fontSize: 34, bold: true, color: v[3], valign: "middle" });
    T(s, v[1], MX + 0.3, y + 0.85, 3.6, 0.35, { fontSize: 13.5, color: v[4] });
  });
  T(s, "동별 85세 이상 1인세대(세대)", 5.3, 1.5, 6, 0.4, { fontSize: 16, bold: true, color: NAVY });
  s.addChart(pres.charts.BAR, [{ name: "85세 이상 1인세대", labels: ["부곡동", "청계동", "내손1동", "오전동", "내손2동", "고천동"], values: [292, 211, 177, 156, 114, 96] }],
    { x: 5.2, y: 1.9, w: 7.55, h: 3.6, barDir: "col", barGapWidthPct: 45, chartColors: [H.accent2, H.accent1, H.accent1, H.accent1, H.accent1, H.accent1], invertedColors: [H.accent2],
      showValue: true, dataLabelPosition: "outEnd", showLegend: false, valAxisHidden: true, valAxisMaxVal: 340, valGridLine: { style: "none" }, catGridLine: { style: "none" }, ...chartText, objectName: nm("chart") });
  box(s, 5.3, 5.6, 7.43, 0.4 + 0.0, WHITE);
  T(s, "부곡동: 노인 기초생활 수급률 12.0% · 단독‧연립‧다세대 32.8%로 시 최고", 5.3, 5.55, 7.43, 0.45, { fontSize: 13.5, bold: true, color: NAVY, valign: "middle" });
}

// ================= 7. 진단 → 대응 =================
{
  const s = slide(S3, "진단에서 2027 과제로", "다섯 가지 진단 결과를 2027년 과제로 연결했습니다. 신청을 기다리지 않는 발굴, 방문의료 이용 확대, 퇴원환자 의뢰 경로 개선, 서비스 개시 확인, 주민 인지도 제고입니다.");
  T(s, "지역의 문제", MX, 1.45, 5.4, 0.35, { fontSize: 14, bold: true, color: MUTED });
  T(s, "2027 대응", 7.3, 1.45, 5.4, 0.35, { fontSize: 14, bold: true, color: MUTED });
  const rows = [
    ["48.6%", "핵심군 중 85세 이상(경기 3위)\n신청주의로는 발굴 한계", "초고령 위험군 선제발굴"],
    ["10.5%", "방문의료 수요 대비 이용\n(668명 중 70명)", "방문의료 이용 활성화"],
    ["0건", "협약병원 공식 의뢰\n(퇴원일 전날 확정)", "퇴원환자 의뢰 경로 이원화"],
    ["지연", "의뢰 증가로\n서비스 개시 확인 늦어짐", "개인별지원계획 수립‧모니터링 내실화"],
    ["6.4%", "돌봄 인지 52.9% → 이용 6.4%", "찾아가는 통합돌봄 설명회"],
  ];
  rows.forEach((r, i) => {
    const y = 1.85 + i * 0.98;
    box(s, MX, y, 5.95, 0.84, TINT);
    T(s, r[0], MX + 0.2, y, 1.6, 0.84, { fontSize: r[0].length > 3 ? 22 : 26, bold: true, color: NAVY, valign: "middle" });
    T(s, r[1], MX + 1.9, y, 3.95, 0.84, { fontSize: 13, valign: "middle" });
    arrowR(s, 6.67, y + 0.24, 0.45, 0.36);
    box(s, 7.3, y, CW - 6.7, 0.84, NAVY);
    T(s, r[2], 7.55, y, CW - 7.2, 0.84, { fontSize: 16, bold: true, color: WHITE, valign: "middle" });
  });
}

// ================= 8. 비전 =================
pres.addSection({ title: S4 });
divider(S4, "03", "2027년 목표");
{
  const s = slide(S4, "비전과 추진체계", "비전은 ‘살던 곳에서 건강하게, 함께 살피는 의왕’입니다. 3대 목표 아래 4개 추진과업, 12개 핵심사업으로 구성했습니다.");
  fig(s, "의왕시_비전체계도.png", 0.97, 1.35, 11.4, 2280 / 1144, "2027 의왕시 통합지원 비전 체계도: 비전, 3대 목표, 4개 추진과업과 12개 핵심사업");
}

// ================= 9. 목표 대상자 =================
{
  const s = slide(S4, "2027년 통합지원 대상자 541명", "2027년 목표 대상자는 541명으로 올해 276명의 약 2배입니다. 우선발굴필요군 6,506명 가운데 421명, 6.5%를 지원합니다.");
  box(s, MX, 1.55, 4.3, 2.35, NAVY);
  T(s, "541명", MX + 0.3, 1.75, 3.7, 1.1, { fontSize: 56, bold: true, color: WHITE, valign: "middle" });
  T(s, "’26년 276명 → 약 2배", MX + 0.3, 2.95, 3.7, 0.4, { fontSize: 15, color: PALE });
  // 구성 막대
  const total = 541, bw = 4.3, y0 = 4.15;
  const a = bw * 421 / total;
  box(s, MX, y0, a, 0.55, BLUE, { r: 0.05 }); box(s, MX + a + 0.04, y0, bw - a - 0.04, 0.55, ORANGE, { r: 0.05 });
  T(s, "우선발굴필요군 421", MX, y0, a, 0.55, { fontSize: 13, bold: true, color: WHITE, align: "center", valign: "middle" });
  T(s, "기타 120", MX + a + 0.04, y0, bw - a - 0.04, 0.55, { fontSize: 12, bold: true, color: WHITE, align: "center", valign: "middle" });
  box(s, MX, 4.95, 4.3, 1.0, TINT);
  T(s, "421 ÷ 우선발굴필요군 6,506명 = 6.5%", MX + 0.25, 4.95, 3.9, 1.0, { fontSize: 14, bold: true, color: NAVY, valign: "middle" });
  T(s, "유형별 목표(명)", 5.3, 1.5, 5, 0.4, { fontSize: 16, bold: true, color: NAVY });
  const lab = ["재가급여 수급자", "고령 장애인", "퇴원환자", "등급외·노인맞춤돌봄", "장기요양 기각·각하", "65세 미만 중증장애인", "시설 퇴소자", "기타(우선발굴필요군 외)"];
  const val = [145, 107, 60, 50, 40, 17, 2, 120];
  s.addChart(pres.charts.BAR, [{ name: "목표", labels: lab.slice().reverse(), values: val.slice().reverse() }],
    { x: 5.2, y: 1.9, w: 7.55, h: 4.3, barDir: "bar", barGapWidthPct: 40, chartColors: [H.accent1], showValue: true, dataLabelPosition: "outEnd",
      showLegend: false, valAxisHidden: true, valAxisMaxVal: 170, valGridLine: { style: "none" }, catGridLine: { style: "none" }, ...chartText, objectName: nm("chart") });
}

// ================= 10. 성과지표 =================
{
  const s = slide(S4, "핵심 성과지표", "중앙 지표는 행안부 합동평가와 복지부 평가 기준을 따랐습니다. 가장 중요한 지표는 직권·자체발굴 발굴률 20%입니다.");
  const k = [["20%", "발굴률", "직권·자체발굴 / 신청", ORANGE], ["6%", "지원계획 수립률", "우선발굴필요군 6,506명 대비", BLUE], ["80%", "모니터링 수행률", "3개월 단위 수행", BLUE], ["25%", "보건소 서비스 이용률", "2가지 이상 연계", BLUE]];
  const cw = (CW - 3 * 0.3) / 4;
  k.forEach((v, i) => {
    const x = MX + i * (cw + 0.3);
    box(s, x, 1.6, cw, 2.6, i === 0 ? NAVY : TINT);
    T(s, v[0], x, 1.8, cw, 1.0, { fontSize: 50, bold: true, color: i === 0 ? WHITE : NAVY, align: "center", valign: "middle" });
    T(s, v[1], x + 0.15, 2.9, cw - 0.3, 0.45, { fontSize: 17, bold: true, color: i === 0 ? WHITE : NAVY, align: "center" });
    T(s, v[2], x + 0.15, 3.4, cw - 0.3, 0.4, { fontSize: 12.5, color: i === 0 ? PALE : MUTED, align: "center" });
    if (i === 0) badge(s, "합동평가 70점", x + cw / 2 - 0.7, 3.85, 1.4);
  });
  T(s, "그 밖의 지표", MX, 4.55, 4, 0.4, { fontSize: 15, bold: true, color: NAVY });
  const others = ["우선발굴필요군 중 대상자 6% 이상", "대상자 중 우선발굴필요군 70% 이상", "전담인력 배치율 100%", "보건소 회의 참석률 90%", "통합지원회의 월 2회 이상", "자체교육 연 5회 · 설명회 연 6회"];
  const ow = (CW - 2 * 0.25) / 3;
  others.forEach((o, i) => {
    const x = MX + (i % 3) * (ow + 0.25), y = 5.0 + Math.floor(i / 3) * 0.72;
    box(s, x, y, ow, 0.58, WHITE, { line: PALE, lw: 1.5, r: 0.08 });
    T(s, o, x + 0.15, y, ow - 0.3, 0.58, { fontSize: 13.5, color: INK, valign: "middle", align: "center" });
  });
}

// ================= 11. 선제발굴 =================
pres.addSection({ title: S5 });
divider(S5, "04", "2027년 추진과제");
{
  const s = slide(S5, "① 신청을 기다리지 않고 먼저 찾아갑니다", "기초연금 수급자 중 85세 이상 1인가구를 상·하반기 추출해 동에 제공합니다. 노인맞춤돌봄 이용자는 생활지원사가 먼저 선별하고, 서비스를 전혀 받지 않는 분은 동 전담인력이 직접 방문합니다.");
  // 깔때기형 흐름
  const steps = [["명단 추출", "기초연금 수급자 중\n85세 이상 1인가구", NAVY], ["제외", "통합돌봄 신청자·대상자", BLUE], ["선별", "노인맞춤돌봄 이용자\n→ 생활지원사 확인", BLUE], ["방문", "서비스 미이용자\n→ 동 전담인력 방문", ORANGE], ["직권신청", "필요 시 바로 신청", NAVY]];
  const sw = 2.1, gap = 0.35;
  steps.forEach((st, i) => {
    const x = MX + i * (sw + gap);
    box(s, x, 1.6, sw, 0.6, st[2]);
    T(s, st[0], x, 1.6, sw, 0.6, { fontSize: 17, bold: true, color: WHITE, align: "center", valign: "middle" });
    box(s, x, 2.25, sw, 1.05, TINT);
    T(s, st[1], x + 0.1, 2.25, sw - 0.2, 1.05, { fontSize: 13, color: INK, align: "center", valign: "middle" });
    if (i < steps.length - 1) arrowR(s, x + sw + 0.04, 1.73, 0.27, 0.34, MID);
  });
  T(s, "상·하반기 명단 제공 · 연중 조사 · 월별 실적 점검  |  건보 빅데이터 반기 1회 · 행복e음 위기정보 매월", MX, 3.45, CW, 0.4, { fontSize: 13, color: MUTED, align: "center" });
  fig(s, "의왕시_연간발굴로드맵.png", MX, 3.95, CW, 2480 / 540, "연간 발굴 로드맵: 상·하반기 명단 추출, 3~11월 동별 조사, 건보 빅데이터 상·하반기 수신, 매월 실적 점검");
}

// ================= 12. 퇴원환자 =================
{
  const s = slide(S5, "② 퇴원환자 의뢰 경로를 둘로 나눕니다", "상급종합병원은 퇴원일이 전날 확정되어 사전조사가 어렵습니다. 그래서 안내문을 주고 퇴원 후 동에서 신청받습니다. 종합병원은 협약을 넓혀 퇴원 전에 시로 직접 의뢰받고, 건당 5만원 이내 연계수당을 지급합니다.");
  fig(s, "의왕시_퇴원환자_의뢰경로이원화.png", MX, 1.5, CW, 2480 / 618, "퇴원환자 의뢰 경로 이원화: 상급종합병원은 안내문 배부 후 동 신청, 종합병원은 시 전담조직에 직접 의뢰");
  const cs = [["60명", "’27 퇴원환자\n대상자 목표"], ["10명 → 50명", "협약병원 공식 의뢰\n’27 → ’30"], ["5만원", "건당 연계수당\n(협약병원)"]];
  const cw = (CW - 0.6) / 3;
  cs.forEach((c, i) => {
    const x = MX + i * (cw + 0.3);
    box(s, x, 4.75, cw, 1.75, i === 0 ? NAVY : TINT);
    T(s, c[0], x, 4.85, cw, 0.85, { fontSize: 34, bold: true, color: i === 0 ? WHITE : NAVY, align: "center", valign: "middle" });
    T(s, c[1], x, 5.7, cw, 0.7, { fontSize: 13.5, color: i === 0 ? PALE : MUTED, align: "center" });
  });
}

// ================= 13. 모니터링 =================
{
  const s = slide(S5, "③ 서비스가 시작됐는지 끝까지 확인합니다", "동 전담인력이 계획 안내 후 1개월 안에 최초 모니터링으로 서비스 시작을 확인하고, 이후 3개월마다 정기 점검합니다. 변화가 있으면 변경안을 만들어 통합지원회의에 다시 올립니다.");
  // 타임라인
  T(s, "모니터링 주기 (동 전담인력)", MX, 1.5, 6, 0.4, { fontSize: 16, bold: true, color: NAVY });
  s.addShape(pres.ShapeType.line, { x: MX + 0.3, y: 2.65, w: CW - 0.6, h: 0, line: { color: H.accent3, width: 3 }, objectName: nm("line") });
  const tl = [["계획 안내", "개인별지원계획 확정 통보", 0.3, NAVY], ["1개월 이내", "최초 모니터링\n서비스 개시 확인(대면·유선)", 3.6, ORANGE], ["3개월마다", "정기 모니터링\n제공기록 확인 후 유선 점검", 7.0, BLUE], ["종결", "통합지원회의 거쳐 종결", 9.95, NAVY]];
  tl.forEach((t) => {
    const x = MX + t[2];
    dot(s, "", x + 0.85, 2.43, 0.44, t[3]);
    T(s, t[0], x, 1.95, 2.15, 0.42, { fontSize: 15, bold: true, color: t[3] === ORANGE ? ORANGE : NAVY, align: "center" });
    T(s, t[1], x - 0.2, 2.98, 2.55, 0.75, { fontSize: 12.5, color: INK, align: "center" });
  });
  // 변경 흐름
  T(s, "변화가 있으면", MX, 4.05, 6, 0.4, { fontSize: 16, bold: true, color: NAVY });
  const fl = ["전담인력\n모니터링", "변경(안)\n작성", "통합지원회의\n재상정", "시가 제공기관에\n변경 의뢰"];
  const fw = 2.55, fg = 0.55;
  fl.forEach((f, i) => {
    const x = MX + i * (fw + fg);
    box(s, x, 4.55, fw, 1.2, i === 2 ? NAVY : TINT);
    T(s, f, x, 4.55, fw, 1.2, { fontSize: 15, bold: true, color: i === 2 ? WHITE : NAVY, align: "center", valign: "middle" });
    if (i < fl.length - 1) arrowR(s, x + fw + 0.07, 4.98, 0.4, 0.36);
  });
  T(s, "목표: 모니터링 수행률 80% · 모니터링 후 계획 변경·종결 90%", MX, 6.0, CW, 0.4, { fontSize: 13.5, bold: true, color: ORANGE });
}

// ================= 14. 서비스 =================
{
  const s = slide(S5, "④ 재가의료 이용을 늘리고 돌봄 공백을 메웁니다", "방문의료 이용자를 매년 30%씩 늘려 2030년 199명을 목표로 합니다. 의왕살핌 식사지원은 퇴원·급성기 치료식 중심 맞춤형 영양식으로 개편하고, 이동지원·회복살핌 꾸러미·AI 돌봄로봇 3종을 새로 도입합니다.");
  T(s, "방문의료(건강살핌) 이용자(명)", MX, 1.5, 5.5, 0.4, { fontSize: 16, bold: true, color: NAVY });
  s.addChart(pres.charts.BAR, [{ name: "이용자", labels: ["’26", "’27", "’30"], values: [70, 91, 199] }],
    { x: MX - 0.1, y: 1.9, w: 5.4, h: 3.4, barDir: "col", barGapWidthPct: 55, chartColors: [H.accent3, H.accent1, H.accent2], invertedColors: [H.accent3],
      showValue: true, dataLabelPosition: "outEnd", showLegend: false, valAxisHidden: true, valAxisMaxVal: 230, valGridLine: { style: "none" }, catGridLine: { style: "none" }, ...chartText, objectName: nm("chart") });
  badge(s, "매년 +30%", MX + 3.6, 1.95, 1.3);
  box(s, MX, 5.45, 5.2, 0.95, TINT);
  T(s, "의왕살핌 식사지원 → 치료식 맞춤형 영양식으로 개편", MX + 0.2, 5.45, 4.85, 0.95, { fontSize: 14, bold: true, color: NAVY, valign: "middle" });
  // 신규 3종
  const x = 6.25, w = W - MX - x;
  T(s, "돌봄 공백 신규서비스", x, 1.5, w, 0.4, { fontSize: 16, bold: true, color: NAVY });
  const nw = [["이동지원", "승강기 없는 주택 거동불편자\n병원 진료 이동", "30명"], ["회복살핌 꾸러미", "치매·퇴원환자\n건강관리용품 지원", "100명"], ["AI 돌봄로봇", "독거·안전확인 필요 대상\n안부 확인", "100명"]];
  nw.forEach((v, i) => {
    const y = 1.95 + i * 1.5;
    box(s, x, y, w, 1.3, WHITE, { line: MID, lw: 1.5 });
    badge(s, "신규", x + 0.25, y + 0.2, 0.65);
    T(s, v[0], x + 1.05, y + 0.12, w - 2.8, 0.45, { fontSize: 18, bold: true, color: NAVY, valign: "middle" });
    T(s, v[1], x + 0.25, y + 0.62, w - 2.4, 0.6, { fontSize: 12.5, color: MUTED });
    T(s, v[2], x + w - 1.8, y, 1.55, 1.3, { fontSize: 30, bold: true, color: ORANGE, align: "right", valign: "middle" });
  });
  T(s, "+ 장애인 건강주치의 참여 의원 1개소 이상 발굴", x, 6.45, w, 0.35, { fontSize: 13, bold: true, color: MUTED });
}

// ================= 15. 추진체계 =================
{
  const s = slide(S5, "⑤ 함께 운영합니다", "복지정책과 통합돌봄팀이 총괄하고, 6개 동 전담인력이 발굴부터 모니터링까지 맡습니다. 공단·보건소·협업부서·제공기관이 통합지원회의에서 함께 계획을 심의합니다.");
  // 상단 회의체
  box(s, 3.3, 1.45, 3.2, 0.8, TINT, { line: MID });
  T(s, "통합지원협의체\n연 2회 · 지역계획 심의", 3.3, 1.45, 3.2, 0.8, { fontSize: 13, bold: true, color: NAVY, align: "center", valign: "middle" });
  box(s, 6.83, 1.45, 3.2, 0.8, TINT, { line: MID });
  T(s, "통합지원회의\n월 2회 이상 · 계획 심의", 6.83, 1.45, 3.2, 0.8, { fontSize: 13, bold: true, color: NAVY, align: "center", valign: "middle" });
  // 중앙
  box(s, 3.3, 2.55, 6.73, 1.35, NAVY);
  T(s, "복지정책과 통합돌봄팀", 3.3, 2.62, 6.73, 0.65, { fontSize: 24, bold: true, color: WHITE, align: "center", valign: "middle" });
  T(s, "사업 총괄 · 지원계획 확정 · 서비스 연계·조정", 3.3, 3.25, 6.73, 0.5, { fontSize: 13, color: PALE, align: "center", valign: "middle" });
  arrowD(s, 6.48, 3.97, 0.36, 0.3);
  box(s, 3.3, 4.35, 6.73, 1.35, ORANGE);
  T(s, "6개 동 주민센터 · 전담인력 6명", 3.3, 4.42, 6.73, 0.6, { fontSize: 20, bold: true, color: WHITE, align: "center", valign: "middle" });
  T(s, "발굴 · 조사 · 개인별지원계획 · 연계 · 모니터링", 3.3, 5.0, 6.73, 0.5, { fontSize: 13, color: WHITE, align: "center", valign: "middle" });
  // 좌우
  const side = (x, title, items) => {
    box(s, x, 2.55, 2.4, 3.15, TINT);
    T(s, title, x, 2.65, 2.4, 0.45, { fontSize: 15, bold: true, color: NAVY, align: "center" });
    T(s, items, x + 0.15, 3.15, 2.1, 2.45, { fontSize: 12.5, color: INK, align: "center", valign: "top", paraSpaceAfter: 4 });
  };
  side(MX, "전문기관", "건보공단 의왕지사\n노인 통합판정·빅데이터\n\n연금공단 군포의왕지사\n장애인 종합판정");
  side(W - MX - 2.4, "보건소·협업부서", "지역보건팀·치매안심센터\n정신건강복지센터\n\n노인장애인과\n복지지원·생활보장팀\n건축과");
  box(s, 3.3, 5.95, 6.73, 0.7, TINT, { line: MID });
  T(s, "서비스 제공기관 18개소 · 방문진료 의원 12개소 · 퇴원 협약병원", 3.3, 5.95, 6.73, 0.7, { fontSize: 13, bold: true, color: NAVY, align: "center", valign: "middle" });
}

// ================= 16. 교육·홍보 =================
{
  const s = slide(S5, "⑥ 사람을 키우고 주민에게 알립니다", "자체교육은 ‘의왕돌봄마당’으로 묶어 연 4회, 경기복지재단 시군 맞춤형 교육 1회를 더해 연 5회 운영합니다. 동별 찾아가는 설명회를 연 1회 이상, 총 6회 엽니다.");
  const k = [["10명", "전담인력 전원\n중앙·전문기관 교육 이수"], ["연 5회", "「의왕돌봄마당」 4회\n+ 시군 맞춤형 교육 1회"], ["연 6회", "찾아가는 통합돌봄 설명회\n동별 1회 이상"], ["1,000건", "홍보자료 배포\n경로당·약국·의원 등"]];
  const cw = (CW - 3 * 0.3) / 4;
  k.forEach((v, i) => {
    const x = MX + i * (cw + 0.3);
    box(s, x, 1.6, cw, 2.7, i < 2 ? NAVY : TINT);
    T(s, v[0], x, 1.85, cw, 1.0, { fontSize: 40, bold: true, color: i < 2 ? WHITE : NAVY, align: "center", valign: "middle" });
    T(s, v[1], x + 0.15, 3.0, cw - 0.3, 1.0, { fontSize: 14, color: i < 2 ? PALE : INK, align: "center" });
  });
  T(s, "인력 양성", MX, 4.5, cw * 2 + 0.3, 0.35, { fontSize: 13, bold: true, color: MUTED, align: "center" });
  T(s, "주민 홍보", MX + 2 * (cw + 0.3), 4.5, cw * 2 + 0.3, 0.35, { fontSize: 13, bold: true, color: MUTED, align: "center" });
  box(s, MX, 5.05, CW, 1.2, TINT);
  dot(s, "↺", MX + 0.3, 5.3, 0.7, ORANGE, 20);
  T(s, "이용자 만족도 조사 연 300명 → 건의사항 검토 → 사업 운영에 반영", MX + 1.2, 5.05, CW - 1.4, 1.2, { fontSize: 17, bold: true, color: NAVY, valign: "middle" });
}

// ================= 17. 예산 =================
pres.addSection({ title: S6 });
{
  const s = slide(S6, "2027년 예산: 706,636천원", "총 7억 664만원으로 사업비 6억 8,964만원, 운영비 1,700만원입니다. 재원은 국비 50%, 도비 5%, 시비 45%입니다. 사업비의 75%가 의왕살핌입니다.");
  T(s, "재원", MX, 1.5, 4, 0.4, { fontSize: 16, bold: true, color: NAVY });
  s.addChart(pres.charts.DOUGHNUT, [{ name: "재원", labels: ["국비 50%", "시비 45%", "도비 5%"], values: [353318, 317986, 35332] }],
    { x: MX - 0.1, y: 1.85, w: 4.4, h: 3.9, holeSize: 60, chartColors: [H.dk2, H.accent3, H.accent2], showPercent: false, showValue: false, showLegend: true, legendPos: "b",
      legendFontSize: 13, legendColor: H.dk1, legendFontFace: "+mn-lt", dataBorder: { pt: 1, color: H.lt1 }, objectName: nm("chart") });
  T(s, "사업별(천원)", 5.3, 1.5, 6, 0.4, { fontSize: 16, bold: true, color: NAVY });
  s.addChart(pres.charts.BAR, [{ name: "예산", labels: ["운영비", "퇴원환자 연계수당", "재택의료센터 지원", "신규 3종", "신규특화(보조)", "의왕살핌"], values: [17000, 500, 5000, 110000, 60000, 514136] }],
    { x: 5.2, y: 1.85, w: 7.55, h: 3.9, barDir: "bar", barGapWidthPct: 40, chartColors: [H.accent1], showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: "#,##0",
      showLegend: false, valAxisHidden: true, valAxisMaxVal: 640000, valGridLine: { style: "none" }, catGridLine: { style: "none" }, ...chartText, objectName: nm("chart") });
  box(s, MX, 5.95, CW, 0.75, TINT);
  T(s, "사업비 689,636  +  운영비 17,000  =  706,636천원   |   신규 3종: 이동지원 · 회복살핌 꾸러미 · AI 돌봄로봇", MX + 0.25, 5.95, CW - 0.5, 0.75, { fontSize: 14, bold: true, color: NAVY, valign: "middle", align: "center" });
}

// ================= 18. 심의 요청 =================
{
  const s = pres.addSlide({ masterName: "CLOSING", sectionTitle: S6 });
  s.addText("심의 요청 사항", { placeholder: "title" });
  s.addNotes("이상 네 가지에 대해 위원님들의 심의와 의견을 부탁드립니다. 주신 의견은 계획에 반영해 확정하겠습니다.");
  const it = [["목표 대상자", "통합지원 대상자 541명"], ["성과지표", "발굴률 20% 등 중앙지표 · 자체지표"], ["추진과제", "4개 과업 · 12개 핵심사업"], ["예산", "706,636천원"]];
  const cw = (CW - 3 * 0.3) / 4;
  it.forEach((v, i) => {
    const x = MX + i * (cw + 0.3);
    box(s, x, 1.75, cw, 2.6, BLUE);
    dot(s, String(i + 1), x + cw / 2 - 0.3, 1.95, 0.6, ORANGE, 18);
    T(s, v[0], x, 2.7, cw, 0.55, { fontSize: 21, bold: true, color: WHITE, align: "center", valign: "middle" });
    T(s, v[1], x + 0.15, 3.3, cw - 0.3, 0.9, { fontSize: 14, color: PALE, align: "center" });
  });
  T(s, "향후 절차", MX, 4.85, CW, 0.4, { fontSize: 15, bold: true, color: ORANGE });
  const fl = ["협의체 심의\n11. 26.", "의견 반영", "지역계획 확정·제출", "2027. 1. 시행"];
  const fw = 2.55, fg = 0.55;
  fl.forEach((f, i) => {
    const x = MX + i * (fw + fg);
    box(s, x, 5.35, fw, 0.95, i === 0 ? ORANGE : WHITE);
    T(s, f, x, 5.35, fw, 0.95, { fontSize: 15, bold: true, color: i === 0 ? WHITE : NAVY, align: "center", valign: "middle" });
    if (i < fl.length - 1) arrowR(s, x + fw + 0.07, 5.65, 0.4, 0.36, PALE);
  });
}

(async () => {
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT);
})();
