// 협의체 심의 발표자료 v3(레퍼런스: 과기정통부 업무계획형, 15장)
// 실행: NODE_PATH=<node_modules> APPLY_THEME=<apply_theme.js> node build_v3.js <out.pptx>
const path = require("path");
const pptxgen = require("pptxgenjs");
const React = require("react");
const RDS = require("react-dom/server");
const sharp = require("sharp");
const L = require("react-icons/lu");
const { applyTheme } = require(process.env.APPLY_THEME);

const OUT = process.argv[2] || "deck_v3.pptx";
const V3 = path.join(__dirname, "..", "figures", "v3");
const FONT = process.env.DECK_FONT || "Pretendard"; // 본문 글꼴(제목은 TITLE_FONT)
const TITLE_FONT = process.env.DECK_TITLE_FONT || "Pretendard ExtraBold"; // 제목 글꼴
const THEME = { name: "Uiwang Care 2027 v3", headFontFace: FONT, bodyFontFace: FONT,
  colors: { dk1: "111B3D", lt1: "FFFFFF", dk2: "13277A", lt2: "EEF3FF", accent1: "2F5BEA", accent2: "F08A24", accent3: "8FAEFF",
    accent4: "DCE6FF", accent5: "5B6785", accent6: "0E1F5B", hlink: "2F5BEA", folHlink: "5B6785" } };
const H = THEME.colors;
const NAVY = H.dk2, INK = H.dk1, BLUE = H.accent1, SKY = H.accent3, PALE = H.accent4, GRAY = H.accent5, BAND = H.accent6, ORG = H.accent2, WHITE = "FFFFFF", TINT = H.lt2;
const TAG = { 신규: "F08A24", 확대: "2F5BEA", 개선: "0E9F6E" };

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
pres.theme = { headFontFace: FONT, bodyFontFace: FONT };
pres.title = "2027년도 의왕시 통합지원 지역계획(안)";
pres.author = "의왕시 복지정책과 통합돌봄팀";
const W = 13.333, MX = 0.6, CW = W - 2 * MX;

let n = 0; const nm = (p) => `${p}-${++n}`;
const icons = {};
async function icon(name, color, size = 256) {
  const k = name + color; if (icons[k]) return icons[k];
  const svg = RDS.renderToStaticMarkup(React.createElement(L[name], { color: "#" + color, size: String(size) }));
  const png = await sharp(Buffer.from(svg)).png().toBuffer();
  return (icons[k] = "image/png;base64," + png.toString("base64"));
}
const T = (s, text, x, y, w, h, o = {}) => s.addText(text, { x, y, w, h, isTextBox: true, margin: 0, fontFace: FONT, fontSize: 14, color: INK, valign: "top", objectName: nm("t"), ...o });
const R = (s, x, y, w, h, fill, o = {}) => s.addShape(o.round ? pres.ShapeType.roundRect : pres.ShapeType.rect, { x, y, w, h, fill: { color: fill, transparency: o.tr || 0 }, line: o.line ? { color: o.line, width: o.lw || 1 } : { type: "none" }, rectRadius: o.round ? (o.r ?? 0.1) : undefined, objectName: nm("r") });
const C = (s, x, y, d, fill, o = {}) => s.addShape(pres.ShapeType.ellipse, { x, y, w: d, h: d, fill: { color: fill, transparency: o.tr || 0 }, line: o.line ? { color: o.line, width: o.lw || 1 } : { type: "none" }, objectName: nm("c") });
async function ic(s, name, x, y, d, color = BLUE, bg = PALE) {
  C(s, x, y, d, bg);
  s.addImage({ data: await icon(name, color), x: x + d * 0.24, y: y + d * 0.24, w: d * 0.52, h: d * 0.52, altText: "", objectName: nm("i") });
}
async function lett(s, name, x, y, w, alt, center = false) { // 레터링 이미지(비율 자동)
  const m = await sharp(path.join(V3, name + ".png")).metadata();
  const h = w * m.height / m.width; if (center) x = (W - w) / 2;
  s.addImage({ path: path.join(V3, name + ".png"), x, y, w, h, altText: alt, objectName: nm("title") }); return h;
}
const bg = (s, f) => s.addImage({ path: path.join(V3, f), x: 0, y: 0, w: W, h: 7.5, altText: "", objectName: nm("bg") });
function head(s, no, kicker, title) {
  bg(s, "bg_content.png");
  T(s, [{ text: no + "  ", options: { bold: true, color: BLUE } }, { text: "—  " + kicker, options: { color: GRAY } }], MX, 0.42, 10, 0.32, { fontSize: 13 });
  T(s, title, MX, 0.78, CW, 1.05, { fontFace: TITLE_FONT, fontSize: 27, bold: true, color: INK, valign: "top", lineSpacingMultiple: 1.05 });
}
function tab(s, text, x, y, w) { // 남색 머리띠 + 아래 가는 선
  R(s, x, y, Math.min(w, Math.max(2.2, text.length * 0.2 + 0.6)), 0.36, NAVY);
  T(s, text, x + 0.15, y, Math.min(w, Math.max(2.2, text.length * 0.2 + 0.6)) - 0.2, 0.36, { fontSize: 13.5, bold: true, color: WHITE, valign: "middle" });
  s.addShape(pres.ShapeType.line, { x, y: y + 0.36, w, h: 0, line: { color: NAVY, width: 1.5 }, objectName: nm("l") });
}
function band(s, label, items) {
  R(s, 0, 6.82, W, 0.68, BAND);
  T(s, label, MX, 6.82, 1.7, 0.68, { fontSize: 14, bold: true, color: WHITE, valign: "middle" });
  const w = (CW - 1.8) / items.length;
  items.forEach((t, i) => T(s, "▪ " + t, MX + 1.8 + i * w, 6.82, w, 0.68, { fontSize: 13, color: "D6E0FF", valign: "middle" }));
}
function badge(s, tag, x, y) {
  R(s, x, y, 0.58, 0.28, TAG[tag], { round: true, r: 0.14 });
  T(s, tag, x, y, 0.58, 0.28, { fontSize: 11, bold: true, color: WHITE, align: "center", valign: "middle" });
}
const arrowR = (s, x, y, w = 0.34, h = 0.3, fill = SKY) => s.addShape(pres.ShapeType.chevron, { x, y, w, h, fill: { color: fill }, line: { type: "none" }, objectName: nm("a") });
const chartTxt = { catAxisLabelFontFace: "+mn-lt", valAxisLabelFontFace: "+mn-lt", dataLabelFontFace: "+mn-lt", legendFontFace: "+mn-lt",
  catAxisLabelColor: H.dk1, dataLabelColor: H.dk1, catAxisLabelFontSize: 12, dataLabelFontSize: 12, valGridLine: { style: "none" }, catGridLine: { style: "none" }, valAxisHidden: true, showLegend: false };

(async () => {
  // 1 ───────── 표지
  {
    const s = pres.addSlide();
    bg(s, "bg_cover.png");
    s.addImage({ data: await icon("LuHeartHandshake", NAVY), x: MX, y: 0.45, w: 0.32, h: 0.32, objectName: nm("i") });
    T(s, "의왕시 복지정책과 통합돌봄팀", MX + 0.42, 0.45, 5, 0.32, { fontSize: 12, bold: true, color: NAVY, valign: "middle" });
    T(s, "의왕시 2027년도 통합지원 지역계획(안)", MX, 2.0, 8, 0.45, { fontSize: 20, color: INK, bold: true });
    await lett(s, "L_cover", MX - 0.05, 2.5, 7.9, "살던 곳에서 건강하게, 함께 살피는 의왕");
    T(s, "제2차 통합지원협의체 심의  |  2026. 11. 26.", MX, 5.45, 7, 0.4, { fontSize: 14, color: GRAY });
    s.addNotes("2027년도 의왕시 통합지원 지역계획(안)을 보고드리고 심의를 요청드립니다.");
  }
  // 2 ───────── 2026 성과 (레퍼런스 2쪽형: 3단 패널)
  {
    const s = pres.addSlide(); bg(s, "bg_content.png");
    await lett(s, "L_s2", 0, 0.32, 5.6, "법 시행 첫해, 먼저 찾고 연결한 2026년", true);

    const cols = [
      ["대상자 발굴·연계", [["LuUsers", "신청 390 → 대상자 276명", "우선발굴필요군 비중 80.4%"], ["LuTrendingUp", "10월 신청 147명", "1~9월 월평균의 5배"], ["LuNetwork", "서비스 연계 740건", "대상자 1인당 2.7건"]]],
      ["재가돌봄 인프라", [["LuStethoscope", "방문진료 참여의원 12개소", "무공급 → 의과 3·한의 9"], ["LuHospital", "재택의료센터 2개소", "관내 1 · 관외(안양) 1"], ["LuHouse", "의왕살핌 340명 이용", "일상 138 · 건강 103 · 내집 99"]]],
      ["운영 기반 구축", [["LuBuilding2", "통합돌봄팀 신설", "’26. 2. · 조례 시행 ’26. 3."], ["LuClipboardList", "통합지원회의 15회", "신규 231 · 변경 29 · 종결 40건"], ["LuUserCheck", "동 전담인력 6명 배치", "’26. 11. 30. 겸직 → 전담 전환"]]],
    ];
    const cw = (CW - 0.4) / 3;
    for (let c = 0; c < 3; c++) {
      const x = MX + c * (cw + 0.2);
      R(s, x, 1.7, cw, 0.45, NAVY);
      T(s, cols[c][0], x, 1.7, cw, 0.45, { fontSize: 15, bold: true, color: WHITE, align: "center", valign: "middle" });
      R(s, x, 2.15, cw, 3.75, WHITE, { tr: 15 });
      for (let i = 0; i < 3; i++) {
        const y = 2.35 + i * 1.15;
        await ic(s, cols[c][1][i][0], x + 0.2, y, 0.72);
        T(s, cols[c][1][i][1], x + 1.05, y + 0.02, cw - 1.2, 0.38, { fontSize: 15, bold: true, color: INK, valign: "middle" });
        T(s, cols[c][1][i][2], x + 1.05, y + 0.4, cw - 1.2, 0.34, { fontSize: 12.5, color: GRAY, valign: "middle" });
      }
      if (c < 2) s.addShape(pres.ShapeType.line, { x: x + cw + 0.1, y: 2.25, w: 0, h: 3.55, line: { color: PALE, width: 1 }, objectName: nm("l") });
    }
    s.addShape(pres.ShapeType.line, { x: MX, y: 6.42, w: 1.6, h: 0, line: { color: NAVY, width: 1 }, objectName: nm("l") });
    s.addShape(pres.ShapeType.line, { x: W - MX - 1.6, y: 6.42, w: 1.6, h: 0, line: { color: NAVY, width: 1 }, objectName: nm("l") });
    T(s, [{ text: "기반을 토대로 2027년에는 ", options: { color: INK } }, { text: "끝까지", options: { color: "C8157F" } }, { text: " 살피는 체계로", options: { color: INK } }], MX + 1.7, 6.15, CW - 3.4, 0.55, { fontSize: 19, bold: true, align: "center", valign: "middle" });
    s.addNotes("법 시행 첫해, 390명이 신청해 276명을 지원했고 740건을 연계했습니다. 없던 방문진료·재택의료를 확보하고, 통합돌봄팀과 조례, 동 전담인력까지 운영 기반을 갖췄습니다.");
  }
  // 2-1 ───────── 의왕살핌 3대 패키지(사진 자리)
  async function photo(s, x, y, w, h, label) {
    s.addShape(pres.ShapeType.roundRect, { x, y, w, h, fill: { color: "F3F5FC" }, line: { color: "AEB9D8", width: 1.25, dashType: "dash" }, rectRadius: 0.1, objectName: nm("photo") });
    s.addImage({ data: await icon("LuCamera", "9AA6C8"), x: x + w / 2 - 0.22, y: y + h / 2 - 0.42, w: 0.44, h: 0.44, objectName: nm("i") });
    T(s, label, x + 0.1, y + h / 2 + 0.08, w - 0.2, 0.5, { fontSize: 12, color: "7A86A8", align: "center" });
  }
  {
    const s = pres.addSlide();
    head(s, "성과", "2026년 지역특화서비스", "국가서비스가 닿지 않는 공백을\n의왕살핌으로 메웠습니다.");
    const cards = [["LuHouse", "일상살핌", "생활·동행돌봄, 식사지원", "138명", "1인 200만원 한도", "사진: 병원 동행 · 식사 배달"],
      ["LuStethoscope", "건강살핌", "방문진료 · 방문재활 · 가정간호", "103명", "1인 60만원 한도", "사진: 방문재활 · 방문진료"],
      ["LuShieldCheck", "내집살핌", "안전손잡이 · 미끄럼방지 · 문턱제거", "99명", "생애 200만원 한도", "사진: 주거개선 시공 전 → 후"]];
    const cw = (CW - 0.6) / 3;
    for (let i = 0; i < 3; i++) {
      const x = MX + i * (cw + 0.3), y = 1.95, c = cards[i];
      R(s, x, y, cw, 4.7, WHITE, { round: true, r: 0.1, line: PALE });
      await photo(s, x + 0.18, y + 0.18, cw - 0.36, 2.15, c[5]);
      await ic(s, c[0], x + 0.22, y + 2.55, 0.62);
      T(s, c[1], x + 0.98, y + 2.52, cw - 1.1, 0.4, { fontSize: 19, bold: true, color: INK, fontFace: TITLE_FONT });
      T(s, c[2], x + 0.98, y + 2.92, cw - 1.1, 0.35, { fontSize: 12.5, color: GRAY });
      s.addShape(pres.ShapeType.line, { x: x + 0.22, y: y + 3.5, w: cw - 0.44, h: 0, line: { color: PALE, width: 1 }, objectName: nm("l") });
      T(s, c[3], x + 0.22, y + 3.62, cw / 2, 0.85, { fontSize: 34, bold: true, color: i === 1 ? "C8157F" : NAVY, fontFace: TITLE_FONT, valign: "middle" });
      T(s, ["’26 이용", c[4]].join("\n"), x + cw / 2, y + 3.68, cw / 2 - 0.22, 0.75, { fontSize: 12, color: GRAY, align: "right", valign: "middle" });
    }
    band(s, "운영 원칙", ["국가돌봄서비스 우선, 공백만 의왕살핌", "소득 무관 신청 · 중위소득별 본인부담 차등", "방문진료 의원 12개소 실제 진료 시작"]);
    s.addNotes("의왕살핌은 장기요양이나 노인맞춤돌봄 같은 국가서비스를 먼저 연계하고, 그래도 남는 공백을 메우는 우리 시 서비스입니다. 일상살핌은 생활·동행돌봄과 식사, 건강살핌은 집으로 찾아가는 진료·재활·간호, 내집살핌은 낙상을 막는 주거 개선입니다. 올해 340명이 이용했습니다.");
  }
  // 2-2 ───────── 공백보완 특화사업(사진 자리)
  {
    const s = pres.addSlide();
    head(s, "성과", "2026년 지역특화서비스", "운영하며 찾은 공백은\n3개월 만에 새 서비스로 메웠습니다.");
    const cards = [["LuUtensils", "맞춤형 영양식", "저작완화식 · 만성질환식 · 유동식\n주 2회 배달, 안부 확인", "55명", "사진: 영양식 도시락 구성"],
      ["LuHeartPulse", "건강돌봄가", "방문 스트레칭, 건강·영양 모니터링\n복약 지원", "50명", "사진: 건강돌봄가 방문 활동"],
      ["LuPuzzle", "동고동락", "치매 대상 방문형 프로그램\n인지·운동·만들기(주 2회, 14회기)", "22명", "사진: 인지활동 · 만들기"]];
    const cw = (CW - 0.6) / 3;
    for (let i = 0; i < 3; i++) {
      const x = MX + i * (cw + 0.3), y = 1.95, c = cards[i];
      R(s, x, y, cw, 4.7, WHITE, { round: true, r: 0.1, line: PALE });
      await photo(s, x + 0.18, y + 0.18, cw - 0.36, 2.15, c[4]);
      badge(s, "신규", x + cw - 0.82, y + 2.58);
      await ic(s, c[0], x + 0.22, y + 2.55, 0.62);
      T(s, c[1], x + 0.98, y + 2.52, cw - 1.9, 0.4, { fontSize: 19, bold: true, color: INK, fontFace: TITLE_FONT });
      T(s, c[2], x + 0.22, y + 3.25, cw - 0.44, 0.65, { fontSize: 12.5, color: GRAY });
      T(s, c[3], x + 0.22, y + 3.85, cw - 0.44, 0.7, { fontSize: 30, bold: true, color: NAVY, fontFace: TITLE_FONT, align: "right", valign: "middle" });
    }
    band(s, "추진 경과", ["상반기 운영 중 치매·영양·건강 공백 확인", "’26. 8. 공모로 3개 사업 개시", "이용 127명(’26. 10.)"]);
    s.addNotes("상반기에 운영해 보니 치매, 영양, 건강 모니터링 공백이 보였습니다. 그래서 8월에 맞춤형 영양식, 건강돌봄가, 방문형 치매 프로그램 동고동락을 시작했고, 10월까지 127명이 이용했습니다.");
  }
  // 3 ───────── 지역 진단
  {
    const s = pres.addSlide();
    head(s, "진단", "지역의 돌봄 수요", "노인 인구는 5년 새 43.5% 늘었고,\n초고령 1인 가구는 부곡동에 가장 많습니다.");
    tab(s, "초고령 · 1인 가구", MX, 1.95, 5.2);
    const st = [["33,930명", "65세 이상 · 고령화율 20.3%"], ["1,046세대", "85세 이상 1인세대"], ["48.6%", "핵심군 중 85세 이상 · 경기 3위"]];
    st.forEach((v, i) => {
      const y = 2.5 + i * 1.05;
      R(s, MX, y, 5.2, 0.92, WHITE, { round: true, r: 0.08, line: PALE });
      T(s, v[0], MX + 0.25, y, 2.1, 0.92, { fontSize: 26, bold: true, color: i === 1 ? BLUE : NAVY, valign: "middle" });
      T(s, v[1], MX + 2.4, y, 2.7, 0.92, { fontSize: 12.5, color: GRAY, valign: "middle" });
    });
    tab(s, "동별 85세 이상 1인세대(세대)", 6.15, 1.95, W - MX - 6.15);
    s.addChart(pres.charts.BAR, [{ name: "85세 이상 1인세대", labels: ["부곡동", "청계동", "내손1동", "오전동", "내손2동", "고천동"], values: [292, 211, 177, 156, 114, 96] }],
      { x: 6.1, y: 2.4, w: W - MX - 6.1, h: 3.35, barDir: "col", barGapWidthPct: 35, chartColors: [H.accent1, "6F8FE8", "6F8FE8", "6F8FE8", "6F8FE8", "6F8FE8"], invertedColors: [H.accent1],
        showValue: true, dataLabelPosition: "outEnd", valAxisMaxVal: 345, ...chartTxt,
        catAxisLabelFontSize: 16, catAxisLabelFontBold: true, catAxisLabelColor: H.dk1, catAxisLineShow: true, catAxisLineColor: "9FB0D8",
        dataLabelFontSize: 18, dataLabelFontBold: true, dataLabelColor: H.dk2, objectName: nm("ch") });
    band(s, "핵심 진단", ["방문의료 수요 대비 이용 10.5%", "퇴원환자 협약병원 의뢰 0건", "돌봄 인지 52.9% → 이용 6.4%"]);
    s.addNotes("노인은 3만 3,930명으로 5년 새 43.5% 늘었습니다. 85세 이상 1인세대 1,046세대 중 부곡동이 292세대로 가장 많습니다. 방문의료 이용, 퇴원환자 연계, 주민 인지도가 핵심 과제입니다.");
  }
  // 4 ───────── 비전과 3대 목표 (레퍼런스 3쪽형: 원 3개)
  {
    const s = pres.addSlide(); bg(s, "bg_mid.png");
    await lett(s, "L_s4", 0, 0.8, 8.4, "살던 곳에서 건강하게, 함께 살피는 의왕 — 의왕을 위한 2027년 3대 목표", true);

    // 3대 목표: 아이콘 + 핵심어(크게) + 목표 문장(작게) + 연결 과업
    const g = [["01", "LuSearch", "끝까지 살피는\n대상자 관리", "먼저 찾고 끝까지 살피는\n대상자 관리체계 확립", "13277A", "과업 01  대상자 발굴·관리"],
      ["02", "LuHouse", "재가돌봄\n공백 해소", "재가의료 이용 확대와\n의왕살핌 고도화", "2F5BEA", "과업 02  서비스 확충·품질향상"],
      ["03", "LuHandshake", "함께하는\n운영기반", "보건·복지·민간이 함께하는\n수요자 중심 운영기반 강화", "5A82F5", "과업 03·04  운영기반·교육홍보"]];
    const d = 3.85, ov = 0.25, x0 = (W - (3 * d - 2 * ov)) / 2, y0 = 2.55;
    for (let i = 0; i < 3; i++) {
      const v = g[i], x = x0 + i * (d - ov), cx = x + d / 2;
      s.addShape(pres.ShapeType.ellipse, { x, y: y0, w: d, h: d, fill: { color: v[4] }, line: { color: "FFFFFF", width: 3 }, shadow: { type: "outer", color: "1B2F7A", opacity: 0.25, blur: 10, offset: 4, angle: 90 }, objectName: nm("c") });
      R(s, cx - 0.42, y0 + 0.38, 0.84, 0.34, "FFFFFF", { round: true, r: 0.17 });
      T(s, v[0], cx - 0.42, y0 + 0.38, 0.84, 0.34, { fontSize: 14, bold: true, color: v[4], align: "center", valign: "middle" });
      s.addImage({ data: await icon(v[1], "FFFFFF"), x: cx - 0.3, y: y0 + 0.88, w: 0.6, h: 0.6, objectName: nm("i") });
      T(s, v[2], x + 0.3, y0 + 1.58, d - 0.6, 1.05, { fontSize: 25, bold: true, color: WHITE, align: "center", valign: "middle", lineSpacingMultiple: 0.95 });
      s.addShape(pres.ShapeType.line, { x: cx - 0.45, y: y0 + 2.72, w: 0.9, h: 0, line: { color: "FFFFFF", width: 1.25, transparency: 40 }, objectName: nm("l") });
      T(s, v[3], x + 0.45, y0 + 2.82, d - 0.9, 0.62, { fontSize: 12.5, color: "E3EAFF", align: "center" });
      R(s, cx - 1.6, y0 + d + 0.1, 3.2, 0.38, "FFFFFF", { round: true, r: 0.19, line: PALE });
      T(s, v[5], cx - 1.6, y0 + d + 0.1, 3.2, 0.38, { fontSize: 12.5, bold: true, color: i === 2 ? "3C63E0" : v[4], align: "center", valign: "middle" });
    }
    s.addNotes("비전 아래 3대 목표입니다. 먼저 찾고 끝까지 살피는 관리체계, 재가의료 이용 확대와 의왕살핌 고도화, 보건·복지·민간이 함께하는 운영기반입니다.");
  }
  // 5 ───────── 4대 과업·12개 핵심사업
  {
    const s = pres.addSlide();
    head(s, "체계", "추진과업 및 핵심사업", "4대 과업, 12개 핵심사업으로\n2027년 목표를 달성하겠습니다.");
    const cols = [["대상자 발굴·관리", "LuSearch", [["초고령 위험군 선제발굴", "확대"], ["퇴원환자 의뢰 경로 이원화", "개선"], ["개인별지원계획 수립·모니터링 내실화", "개선"]]],
      ["서비스 확충·품질향상", "LuHeartPulse", [["의왕살핌 서비스 고도화", "확대"], ["돌봄 공백 신규서비스 도입", "신규"], ["장애인 건강주치의 참여 의원 발굴", "신규"]]],
      ["운영기반·민관협력", "LuHandshake", [["서비스 의뢰·회신 표준체계 구축", "신규"], ["전문기관 협업 정례화", "확대"], ["통합지원회의 운영 내실화", "개선"]]],
      ["교육·홍보", "LuMegaphone", [["「의왕돌봄마당」 민관 역량강화", "확대"], ["찾아가는 통합돌봄 설명회", "신규"], ["주민 생활 밀착 홍보", "신규"]]]];
    const cw = (CW - 0.45) / 4;
    for (let c = 0; c < 4; c++) {
      const x = MX + c * (cw + 0.15);
      R(s, x, 2.0, cw, 0.62, NAVY);
      s.addImage({ data: await icon(cols[c][1], WHITE), x: x + 0.2, y: 2.13, w: 0.36, h: 0.36, objectName: nm("i") });
      T(s, cols[c][0], x + 0.65, 2.0, cw - 0.75, 0.62, { fontSize: 15, bold: true, color: WHITE, valign: "middle" });
      cols[c][2].forEach((it, i) => {
        const y = 2.8 + i * 1.3;
        R(s, x, y, cw, 1.15, WHITE, { round: true, r: 0.08, line: PALE });
        badge(s, it[1], x + 0.18, y + 0.15);
        T(s, it[0], x + 0.18, y + 0.5, cw - 0.36, 0.6, { fontSize: 14, bold: true, color: INK, valign: "top" });
      });
    }
    const leg = [["신규", "5개"], ["확대", "4개"], ["개선", "3개"]];
    leg.forEach((l, i) => { badge(s, l[0], W - MX - 3.6 + i * 1.2, 1.55); T(s, l[1], W - MX - 3.6 + i * 1.2 + 0.63, 1.55, 0.5, 0.28, { fontSize: 12, bold: true, color: TAG[l[0]], valign: "middle" }); });
    s.addNotes("4대 과업 아래 12개 핵심사업입니다. 신규 5개, 확대 4개, 개선 3개입니다.");
  }
  // 6 ───────── 01 선제발굴
  {
    const s = pres.addSlide();
    head(s, "01", "대상자 발굴 및 관리", "신청을 기다리지 않고,\n85세 이상 1인 가구를 먼저 찾아가겠습니다.");
    tab(s, "초고령 위험군 선제 조사", MX, 1.95, CW);
    const st = [["LuDatabase", "명단 추출", "기초연금 수급자 중\n85세 이상 1인가구"], ["LuListChecks", "대상 정리", "통합돌봄\n신청자·대상자 제외"], ["LuUsers", "1차 선별", "노인맞춤돌봄 이용자\n→ 생활지원사 확인"], ["LuMapPin", "가정 방문", "서비스 미이용자\n→ 동 전담인력 방문"], ["LuFileText", "직권신청", "필요 시\n바로 신청"]];
    const sw = 2.05, gap = (CW - 5 * sw) / 4;
    for (let i = 0; i < 5; i++) {
      const x = MX + i * (sw + gap);
      await ic(s, st[i][0], x + sw / 2 - 0.45, 2.55, 0.9, i === 3 ? WHITE : BLUE, i === 3 ? BLUE : PALE);
      T(s, st[i][1], x, 3.55, sw, 0.4, { fontSize: 16, bold: true, color: i === 3 ? BLUE : INK, align: "center" });
      T(s, st[i][2], x, 3.98, sw, 0.7, { fontSize: 12.5, color: GRAY, align: "center" });
      if (i < 4) arrowR(s, x + sw + gap / 2 - 0.15, 2.85, 0.3, 0.32);
    }
    tab(s, "공공데이터 · 관계기관 발굴", MX, 4.95, CW);
    const pd = [["LuDatabase", "건보 빅데이터", "입원위험 재가급여자·등급외자, 반기 1회 동별 배분"], ["LuShieldCheck", "행복e음 위기정보", "복지사각지대 위기정보 매월 확인"], ["LuMegaphone", "주민·기관 의뢰", "설명회·통장·보건소·복지관 의뢰 → 직권신청"]];
    const pw = (CW - 0.4) / 3;
    for (let i = 0; i < 3; i++) {
      const x = MX + i * (pw + 0.2);
      await ic(s, pd[i][0], x, 5.48, 0.6);
      T(s, pd[i][1], x + 0.72, 5.45, pw - 0.8, 0.32, { fontSize: 14, bold: true, color: INK });
      T(s, pd[i][2], x + 0.72, 5.8, pw - 0.8, 0.6, { fontSize: 12, color: GRAY });
    }
    band(s, "핵심 목표", ["직권·자체발굴 발굴률 20% 이상", "상·하반기 명단 제공 · 월별 실적 점검", "지원계획 수립률 6% 이상"]);
    s.addNotes("기초연금 수급자 중 85세 이상 1인가구를 상·하반기 뽑아 동에 제공합니다. 노인맞춤돌봄 이용자는 생활지원사가 먼저 확인하고, 서비스를 받지 않는 분은 동 전담인력이 직접 방문합니다.");
  }
  // 7 ───────── 01 퇴원환자
  {
    const s = pres.addSlide();
    head(s, "01", "대상자 발굴 및 관리", "퇴원 후 돌봄 공백 없이\n지역으로 바로 연결하겠습니다.");
    const lanes = [["상급종합병원", "퇴원일이 전날 확정", [["LuFileText", "병원 안내문 배부"], ["LuHouse", "퇴원 후 동 신청"], ["LuMapPin", "동 전담인력 확인 방문"], ["LuClipboardList", "통합지원회의 계획 수립"]]],
      ["종합병원(협약)", "퇴원 전 공식 의뢰", [["LuHospital", "선별·환자평가"], ["LuFileText", "시 전담조직에 직접 의뢰"], ["LuListChecks", "서류 검토"], ["LuWallet", "계획 수립 · 연계수당 지급"]]]];
    for (let l = 0; l < 2; l++) {
      const y = 2.05 + l * 1.85;
      R(s, MX, y, 2.4, 1.55, l ? NAVY : BLUE);
      T(s, lanes[l][0], MX, y + 0.3, 2.4, 0.45, { fontSize: 17, bold: true, color: WHITE, align: "center" });
      T(s, lanes[l][1], MX, y + 0.8, 2.4, 0.4, { fontSize: 12.5, color: "D6E0FF", align: "center" });
      const sw = 2.05, x0 = MX + 2.6;
      for (let i = 0; i < 4; i++) {
        const x = x0 + i * (sw + 0.32);
        R(s, x, y, sw, 1.55, WHITE, { round: true, r: 0.08, line: PALE });
        await ic(s, lanes[l][2][i][0], x + sw / 2 - 0.33, y + 0.18, 0.66);
        T(s, lanes[l][2][i][1], x + 0.1, y + 0.95, sw - 0.2, 0.5, { fontSize: 13.5, bold: true, color: INK, align: "center" });
        if (i < 3) arrowR(s, x + sw + 0.03, y + 0.62, 0.26, 0.3);
      }
    }
    const k = [["60명", "퇴원환자 대상자 목표"], ["10명 → 50명", "협약병원 공식 의뢰('27 → '30)"], ["건당 5만원", "연계수당(협약병원)"]];
    const kw = (CW - 0.4) / 3;
    k.forEach((v, i) => { const x = MX + i * (kw + 0.2); T(s, v[0], x, 5.72, kw, 0.5, { fontSize: 24, bold: true, color: BLUE, align: "center", valign: "middle" }); T(s, v[1], x, 6.22, kw, 0.35, { fontSize: 12.5, color: GRAY, align: "center" }); });
    band(s, "퇴원 직후", ["회복살핌 꾸러미·노인맞춤 단기집중", "미해당 시 일상살핌으로 보완", "퇴원 후 1개월 이내 최초 모니터링"]);
    s.addNotes("상급종합병원은 퇴원일이 전날 확정되어 안내문 방식으로, 종합병원은 협약을 넓혀 퇴원 전에 시로 직접 의뢰받습니다.");
  }
  // 8 ───────── 01 모니터링
  {
    const s = pres.addSlide();
    head(s, "01", "대상자 발굴 및 관리", "서비스가 시작됐는지,\n끝까지 확인하겠습니다.");
    tab(s, "모니터링 주기 · 동 전담인력", MX, 1.95, CW);
    s.addShape(pres.ShapeType.line, { x: MX + 0.8, y: 3.25, w: CW - 1.6, h: 0, line: { color: SKY, width: 4 }, objectName: nm("l") });
    const tl = [["LuClipboardList", "계획 안내", "개인별지원계획\n확정 통보"], ["LuCalendarCheck", "1개월 이내", "최초 모니터링\n서비스 개시 확인(대면·유선)"], ["LuRefreshCw", "3개월마다", "정기 모니터링\n제공기록 확인 후 유선 점검"], ["LuCircleCheck", "종결", "통합지원회의를\n거쳐 종결"]];
    const tw = CW / 4;
    for (let i = 0; i < 4; i++) {
      const cx = MX + tw * i + tw / 2;
      await ic(s, tl[i][0], cx - 0.42, 2.83, 0.84, i === 1 ? WHITE : BLUE, i === 1 ? BLUE : "FFFFFF");
      T(s, tl[i][1], cx - 1.4, 3.75, 2.8, 0.4, { fontSize: 17, bold: true, color: i === 1 ? BLUE : INK, align: "center" });
      T(s, tl[i][2], cx - 1.4, 4.15, 2.8, 0.7, { fontSize: 12.5, color: GRAY, align: "center" });
    }
    tab(s, "변화가 있으면", MX, 5.05, CW);
    const fl = ["전담인력 모니터링", "변경(안) 작성", "통합지원회의 재상정", "시가 제공기관에 변경 의뢰"];
    const fw = (CW - 3 * 0.45) / 4;
    fl.forEach((f, i) => { const x = MX + i * (fw + 0.45); R(s, x, 5.6, fw, 0.75, i === 2 ? NAVY : WHITE, { round: true, r: 0.1, line: i === 2 ? null : PALE }); T(s, f, x, 5.6, fw, 0.75, { fontSize: 14, bold: true, color: i === 2 ? WHITE : INK, align: "center", valign: "middle" }); if (i < 3) arrowR(s, x + fw + 0.08, 5.83, 0.28, 0.3); });
    band(s, "핵심 목표", ["모니터링 수행률 80%", "모니터링 후 계획 변경·종결 90%", "제공기록 월 1회 행복e음 입력"]);
    s.addNotes("동 전담인력이 계획 안내 후 1개월 안에 서비스 시작을 확인하고, 3개월마다 정기 점검합니다. 변화가 있으면 변경안을 통합지원회의에 다시 올립니다.");
  }
  // 9 ───────── 02 재가의료·의왕살핌
  {
    const s = pres.addSlide();
    head(s, "02", "서비스 확충 및 품질향상", "집에서 진료받는 재가의료를 넓히고,\n의왕살핌을 고도화하겠습니다.");
    tab(s, "방문의료(건강살핌) 이용 확대", MX, 1.95, 5.6);
    s.addChart(pres.charts.BAR, [{ name: "이용자", labels: ["2026년", "2027년", "2030년"], values: [70, 91, 199] }],
      { x: MX, y: 2.4, w: 5.6, h: 3.35, barDir: "col", barGapWidthPct: 45, chartColors: ["6F8FE8", H.accent1, H.dk2], invertedColors: ["6F8FE8"], showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: '0"명"', valAxisMaxVal: 235, ...chartTxt,
        catAxisLabelFontSize: 16, catAxisLabelFontBold: true, catAxisLabelColor: H.dk1, catAxisLineShow: true, catAxisLineColor: "9FB0D8",
        dataLabelFontSize: 18, dataLabelFontBold: true, dataLabelColor: H.dk2, objectName: nm("ch") });
    R(s, MX + 0.15, 2.5, 1.45, 0.36, ORG, { round: true, r: 0.18 });
    T(s, "매년 +30%", MX + 0.15, 2.5, 1.45, 0.36, { fontSize: 13, bold: true, color: WHITE, align: "center", valign: "middle" });
    tab(s, "재가의료 기반 · 의왕살핌", 6.55, 1.95, W - MX - 6.55);
    const it = [["LuStethoscope", "방문진료 참여의원 12개소", "등록만 돼 있던 의원이 실제 진료 · 대상자 집중 발굴로 이용 활성화"], ["LuHospital", "재택의료센터 관내 1 · 관외 1", "추가 확충보다 연계 유지 · 이용 확대"], ["LuUtensils", "식사지원 개편", "일반식은 기존 서비스로 연계, 퇴원·급성기 치료식 맞춤형 영양식으로"], ["LuHouse", "의왕살핌 3대 패키지", "일상·건강·내집살핌 이용자 전년 대비 30% 확대"]];
    for (let i = 0; i < 4; i++) {
      const y = 2.5 + i * 0.98;
      await ic(s, it[i][0], 6.55, y, 0.66);
      T(s, it[i][1], 7.35, y - 0.02, W - MX - 7.35, 0.36, { fontSize: 15, bold: true, color: INK });
      T(s, it[i][2], 7.35, y + 0.34, W - MX - 7.35, 0.55, { fontSize: 12, color: GRAY });
    }
    band(s, "핵심 목표", ["방문의료 이용자 ’27. 91명", "보건소 서비스 이용률 25%", "예산 의왕살핌 514,136천원"]);
    s.addNotes("방문의료 이용자를 매년 30%씩 늘려 2030년 199명을 목표로 합니다. 기관을 더 늘리기보다 대상자를 찾아 이용을 늘리는 데 집중합니다.");
  }
  // 10 ───────── 02 신규서비스
  {
    const s = pres.addSlide();
    head(s, "02", "서비스 확충 및 품질향상", "공적서비스가 닿지 않는 돌봄 공백을\n새로운 서비스로 메우겠습니다.");
    const nw = [["LuCar", "이동지원", "승강기 없는 주택 거동불편자\n병원 진료 이동", "30명", "20,000천원"], ["LuPackage", "회복살핌 꾸러미", "치매·퇴원환자\n건강관리용품 지원", "100명", "10,000천원"], ["LuBot", "AI 돌봄로봇", "독거·안전확인 필요 대상\nAI 안부 확인", "100명", "80,000천원"], ["LuAccessibility", "장애인 건강주치의", "65세 미만 중증장애인\n참여 의원 발굴", "1개소+", "보건소 재활 연계"]];
    const cw = (CW - 0.6) / 4;
    for (let i = 0; i < 4; i++) {
      const x = MX + i * (cw + 0.2);
      R(s, x, 2.0, cw, 4.45, WHITE, { round: true, r: 0.1, line: PALE });
      badge(s, "신규", x + 0.25, 2.2);
      await ic(s, nw[i][0], x + cw / 2 - 0.55, 2.65, 1.1);
      T(s, nw[i][1], x, 3.9, cw, 0.45, { fontSize: 19, bold: true, color: INK, align: "center" });
      T(s, nw[i][2], x + 0.15, 4.4, cw - 0.3, 0.7, { fontSize: 12.5, color: GRAY, align: "center" });
      s.addShape(pres.ShapeType.line, { x: x + 0.3, y: 5.25, w: cw - 0.6, h: 0, line: { color: PALE, width: 1 }, objectName: nm("l") });
      T(s, nw[i][3], x, 5.35, cw, 0.6, { fontSize: 26, bold: true, color: BLUE, align: "center", valign: "middle" });
      T(s, nw[i][4], x, 5.95, cw, 0.35, { fontSize: 12, color: GRAY, align: "center" });
    }
    band(s, "운영 원칙", ["국가돌봄서비스 우선 연계", "공백만 지역특화서비스로 보완", "세부 내용 확정 후 단계적 개시"]);
    s.addNotes("이동지원, 회복살핌 꾸러미, AI 돌봄로봇 3종을 새로 도입하고, 장애인 건강주치의 참여 의원을 1개소 이상 발굴합니다.");
  }
  // 11 ───────── 03 운영기반 허브 (레퍼런스 9쪽형)
  {
    const s = pres.addSlide();
    head(s, "03", "운영기반 및 민관협력", "보건·복지·민간의 돌봄 역량을\n통합지원회의로 하나로 모으겠습니다.");
    const cx = W / 2, cy = 4.05, d = 2.5;
    const left = [["LuLandmark", "건보공단 의왕지사", "통합판정 · 빅데이터"], ["LuAccessibility", "연금공단 군포의왕지사", "장애인 종합판정"], ["LuHeartPulse", "보건소", "방문건강 · 치매 · 정신건강"]];
    const right = [["LuBuilding2", "협업부서", "노인장애인과 · 건축과 등"], ["LuHeartHandshake", "서비스 제공기관", "의왕살핌 협약 18개소"], ["LuStethoscope", "의료기관", "방문진료 12 · 퇴원 협약병원"]];
    const nw = 3.0, nh = 0.86, ys = [2.2, 3.62, 5.04];
    for (const [arr, x] of [[left, MX + 0.2], [right, W - MX - 0.2 - nw]]) {
      for (let i = 0; i < 3; i++) {
        const y = ys[i] + nh / 2, xe = x === MX + 0.2 ? x + nw : x;
        const x1 = Math.min(xe, cx), x2 = Math.max(xe, cx), y1 = Math.min(y, cy), y2 = Math.max(y, cy);
        const flip = (xe < cx) !== (y < cy);
        s.addShape(pres.ShapeType.line, { x: x1, y: y1, w: x2 - x1, h: Math.max(y2 - y1, 0.001), flipV: flip, line: { color: SKY, width: 1.25, dashType: "dash" }, objectName: nm("l") });
      }
    }
    for (const [arr, x] of [[left, MX + 0.2], [right, W - MX - 0.2 - nw]]) {
      for (let i = 0; i < 3; i++) {
        const y = ys[i];
        R(s, x, y, nw, nh, WHITE, { round: true, r: 0.43, line: SKY, lw: 1.25 });
        await ic(s, arr[i][0], x + 0.12, y + 0.11, 0.64);
        T(s, arr[i][1], x + 0.86, y + 0.1, nw - 0.95, 0.36, { fontSize: 13.5, bold: true, color: INK, valign: "middle" });
        T(s, arr[i][2], x + 0.86, y + 0.45, nw - 0.95, 0.3, { fontSize: 11, color: GRAY, valign: "middle" });
      }
    }
    C(s, cx - d / 2 - 0.15, cy - d / 2 - 0.15, d + 0.3, PALE);
    C(s, cx - d / 2, cy - d / 2, d, NAVY);
    s.addImage({ data: await icon("LuUsers", WHITE), x: cx - 0.25, y: cy - 0.95, w: 0.5, h: 0.5, objectName: nm("i") });
    T(s, "통합지원회의", cx - 1.2, cy - 0.4, 2.4, 0.5, { fontSize: 21, bold: true, color: WHITE, align: "center" });
    T(s, "월 2회 이상 · 계획 심의", cx - 1.2, cy + 0.12, 2.4, 0.32, { fontSize: 12, color: "C9D7FF", align: "center" });
    T(s, "운영: 복지정책과 통합돌봄팀", cx - 1.2, cy + 0.45, 2.4, 0.3, { fontSize: 11, color: "C9D7FF", align: "center" });
    R(s, cx - 2.4, 5.78, 4.8, 0.62, BLUE, { round: true, r: 0.31 });
    s.addImage({ data: await icon("LuMapPin", WHITE), x: cx - 2.15, y: 5.92, w: 0.34, h: 0.34, objectName: nm("i") });
    T(s, "6개 동 전담인력 · 발굴부터 모니터링까지", cx - 1.75, 5.78, 4.2, 0.62, { fontSize: 14, bold: true, color: WHITE, valign: "middle" });
    R(s, 0, 6.82, W, 0.68, BAND);
    T(s, "핵심사업", MX, 6.82, 1.5, 0.68, { fontSize: 14, bold: true, color: WHITE, valign: "middle" });
    const core = [["신규", "서비스 의뢰·회신 표준체계 구축"], ["확대", "전문기관 협업 정례화"], ["개선", "통합지원회의 운영 내실화"]];
    const bw = (CW - 1.6) / 3;
    core.forEach((c, i) => { const x = MX + 1.6 + i * bw; badge(s, c[0], x, 7.02); T(s, c[1], x + 0.7, 6.82, bw - 0.75, 0.68, { fontSize: 13, color: "D6E0FF", valign: "middle" }); });
    s.addNotes("통합지원회의를 중심으로 공단, 보건소, 협업부서, 제공기관, 의원, 협약병원이 함께 계획을 심의합니다. 6개 동 전담인력이 발굴부터 모니터링까지 맡습니다.");
  }
  // 12 ───────── 04 교육·홍보
  {
    const s = pres.addSlide();
    head(s, "04", "교육 및 홍보", "현장의 역량을 키우고,\n주민에게 먼저 다가가겠습니다.");
    tab(s, "인력 양성", MX, 1.95, 5.75); tab(s, "주민 홍보·소통", 6.98, 1.95, W - MX - 6.98);
    const L1 = [["LuGraduationCap", "10명", "전담인력 전원 중앙·전문기관 교육 이수"], ["LuUsers", "연 5회", "「의왕돌봄마당」 4회 + 시군 맞춤형 교육 1회"]];
    const L2 = [["LuMegaphone", "연 6회", "찾아가는 통합돌봄 설명회(동별 1회 이상)"], ["LuFileText", "1,000건", "홍보자료 배포(경로당·약국·의원 등)"], ["LuSmile", "300명", "이용자 만족도 조사 → 운영 개선 반영"]];
    for (let i = 0; i < 2; i++) { const y = 2.55 + i * 1.75; R(s, MX, y, 5.75, 1.55, WHITE, { round: true, r: 0.1, line: PALE }); await ic(s, L1[i][0], MX + 0.3, y + 0.35, 0.85); T(s, L1[i][1], MX + 1.4, y + 0.2, 4.2, 0.65, { fontSize: 30, bold: true, color: BLUE, valign: "middle" }); T(s, L1[i][2], MX + 1.4, y + 0.85, 4.2, 0.5, { fontSize: 13, color: GRAY }); }
    for (let i = 0; i < 3; i++) { const y = 2.55 + i * 1.17; await ic(s, L2[i][0], 6.98, y + 0.12, 0.75); T(s, L2[i][1], 7.95, y + 0.05, 2.0, 0.5, { fontSize: 24, bold: true, color: NAVY, valign: "middle" }); T(s, L2[i][2], 7.95, y + 0.55, W - MX - 7.95, 0.4, { fontSize: 12.5, color: GRAY }); }
    band(s, "중점", ["신규 동 인력 상반기 직무교육", "제공기관 공동교육으로 절차 통일", "퇴원 예정자에 병원 안내문 배부"]);
    s.addNotes("전담인력 10명 전원 교육 이수, 자체교육 연 5회, 동별 설명회 연 6회를 운영합니다.");
  }
  // 13 ───────── 대상자·성과지표
  {
    const s = pres.addSlide();
    head(s, "목표", "대상자 및 성과지표", "2027년 541명을 지원하고,\n발굴률 20%를 달성하겠습니다.");
    R(s, MX, 1.95, 4.3, 4.6, NAVY, { round: true, r: 0.12 });
    T(s, "통합지원 대상자", MX + 0.35, 2.2, 3.6, 0.4, { fontSize: 15, color: "C9D7FF", bold: true });
    T(s, "541명", MX + 0.35, 2.6, 3.6, 1.0, { fontSize: 54, bold: true, color: WHITE, valign: "middle" });
    T(s, "’26년 276명 → 약 2배", MX + 0.35, 3.65, 3.6, 0.4, { fontSize: 14, color: "C9D7FF" });
    const a = 3.6 * 421 / 541;
    R(s, MX + 0.35, 4.3, a, 0.45, "7FA2F7"); R(s, MX + 0.35 + a + 0.04, 4.3, 3.6 - a - 0.04, 0.45, ORG);
    T(s, "우선발굴필요군 421", MX + 0.35, 4.3, a, 0.45, { fontSize: 12, bold: true, color: WHITE, align: "center", valign: "middle" });
    T(s, "기타 120", MX + 0.39 + a, 4.3, 3.56 - a, 0.45, { fontSize: 11, bold: true, color: WHITE, align: "center", valign: "middle" });
    T(s, "우선발굴필요군 6,506명 중 6.5%", MX + 0.35, 4.95, 3.6, 0.4, { fontSize: 13, color: "C9D7FF" });
    T(s, "재가급여 145 · 고령장애인 107 · 퇴원환자 60 · 등급외 50 · 기각 40 · 중증장애인 17 · 시설퇴소 2", MX + 0.35, 5.4, 3.6, 0.9, { fontSize: 12, color: "C9D7FF" });
    tab(s, "중앙 성과지표", 5.2, 1.95, W - MX - 5.2);
    const k = [["20%", "직권·자체발굴 발굴률"], ["6%", "지원계획 수립률"], ["80%", "모니터링 수행률"], ["25%", "보건소 서비스 이용률"], ["70%", "대상자 중 우선발굴필요군"], ["100%", "전담인력 배치율"], ["90%", "보건소 회의 참석률"], ["421명", "우선발굴필요군 중 대상자"]];
    const kw = (W - MX - 5.2 - 0.45) / 4;
    k.forEach((v, i) => {
      const x = 5.2 + (i % 4) * (kw + 0.15), y = 2.5 + Math.floor(i / 4) * 2.05, hi = i === 0;
      R(s, x, y, kw, 1.85, hi ? BLUE : WHITE, { round: true, r: 0.1, line: hi ? null : "C9D5F2", lw: 1.25 });
      const num = v[0].startsWith("월 ") ? [{ text: "월 ", options: { fontSize: 20 } }, { text: v[0].slice(2), options: { fontSize: 38 } }] : [{ text: v[0], options: { fontSize: v[0].length >= 4 ? 33 : 38 } }];
      T(s, num, x, y + 0.18, kw, 0.95, { bold: true, fontFace: TITLE_FONT, color: hi ? WHITE : NAVY, align: "center", valign: "middle" });
      s.addShape(pres.ShapeType.line, { x: x + kw / 2 - 0.25, y: y + 1.17, w: 0.5, h: 0, line: { color: hi ? "C9D7FF" : "C9D5F2", width: 1.25 }, objectName: nm("l") });
      T(s, v[1], x + 0.08, y + 1.25, kw - 0.16, 0.55, { fontSize: 13.5, bold: true, color: hi ? WHITE : INK, align: "center", valign: "top" });
    });
    s.addNotes("목표 대상자는 541명입니다. 합동평가 핵심인 직권·자체발굴 발굴률 20%를 비롯해 중앙 지표 8개를 관리합니다.");
  }
  // 14 ───────── 예산
  {
    const s = pres.addSlide();
    head(s, "예산", "2027년 재원 조달 및 운용", "총 706,636천원을\n국가서비스가 채우지 못한 공백에 쓰겠습니다.");
    tab(s, "재원 구성", MX, 1.95, 4.6);
    s.addChart(pres.charts.DOUGHNUT, [{ name: "재원", labels: ["국비 50%", "시비 45%", "도비 5%"], values: [353318, 317986, 35332] }],
      { x: MX + 0.6, y: 2.42, w: 3.4, h: 2.85, holeSize: 64, chartColors: [H.dk2, "6F8FE8", H.accent2], showPercent: false, showValue: false, showLegend: false, dataBorder: { pt: 1.5, color: "FFFFFF" }, objectName: nm("ch") });
    T(s, [{ text: "706,636", options: { fontSize: 18, fontFace: TITLE_FONT } }, { text: "\n천원", options: { fontSize: 13 } }], MX + 1.55, 3.35, 1.5, 1.0, { bold: true, color: NAVY, align: "center", valign: "middle" });
    const src = [["국비", "50%", "353,318", H.dk2], ["시비", "45%", "317,986", "6F8FE8"], ["도비", "5%", "35,332", H.accent2]];
    src.forEach((r, i) => {
      const y = 5.33 + i * 0.44;
      C(s, MX + 0.1, y + 0.12, 0.2, r[3]);
      T(s, [{ text: r[0] + " ", options: { bold: true, color: INK } }, { text: r[1], options: { bold: true, color: GRAY } }], MX + 0.42, y, 1.6, 0.44, { fontSize: 15, valign: "middle" });
      T(s, [{ text: r[2], options: { fontSize: 17, bold: true, fontFace: TITLE_FONT, color: NAVY } }, { text: " 천원", options: { fontSize: 12, color: GRAY } }], MX + 1.9, y, 2.7, 0.44, { align: "right", valign: "middle" });
    });
    tab(s, "사업별(천원)", 5.6, 1.95, W - MX - 5.6);
    const items = [["LuHouse", "의왕살핌(일상·건강·내집)", 514136], ["LuBot", "신규 3종(AI로봇·이동·회복)", 110000], ["LuHeartHandshake", "신규특화사업(보조)", 60000], ["LuFileText", "운영비(회의·홍보·교육)", 17000], ["LuHospital", "재택의료센터 운영지원", 5000], ["LuWallet", "퇴원환자 연계수당", 500]];
    const bx = 9.5, ax = W - MX - 1.5, maxw = ax - bx - 0.15;
    for (let i = 0; i < items.length; i++) {
      const y = 2.48 + i * 0.68;
      await ic(s, items[i][0], 5.6, y, 0.52);
      T(s, items[i][1], 6.25, y, 3.2, 0.52, { fontSize: 14, bold: i < 2, color: INK, valign: "middle" });
      const bw = Math.max(0.06, maxw * items[i][2] / 514136);
      R(s, bx, y + 0.1, bw, 0.32, i === 0 ? BLUE : "6F8FE8");
      T(s, items[i][2].toLocaleString("en-US"), ax, y, 1.5, 0.52, { fontSize: 17, bold: true, fontFace: TITLE_FONT, color: NAVY, align: "right", valign: "middle" });
    }
    band(s, "집행 원칙", ["국가서비스 우선, 공백 영역에만 투입", "기존 제도 단가 준용·소득별 차등", "현금성 급여·본인부담 전액지원 배제"]);
    s.addNotes("총 7억 664만원, 국비 50, 도비 5, 시비 45퍼센트입니다. 사업비의 대부분이 의왕살핌과 신규서비스입니다.");
  }
  // 15 ───────── 닫는 슬라이드
  {
    const s = pres.addSlide(); bg(s, "bg_mid.png");
    await lett(s, "L_closing", 0, 1.05, 6.6, "2027년도 통합지원 지역계획(안)을 심의해 주십시오", true);
    const it = [["LuTarget", "목표 대상자", "통합지원 대상자 541명"], ["LuChartBar", "성과지표", "발굴률 20% 등 중앙 8개"], ["LuListChecks", "추진과제", "4대 과업 · 12개 핵심사업"], ["LuWallet", "예산", "706,636천원"]];
    const cw = (CW - 0.6) / 4;
    for (let i = 0; i < 4; i++) { const x = MX + i * (cw + 0.2); R(s, x, 3.35, cw, 2.0, WHITE, { round: true, r: 0.12, tr: 10 }); await ic(s, it[i][0], x + cw / 2 - 0.4, 3.55, 0.8); T(s, it[i][1], x, 4.45, cw, 0.4, { fontSize: 17, bold: true, color: INK, align: "center" }); T(s, it[i][2], x + 0.15, 4.85, cw - 0.3, 0.4, { fontSize: 12, color: GRAY, align: "center" }); }
    const fl = ["협의체 심의 11. 26.", "의견 반영", "지역계획 확정·제출", "2027. 1. 시행"];
    const fw = (CW - 3 * 0.4) / 4;
    s.addImage({ data: await icon("LuLandmark", NAVY), x: MX, y: 6.68, w: 0.3, h: 0.3, objectName: nm("i") });
    T(s, [{ text: "통합지원협의체 ", options: { bold: true, color: NAVY } }, { text: "의료·복지·공단·학계·시의회 위원 · 연 2회 · 지역계획 심의와 기관 연계·협력 자문", options: { color: GRAY } }], MX + 0.4, 6.6, CW - 0.4, 0.45, { fontSize: 13, valign: "middle" });
    fl.forEach((f, i) => { const x = MX + i * (fw + 0.4); R(s, x, 5.85, fw, 0.6, i === 0 ? BLUE : NAVY, { round: true, r: 0.3 }); T(s, f, x, 5.85, fw, 0.6, { fontSize: 14, bold: true, color: WHITE, align: "center", valign: "middle" }); if (i < 3) arrowR(s, x + fw + 0.07, 6.0, 0.26, 0.3, NAVY); });
    s.addNotes("이상 네 가지에 대해 심의와 의견을 부탁드립니다.");
  }

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT);
})();
