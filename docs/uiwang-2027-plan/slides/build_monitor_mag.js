// 회의장 모니터 안내 2장(마젠타 배경 2종): NODE_PATH=<node_modules> node build_monitor_mag.js <out.pptx>
const path = require("path");
const pptxgen = require("pptxgenjs");
const React = require("react"), RDS = require("react-dom/server"), sharp = require("sharp"), L = require("react-icons/lu");
const FIG = path.join(__dirname, "..", "figures");
const OUT = process.argv[2] || "monitor_mag.pptx";
const F = "Pretendard", FT = "Pretendard ExtraBold";
const NAVY = "13277A", INK = "111B3D", GRAY = "5B6785", MAG = "C8157F", MAG_D = "8E0F5B", PINK = "FCE4F1", BAND = "0E1F5B";
const W = 13.333, MX = 0.6, CW = W - 2 * MX;
async function icon(n, c) { const svg = RDS.renderToStaticMarkup(React.createElement(L[n], { color: "#" + c, size: "256" })); return "image/png;base64," + (await sharp(Buffer.from(svg)).png().toBuffer()).toString("base64"); }

(async () => {
  const p = new pptxgen(); p.layout = "LAYOUT_WIDE"; p.title = "제2차 의왕시 통합지원협의체 회의";
  const mk = (s) => (t, x, y, w, h, o = {}) => s.addText(t, { x, y, w, h, isTextBox: true, margin: 0, fontFace: F, fontSize: 14, color: INK, valign: "top", ...o });
  const pill = (s, T, label, x, y, w, h, fill, fs = 15) => { s.addShape(p.ShapeType.roundRect, { x, y, w, h, fill: { color: fill }, line: { type: "none" }, rectRadius: h / 2 }); T(label, x, y, w, h, { fontSize: fs, bold: true, color: "FFFFFF", align: "center", valign: "middle" }); };

  // ── 1장: 띠 없는 배경 — 가운데 정렬 대기 화면
  {
    const s = p.addSlide(), T = mk(s);
    s.addImage({ path: path.join(FIG, "monitor", "bg_mag.png"), x: 0, y: 0, w: W, h: 7.5, altText: "" });
    pill(s, T, "통합지원협의체", (W - 2.2) / 2, 1.25, 2.2, 0.46, MAG, 15);
    T("2026년 제2차 의왕시", 0, 1.95, W, 0.75, { fontFace: FT, fontSize: 36, color: NAVY, align: "center", valign: "middle" });
    T([{ text: "통합지원협의체 ", options: { color: INK } }, { text: "회의", options: { color: MAG } }], 0, 2.7, W, 1.3, { fontFace: FT, fontSize: 70, align: "center", valign: "middle" });
    T("2027년도 통합지원 지역계획(안) 심의", 0, 4.0, W, 0.55, { fontSize: 22, bold: true, color: GRAY, align: "center", valign: "middle" });
    const pw = 9.4, px = (W - pw) / 2, py = 4.95;
    s.addShape(p.ShapeType.roundRect, { x: px, y: py, w: pw, h: 0.8, fill: { color: "FFFFFF", transparency: 15 }, line: { color: "F3C3DD", width: 1.25 }, rectRadius: 0.4 });
    s.addImage({ data: await icon("LuCalendarCheck", MAG), x: px + 0.35, y: py + 0.2, w: 0.4, h: 0.4 });
    T("2026. 11. 26.(목) 14:00", px + 0.85, py, 4.1, 0.8, { fontFace: FT, fontSize: 22, color: INK, valign: "middle" });
    s.addShape(p.ShapeType.line, { x: px + 5.0, y: py + 0.2, w: 0, h: 0.4, line: { color: "F3C3DD", width: 1.5 } });
    s.addImage({ data: await icon("LuMapPin", MAG), x: px + 5.3, y: py + 0.2, w: 0.4, h: 0.4 });
    T("의왕시청 중회의실", px + 5.8, py, 3.4, 0.8, { fontFace: FT, fontSize: 22, color: INK, valign: "middle" });
    T("의왕시 복지정책과 통합돌봄팀", 0, 6.55, W, 0.4, { fontSize: 14, color: GRAY, align: "center", valign: "middle" });
    s.addNotes("회의 시작 전 대기 화면");
  }

  // ── 2장: 하단 띠 배경 — 회의 진행 순서 안내
  {
    const s = p.addSlide(), T = mk(s);
    s.addImage({ path: path.join(FIG, "monitor", "bg_mag_band.png"), x: 0, y: 0, w: W, h: 7.5, altText: "" });
    s.addShape(p.ShapeType.rect, { x: 0, y: 6.82, w: W, h: 0.68, fill: { color: BAND }, line: { type: "none" } });
    T([{ text: "제2차  ", options: { bold: true, color: MAG } }, { text: "—  의왕시 통합지원협의체 회의", options: { color: GRAY } }], MX, 0.5, 10, 0.35, { fontSize: 15 });
    T([{ text: "오늘 회의는 이렇게 ", options: { color: INK } }, { text: "진행", options: { color: MAG } }, { text: "됩니다.", options: { color: INK } }], MX, 0.92, CW, 0.9, { fontFace: FT, fontSize: 38 });
    const st = [["LuMic", "개회", ""], ["LuUsers", "참석자 소개", ""], ["LuMessageSquare", "인사 말씀", "위원장"], ["LuListChecks", "성원 보고", ""], ["LuFileText", "안건", "보고 · 심의"], ["LuCircleCheck", "폐회", ""]];
    const n = st.length, bw = 1.72, gap = (CW - n * bw) / (n - 1), y = 2.25;
    s.addShape(p.ShapeType.line, { x: MX + bw / 2, y: y + 0.55, w: CW - bw, h: 0, line: { color: "E7B3D2", width: 3 } });
    for (let i = 0; i < n; i++) {
      const x = MX + i * (bw + gap), hi = i === 4, d = 1.1;
      s.addShape(p.ShapeType.ellipse, { x: x + (bw - d) / 2, y, w: d, h: d, fill: { color: hi ? MAG : "FFFFFF" }, line: { color: hi ? MAG : "F3C3DD", width: 1.5 } });
      s.addImage({ data: await icon(st[i][0], hi ? "FFFFFF" : MAG), x: x + (bw - 0.5) / 2, y: y + 0.3, w: 0.5, h: 0.5 });
      T(st[i][1], x - 0.1, y + 1.25, bw + 0.2, 0.45, { fontFace: FT, fontSize: 19, color: hi ? MAG : INK, align: "center" });
      if (st[i][2]) T(st[i][2], x - 0.1, y + 1.7, bw + 0.2, 0.35, { fontSize: 13.5, color: GRAY, align: "center" });
    }
    // 안건 카드
    const cw = (CW - 0.3) / 2, ay = 4.75;
    const ag = [["보고", "2026년 통합돌봄 추진사항", "4F6BD8", "FFFFFF", INK], ["심의", "2027년도 통합지원 지역계획(안)", "FFFFFF", MAG_D, "FFFFFF"]];
    ag.forEach((a, i) => {
      const x = MX + i * (cw + 0.3);
      s.addShape(p.ShapeType.roundRect, { x, y: ay, w: cw, h: 1.0, fill: { color: a[3] }, line: i ? { type: "none" } : { color: "F3C3DD", width: 1.25 }, rectRadius: 0.12 });
      pill(s, T, a[0], x + 0.3, ay + 0.31, 0.85, 0.38, a[2], 14);
      if (i) T(a[0], x + 0.3, ay + 0.31, 0.85, 0.38, { fontSize: 14, bold: true, color: MAG, align: "center", valign: "middle" });
      T(a[1], x + 1.35, ay, cw - 1.5, 1.0, { fontSize: 19, bold: true, color: a[4], valign: "middle" });
    });
    T("2026. 11. 26.(목) 14:00  |  의왕시청 중회의실", MX, 6.82, 8, 0.68, { fontSize: 15, bold: true, color: "FFFFFF", valign: "middle" });
    T("의왕시 복지정책과 통합돌봄팀", W - MX - 4, 6.82, 4, 0.68, { fontSize: 13, color: "F3C3DD", align: "right", valign: "middle" });
    s.addNotes("회의 진행 순서 안내 화면");
  }
  await p.writeFile({ fileName: OUT }); console.log("wrote", OUT);
})();
